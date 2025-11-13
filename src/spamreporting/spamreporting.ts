/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

import { AccountManager } from "./authConfig";
import { makeGraphRequest, makeGraphPostRequest } from "./msgraph-helper";

const accountManager = new AccountManager();

// Initialize when Office is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {    
    // Initialize MSAL.
    accountManager.initialize();
  }
});


/**
 * Gets the current user's ID from the Graph API.
 * @param accessToken The access token to use for the Graph API request
 * @returns The user's ID (GUID)
 */
async function getCurrentUserId(accessToken: string): Promise<string> {
  console.log("[getCurrentUserId] Fetching current user ID from Graph API");

  try {
    const userProfile = await makeGraphRequest(accessToken, "/me", "?$select=id");
    console.log("[getCurrentUserId] User profile retrieved:", userProfile);

    if (userProfile && userProfile.id) {
      console.log("[getCurrentUserId] User ID:", userProfile.id);
      return userProfile.id;
    } else {
      console.error("[getCurrentUserId] No user ID found in response");
      throw new Error("Unable to retrieve user ID from Graph API");
    }
  } catch (error) {
    console.error("[getCurrentUserId] Failed to get user ID:", error);
    throw error;
  }
}

/**
 * Gets the currently selected email's message URL in Graph API format.
 * @param accessToken The access token to use for the Graph API request
 * @returns The Graph API URL for the currently selected email message.
 */
async function getCurrentMessageUrl(accessToken: string): Promise<string> {
  console.log("[getCurrentMessageUrl] Starting to get current message URL");

  return new Promise(async (resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.error("[getCurrentMessageUrl] No email item is currently selected");
      reject(new Error("No email item is currently selected."));
      return;
    }

    console.log("[getCurrentMessageUrl] Email item found");
    console.log("[getCurrentMessageUrl] Item type:", typeof item);
    console.log("[getCurrentMessageUrl] Available properties:", Object.keys(item));

    // Check if itemId is directly available (in read mode)
    if (item.itemId) {
      console.log("[getCurrentMessageUrl] Using item.itemId property");
      const itemId = item.itemId;

      try {
        // Get the current user's ID from Graph API
        const userId = await getCurrentUserId(accessToken);
        console.log("[getCurrentMessageUrl] Retrieved user ID:", userId);

        // Convert EWS ID to REST ID if needed
        const restId = Office.context.mailbox.convertToRestId(
          itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );

        // Use Graph API URL format matching the threat submission documentation
        // Use beta endpoint and users/{userId}/messages format as per documentation
        const messageUrl = `https://graph.microsoft.com/beta/users/${userId}/messages/${restId}`;

        console.log("[getCurrentMessageUrl] Successfully got message URL");
        console.log("[getCurrentMessageUrl] EWS Item ID:", itemId);
        console.log("[getCurrentMessageUrl] REST Item ID:", restId);
        console.log("[getCurrentMessageUrl] Graph API message URL:", messageUrl);

        resolve(messageUrl);
      } catch (error) {
        console.error("[getCurrentMessageUrl] Failed to convert to REST ID or get user ID:", error);
        reject(error);
      }
    } else {
      console.error("[getCurrentMessageUrl] Item ID is not available");
      console.error("[getCurrentMessageUrl] item.itemId:", item.itemId);
      reject(new Error("Item ID is not available - cannot get message URL"));
    }
  });
}

/**
 * Gets the recipient email address from the current mailbox user.
 * @returns The email address of the current user.
 */
async function getRecipientEmailAddress(): Promise<string> {
  console.log("[getRecipientEmailAddress] Getting recipient email address");

  return new Promise((resolve, reject) => {
    const userProfile = Office.context.mailbox.userProfile;
    if (userProfile && userProfile.emailAddress) {
      console.log("[getRecipientEmailAddress] Email address:", userProfile.emailAddress);
      resolve(userProfile.emailAddress);
    } else {
      console.error("[getRecipientEmailAddress] Unable to get user email address");
      reject(new Error("Unable to get user email address"));
    }
  });
}

/**
 * Submits the current email as a spam/phishing threat to Microsoft.
 * @param category The category of the threat: "spam", "phishing", or "notJunk"
 */
async function submitEmailThreat(category: "spam" | "phishing" | "notJunk") {
  console.log("[submitEmailThreat] Starting email threat submission");
  console.log("[submitEmailThreat] Category:", category);

  try {
    // Get the access token with required scopes
    console.log("[submitEmailThreat] Requesting access token with scopes: ThreatSubmission.ReadWrite, Mail.Read");
    const accessToken = await accountManager.ssoGetAccessToken([
      "ThreatSubmission.ReadWrite",
      "Mail.Read",
    ]);
    console.log("[submitEmailThreat] Access token obtained successfully");
    console.log("[submitEmailThreat] Token (first 20 chars):", accessToken.substring(0, 20) + "...");

    // Get the recipient email address (required field)
    console.log("[submitEmailThreat] Getting recipient email address...");
    const recipientEmail = await getRecipientEmailAddress();
    console.log("[submitEmailThreat] Recipient email:", recipientEmail);

    // Get the current message URL
    console.log("[submitEmailThreat] Getting current message URL...");
    const messageUrl = await getCurrentMessageUrl(accessToken);
    console.log("[submitEmailThreat] Message URL obtained:", messageUrl);

    // Prepare the request body with required fields
    const body: any = {
      "@odata.type": "#microsoft.graph.security.emailUrlThreatSubmission",
      category: category,
      recipientEmailAddress: recipientEmail,
      messageUrl: messageUrl,
    };

    console.log("[submitEmailThreat] Request body:", JSON.stringify(body, null, 2));

    // Submit the threat
    console.log("[submitEmailThreat] Sending POST request to /security/threatSubmission/emailThreats");
    const response = await makeGraphPostRequest(accessToken, "/security/threatSubmission/emailThreats", body);

    console.log("[submitEmailThreat] ✓ Email threat submitted successfully!");
    console.log("[submitEmailThreat] Response:", JSON.stringify(response, null, 2));
    return response;
  } catch (error) {
    console.error("[submitEmailThreat] ✗ Failed to submit email threat");
    console.error("[submitEmailThreat] Error details:", error);
    throw error;
  }
}

/**
 * Handler for the spam report button. This is called by the Office add-in framework.
 * @param event The event object from the Office add-in framework
 */
async function onSpamReport(event: Office.AddinCommands.Event & { options?: any }) {
  console.log("=".repeat(80));
  console.log("[onSpamReport] *** SPAM REPORT BUTTON CLICKED ***");
  console.log("[onSpamReport] Event received:", event);
  console.log("=".repeat(80));

  try {
    // Ensure AccountManager is initialized
    console.log("[onSpamReport] Initializing AccountManager...");
    await accountManager.initialize();
    console.log("[onSpamReport] AccountManager initialized successfully");

    // Get the reporting options selected by the user from the preprocessing dialog
    const reportingOptions = event.options;
    console.log("[onSpamReport] Reporting options from dialog:", reportingOptions);

    // Determine the category based on user selection
    let category: "phishing" | "spam" | "notJunk" = "phishing"; // Default to phishing

    if (reportingOptions) {
      // Map the selected option to a category based on boolean object with numeric keys
      console.log("[onSpamReport] Available properties in reportingOptions:", Object.keys(reportingOptions));
      
      
      console.log("[onSpamReport] Reporting options object:", reportingOptions);      

      // Check the boolean values directly from the object with numeric keys
      if (reportingOptions[0] === true) {
        category = "phishing";
      } else if (reportingOptions[1] === true) {
        category = "spam";
      } else if (reportingOptions[2] === true) {
        category = "notJunk";
      }      
    }

    console.log("[onSpamReport] Final category:", category);

    // Submit the email threat
    console.log("[onSpamReport] Calling submitEmailThreat...");
    await submitEmailThreat(category);

    // Signal that the spam report was successful
    console.log("[onSpamReport] ✓ Spam report completed successfully!");
    console.log("[onSpamReport] Signaling success to Office framework");
    event.completed({ allowEvent: true });
  } catch (error) {
    console.error("[onSpamReport] ✗ Error in onSpamReport!");
    console.error("[onSpamReport] Full error:", error);

    if (error instanceof Error) {
      console.error("[onSpamReport] Error type:", error.constructor.name);
      console.error("[onSpamReport] Error message:", error.message);
      console.error("[onSpamReport] Stack trace:", error.stack);
    }

    // Signal that the spam report failed
    console.log("[onSpamReport] Signaling failure to Office framework");
    event.completed({ allowEvent: false });
  }
}

// Register the onSpamReport function so it can be called by the Office add-in framework
if (typeof Office !== "undefined") {
  Office.actions.associate("onSpamReport", onSpamReport);
}
