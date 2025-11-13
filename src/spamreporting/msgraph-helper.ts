// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides functionality to get Microsoft Graph data.

/* global console fetch */

/**
 *  Calls a Microsoft Graph API and returns the response.
 *
 * @param accessToken The access token to use for the request.
 * @param path Path component of the URI, e.g., "/me". Should start with "/".
 * @param queryParams Query parameters, e.g., "?$select=name,id". Should start with "?".
 * @returns
 */
export async function makeGraphRequest(accessToken: string, path: string, queryParams: string): Promise<any> {
  if (!path) throw new Error("path is required.");
  if (!path.startsWith("/")) throw new Error("path must start with '/'.");
  if (queryParams && !queryParams.startsWith("?")) throw new Error("queryParams must start with '?'.");

  const response = await fetch(`https://graph.microsoft.com/beta${path}${queryParams}`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    const data = await response.json();
    console.log(data);
    return data;
  } else {
    throw new Error(response.statusText);
  }
}

/**
 * Sends a POST request to Microsoft Graph API with a JSON body.
 *
 * @param accessToken The access token to use for the request.
 * @param path Path component of the URI, e.g., "/security/threatSubmission/emailThreats". Should start with "/".
 * @param body The request body object to be JSON stringified.
 * @returns The response data from the API.
 */
export async function makeGraphPostRequest(accessToken: string, path: string, body: any): Promise<any> {
  console.log("[makeGraphPostRequest] Starting POST request");
  console.log("[makeGraphPostRequest] Path:", path);
  console.log("[makeGraphPostRequest] Body:", JSON.stringify(body, null, 2));

  if (!path) throw new Error("path is required.");
  if (!path.startsWith("/")) throw new Error("path must start with '/'.");

  const url = `https://graph.microsoft.com/beta${path}`;
  console.log("[makeGraphPostRequest] Full URL:", url);

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    console.log("[makeGraphPostRequest] Response status:", response.status, response.statusText);
    console.log("[makeGraphPostRequest] Response ok:", response.ok);

    if (response.ok) {
      const data = await response.json();
      console.log("[makeGraphPostRequest] ✓ Success! Response data:", data);
      return data;
    } else {
      const errorText = await response.text();
      console.error("[makeGraphPostRequest] ✗ Request failed!");
      console.error("[makeGraphPostRequest] Status:", response.status, response.statusText);
      console.error("[makeGraphPostRequest] Error response body:", errorText);

      let errorMessage = `${response.status} ${response.statusText}`;
      try {
        const errorData = JSON.parse(errorText);
        if (errorData.error) {
          errorMessage += `: ${errorData.error.message || JSON.stringify(errorData.error)}`;
        }
      } catch (e) {
        errorMessage += `: ${errorText}`;
      }

      throw new Error(errorMessage);
    }
  } catch (error) {
    console.error("[makeGraphPostRequest] Exception occurred:", error);
    throw error;
  }
}
