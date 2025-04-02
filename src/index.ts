// src/index.ts

import * as dotenv from "dotenv";
dotenv.config();

import axios from "axios";
import { processEntityColumns } from "./entityColumns";
import { processEntityAll } from "./processEntity";

async function getAccessToken() {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
      throw new Error(
        "Missing TENANT_ID, CLIENT_ID or CLIENT_SECRET in environment variables."
      );
    }

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/token`;

    const params = new URLSearchParams();
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("grant_type", "client_credentials");
    params.append("resource", "https://org0b26dba9.crm.dynamics.com");

    const response = await axios.post(tokenUrl, params.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });

    return response.data.access_token;
  } catch (error: any) {
    console.error(
      "Error getting access token:",
      error.response?.data || error.message
    );
    throw error;
  }
}

(async () => {
  try {
    console.log("Fetching access token...");
    const accessToken = await getAccessToken();
    console.log("Access token acquired.");

    // The base URL for your D365 environment
    const orgUrl = "https://org0b26dba9.crm.dynamics.com/api/data/v9.2";

    // For now, let's just do "account"
    await processEntityColumns(accessToken, "account", orgUrl);

    console.log("Done processing account columns!");

    await processEntityAll(accessToken, "account");

    // Later, you could do "contacts", etc.
    // await processEntityColumns(accessToken, "contact", orgUrl);
    // console.log("Done processing contact columns!");
  } catch (err) {
    console.error("Unhandled error:", err);
  }
})();
