// src/index.ts

import * as dotenv from "dotenv";
dotenv.config();

import axios from "axios";
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

    // const listOfEntities: string[] = ["account", "contact"];

    const listOfEntites: string[] = [
      "account",
      "bam_batch",
      "bam_bpf_556396196ab5470b836c103f3dfd232a",
      "bam_city",
      "bam_consignmentbatch",
      "bam_consignmentdetail",
      "bam_field",
      "bam_fieldproductstatus",
      "bam_harvestrecord",
      "bam_inventorytagcount",
      "bam_inventorytransactionlog",
      "bam_lot",
      "bam_lotbagtag",
      "bam_lottesting",
      "bam_mlra",
      "bam_mlra_product",
      "bam_nationalplantlist",
      "bam_orderbagtag",
      "bam_pickingbatch",
      "bam_plsrate",
      "bam_processingrecord",
      "bam_processtrigger",
      "bam_productfeedback",
      "bam_productheading",
      "bam_projecttype",
      "bam_purchaseorder",
      "bam_purchaseorderdetail",
      "bam_quoteconfirmation",
      "bam_rebate",
      "bam_state",
      "bam_subfield",
      "bam_term",
      "bam_vendor",
      "bam_vendorproduct",
      "bam_yield",
      "contact",
      "contactleads",
      "customeraddress",
      "discount",
      "discounttype",
      "invoice",
      "invoicedetail",
      "opportunity",
      "product",
      "productassociation",
      "productpricelevel",
      "productsalesliterature",
      "productsubstitute",
    ];

    // Loop through each entity and process it
    for (const entity of listOfEntites) {
      console.log(`Processing entity: ${entity}`);
      await processEntityAll(accessToken, entity);
      console.log(`Finished processing entity: ${entity}`);
    }
  } catch (err) {
    console.error("Unhandled error:", err);
  }
})();
