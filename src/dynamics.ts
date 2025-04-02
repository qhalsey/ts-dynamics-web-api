// src/dynamics.ts

import axios from "axios";

export async function fetchData(
  entity: string,
  accessToken: string,
  dynamicsUrl: string
): Promise<any[]> {
  // For example: dynamicsUrl might be https://org0b26dba9.api.crm.dynamics.com/api/data/v9.1
  // Then we build a request like GET {dynamicsUrl}/{entity}?$select=...

  const url = `${dynamicsUrl}/${entity}`; // e.g., /accounts or /contacts, etc.

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        // If needed, you can set an OData version or other headers here
        // 'OData-Version': '4.0',
        // 'Accept': 'application/json',
      },
      // If you have query params or $select fields, you can do so here or in the URL
    });

    // Usually D365 data is in response.data.value
    // but let's log to confirm the shape
    console.log(`Fetched ${entity} data:`, response.data);

    // Return the array of records
    return response.data.value || [];
  } catch (error: any) {
    console.error(
      `Error fetching ${entity}:`,
      error.response?.data || error.message
    );
    throw error;
  }
}
