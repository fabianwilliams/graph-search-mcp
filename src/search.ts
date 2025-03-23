// Author: Fabian G. Williams
// Purpose: Defines a focused MCP tool for searching DriveItems via Microsoft Graph

import { z } from 'zod';
import { getToken } from './auth.js';

export const inputSchema = z.object({
  query: z.string().describe("The search term (e.g. 'report.xlsx')"),
  user: z
    .string()
    .optional()
    .describe('User ID to search their OneDrive instead of the current user'),
});

export const searchTool = {
  name: 'searchDriveItems',
  description: 'Search for files in Microsoft 365 using the Graph API',
  inputSchema,
  handler: async ({ query, user }: z.infer<typeof inputSchema>) => {
    const token = await getToken();

    const url = user
      ? `https://graph.microsoft.com/v1.0/users/${user}/drive/root/search(q='${query}')`
      : `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${query}')`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: 'eventual',
      },
    });

    const data = (await response.json()) as {
      value?: { name: string; webUrl: string }[];
    };
    const files = data.value ?? [];

    // Format as MCP-compatible content (type: "resource")
    let content = files.map((file) => ({
      type: 'resource' as const,
      resource: {
        text: file.name,
        uri: file.webUrl,
        mimeType: 'text/html',
      },
    }));

    // Add fallback message if no files were found
    if (content.length === 0) {
      content = [
        {
          type: 'resource' as const,
          resource: {
            text: '📁 No files found.',
            uri: 'https://graph.microsoft.com/v1.0/me/drive/root', // Placeholder URI
            mimeType: 'text/plain',
          },
        },
      ];
    }

    return {
      content,
      _meta: {},
      isError: false,
    };
  },
};
