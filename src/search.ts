// Author: Fabian G. Williams
// Purpose: MCP Tool to search OneDrive files using Graph API with app-only fallback logic

import { z } from 'zod';
import { getToken } from './auth.js';

export const inputSchema = z.object({
  query: z.string().describe("The search term (e.g. 'report.xlsx')"),
  user: z
    .string()
    .optional()
    .describe(
      "User's UPN or ID to search their OneDrive. If omitted, defaults to the first user found.",
    ),
});

async function getDefaultUserId(token: string): Promise<string> {
  const response = await fetch(
    'https://graph.microsoft.com/v1.0/users?$top=1',
    {
      headers: { Authorization: `Bearer ${token}` },
    },
  );

  if (!response.ok)
    throw new Error(`Unable to fetch users: ${response.statusText}`);

  const json = await response.json();
  const user = json.value?.[0];
  if (!user?.id) throw new Error('No users found in tenant.');

  console.log(
    `🧠 Defaulting to user: ${user.displayName} (${user.userPrincipalName})`,
  );
  return user.id;
}

async function searchViaGraphSearchAPI(
  token: string,
  userId: string,
  query: string,
) {
  const url = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/search(q='${encodeURIComponent(query)}')`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      ConsistencyLevel: 'eventual',
    },
  });

  return response.ok ? (await response.json()).value : null;
}

async function searchViaFallback(token: string, userId: string, query: string) {
  const url = `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/children?$top=999`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) return [];

  const json = await response.json();
  return (
    json.value?.filter((file: any) =>
      file.name?.toLowerCase().includes(query.toLowerCase()),
    ) ?? []
  );
}

export const searchTool = {
  name: 'searchDriveItems',
  description:
    'Search for files in Microsoft 365 using the Graph API with fallback support',
  inputSchema,
  handler: async ({ query, user }: z.infer<typeof inputSchema>) => {
    const token = await getToken();
    const userId = user ?? (await getDefaultUserId(token));

    let files = await searchViaGraphSearchAPI(token, userId, query);
    let usedFallback = false;

    if (!files) {
      console.warn(
        `🔁 Falling back to drive/root/children listing for user: ${userId}`,
      );
      files = await searchViaFallback(token, userId, query);
      usedFallback = true;
    }

    let content = files.map((file: any) => ({
      type: 'resource' as const,
      resource: {
        text: file.name,
        uri: file.webUrl,
        mimeType: file?.file?.mimeType ?? 'application/octet-stream',
        metadata: {
          id: file.id,
          size: file.size,
          createdBy: file?.createdBy?.user?.displayName ?? 'Unknown',
          lastModifiedBy: file?.lastModifiedBy?.user?.displayName ?? 'Unknown',
          lastModified: file?.lastModifiedDateTime ?? null,
          shared: file?.shared?.scope ?? 'private',
          downloadUrl: file['@microsoft.graph.downloadUrl'] ?? undefined,
        },
      },
    }));

    if (content.length === 0) {
      content = [
        {
          type: 'resource' as const,
          resource: {
            text: '📁 No files found.',
            uri: `https://graph.microsoft.com/v1.0/users/${userId}/drive/root`,
            mimeType: 'text/plain',
          },
        },
      ];
    }

    return {
      content,
      _meta: {
        usedFallback,
        fileCount: content.length,
        query,
        userId,
      },
      isError: false,
    };
  },
};
