// File: src/graph.ts

/** Payload for creating a new Outlook contact. */
export interface ContactInfo {
  name:         string;
  email:        string;
  phone:        string;
  organization: string;
  postcode:     string;
}

/** Create a new contact via Microsoft Graph. */
export async function createContact(token: string, info: ContactInfo): Promise<void> {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
    method:  "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      givenName:      info.name,
      emailAddresses: [{ address: info.email, name: info.name }],
      businessPhones: [ info.phone ],
      companyName:    info.organization,
      homeAddress:    { postalCode: info.postcode }
    })
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`createContact failed: ${res.status} ${text}`);
  }
}

// ────────────────────────────────────────────────────────────────────────────────
// Site & Drive lookup
// ────────────────────────────────────────────────────────────────────────────────
export interface SiteAndDrive {
  siteId:  string;
  driveId: string;
}

/**
 * Returns the SharePoint siteId and the **default** document-library driveId
 * by calling `/sites/{siteId}/drive`. This always works, regardless of the
 * library’s display name.
 */
export async function getSiteAndDrive(token: string): Promise<SiteAndDrive> {
  const headers = { Authorization: `Bearer ${token}` };

  // 1) Get the root site
  const siteRes = await fetch("https://graph.microsoft.com/v1.0/sites/root", { headers });
  if (!siteRes.ok) {
    throw new Error(`getSiteAndDrive: site lookup failed ${siteRes.status}`);
  }
  const siteJson = await siteRes.json();
  const siteId   = siteJson.id;

  // 2) Fetch the default drive for that site
  const driveRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`,
    { headers }
  );
  if (!driveRes.ok) {
    throw new Error(`getSiteAndDrive: default drive lookup failed ${driveRes.status}`);
  }
  const driveJson = await driveRes.json();
  const driveId   = driveJson.id;

  return { siteId, driveId };
}

// ────────────────────────────────────────────────────────────────────────────────
// Folder tree and suggestions
// ────────────────────────────────────────────────────────────────────────────────
export interface FolderNode {
  id:         string;
  name:       string;
  children:   FolderNode[];
  pathIds:    string[];
  pathNames:  string[];
  path:       string;
}

/**
 * Fetch top-level folders under the drive’s root, then recurse two levels
 * deep only under “Prospects” so suggestions remain fast.
 */
export async function getDriveTree(token: string, driveId: string): Promise<FolderNode[]> {
  const headers = { Authorization: `Bearer ${token}` };

  // 1) List first-level folders
  const rootRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder`,
    { headers }
  );
  if (!rootRes.ok) {
    throw new Error(`getDriveTree: root children fetch failed ${rootRes.status}`);
  }
  const rootJson = await rootRes.json();

  const nodes: FolderNode[] = [];
  for (const item of rootJson.value.filter((i: any) => i.folder)) {
    const baseIds   = ["root", item.id];
    const baseNames = ["Shared Documents", item.name];
    const basePath  = baseNames.join(" / ");

    if (item.name === "Prospects") {
      // Only under Prospects do we go two levels deep
      const prospectsNode = await loadProspectsSubtree(
        item.id, baseIds, baseNames, headers, /*depth=*/0
      );
      nodes.push(prospectsNode);
    } else {
      nodes.push({
        id:        item.id,
        name:      item.name,
        children:  [],
        pathIds:   baseIds,
        pathNames: baseNames,
        path:      basePath
      });
    }
  }
  return nodes;
}

/**
 * Helper: recursively fetch children up to 2 levels under “Prospects”.
 */
async function loadProspectsSubtree(
  itemId:    string,
  pathIds:   string[],
  pathNames: string[],
  headers:   Record<string,string>,
  depth:     number
): Promise<FolderNode> {
  if (depth >= 2) {
    return {
      id:        itemId,
      name:      pathNames[pathNames.length - 1],
      children:  [],
      pathIds,
      pathNames,
      path:      pathNames.join(" / ")
    };
  }

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${headers.Authorization!.split(" ")[1]}/items/${itemId}/children?$select=id,name,folder`,
    { headers }
  );
  if (!res.ok) {
    throw new Error(`loadProspectsSubtree: fetch failed ${res.status}`);
  }
  const json    = await res.json();
  const folders = json.value.filter((i: any) => i.folder);

  const children = await Promise.all(
    folders.map((f: any) =>
      loadProspectsSubtree(
        f.id,
        [...pathIds, f.id],
        [...pathNames, f.name],
        headers,
        depth + 1
      )
    )
  );

  return {
    id:        itemId,
    name:      pathNames[pathNames.length - 1],
    children,
    pathIds,
    pathNames,
    path:      pathNames.join(" / ")
  };
}

