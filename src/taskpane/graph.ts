// File: src/graph.ts

/**
 * Data needed to create a new contact.
 */
export interface ContactInfo {
  name:         string;
  email:        string;
  phone:        string;
  organization: string;
  postcode:     string;
}

/**
 * Creates a new Outlook contact via Microsoft Graph.
 */
export async function createContact(token: string, info: ContactInfo): Promise<void> {
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  const body = {
    givenName:      info.name,
    emailAddresses: [{ address: info.email, name: info.name }],
    businessPhones: [ info.phone ],
    companyName:    info.organization,
    homeAddress:    { postalCode: info.postcode }
  };

  const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
    method:  "POST",
    headers,
    body:    JSON.stringify(body)
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Contacts.Create faalde: ${res.status} ${text}`);
  }
}

/**
 * A node in our folder‐tree for suggestions.
 */
export interface FolderNode {
  id:        string;
  name:      string;
  children:  FolderNode[];
  pathIds:   string[];
  pathNames: string[];
  path:      string;               // convenience: joined pathNames
}

/**
 * Look up the site & drive IDs for our SharePoint library.
 */
export async function getSiteAndDrive(token: string): Promise<{ siteId: string; driveId: string }> {
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type":   "application/json",
  };

  // 1) Fetch the site
  const siteRes = await fetch(
    "https://graph.microsoft.com/v1.0/sites/synergiacapital.sharepoint.com:/sites/Data",
    { headers }
  );
  if (!siteRes.ok) {
    throw new Error(`getSiteAndDrive: site lookup failed ${siteRes.status}`);
  }
  const siteJson = await siteRes.json();
  const siteId   = siteJson.id;

  // 2) Fetch its default drive
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

/**
 * Recursively fetch the subtree under “Prospects” up to 2 levels deep.
 * Always uses the original driveId for its children lookups.
 */
async function loadProspectsSubtree(
  itemId:    string,
  driveId:   string,
  pathIds:   string[],
  pathNames: string[],
  headers:   Record<string,string>,
  depth:     number
): Promise<FolderNode> {
  // stop after 2 levels
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

  // fetch this folder’s children using the correct driveId
  const childrenRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$select=id,name,folder`,
    { headers }
  );
  if (!childrenRes.ok) {
    throw new Error(`loadProspectsSubtree: children fetch failed ${childrenRes.status}`);
  }
  const childrenJson = await childrenRes.json();

  const children: FolderNode[] = [];
  for (const child of childrenJson.value.filter((i: any) => i.folder)) {
    const childIds   = [...pathIds, child.id];
    const childNames = [...pathNames, child.name];

    // recurse deeper under Prospects
    const node = await loadProspectsSubtree(
      child.id, driveId, childIds, childNames, headers, depth + 1
    );
    children.push(node);
  }

  return {
    id:        itemId,
    name:      pathNames[pathNames.length - 1],
    children,
    pathIds,
    pathNames,
    path:      pathNames.join(" / ")
  };
}

/**
 * Build the full tree for folder‐suggestions:
 *  • Fetch root‐level folders of the drive.
 *  • If we see “Shared”, drill into it and treat *its* children as our base set.
 *  • Under “Prospects”, recurse two levels deep.
 */
export async function getDriveTree(token: string, driveId: string): Promise<FolderNode[]> {
  const headers = { Authorization: `Bearer ${token}` };

  // 1) Get root‐level child items
  const rootRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder`,
    { headers }
  );
  if (!rootRes.ok) {
    throw new Error(`getDriveTree: root fetch failed ${rootRes.status}`);
  }
  const rootJson = await rootRes.json();

  const nodes: FolderNode[] = [];

  for (const item of rootJson.value.filter((i: any) => i.folder)) {
    // Special case: if the library’s first child is “Shared”, drill into it
    if (item.name === "Shared") {
      const sharedRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/children?$select=id,name,folder`,
        { headers }
      );
      if (!sharedRes.ok) {
        throw new Error(`getDriveTree: Shared children fetch failed ${sharedRes.status}`);
      }
      const sharedJson = await sharedRes.json();

      for (const sharedChild of sharedJson.value.filter((i: any) => i.folder)) {
        const baseIds   = [ "root", item.id, sharedChild.id ];
        const baseNames = [ "Shared Documents", item.name, sharedChild.name ];

        if (sharedChild.name === "Prospects") {
          // Two‐level deep under Prospects
          const prospectsNode = await loadProspectsSubtree(
            sharedChild.id, driveId, baseIds, baseNames, headers, 0
          );
          nodes.push(prospectsNode);
        } else {
          nodes.push({
            id:        sharedChild.id,
            name:      sharedChild.name,
            children:  [],
            pathIds:   baseIds,
            pathNames: baseNames,
            path:      baseNames.join(" / ")
          });
        }
      }
    }
    // Also handle a top‐level “Prospects” folder if someone placed it there
    else if (item.name === "Prospects") {
      const baseIds   = [ "root", item.id ];
      const baseNames = [ "Shared Documents", item.name ];
      const prospectsNode = await loadProspectsSubtree(
        item.id, driveId, baseIds, baseNames, headers, 0
      );
      nodes.push(prospectsNode);
    }
    // ignore all other root folders for suggestion purposes
  }

  return nodes;
}


