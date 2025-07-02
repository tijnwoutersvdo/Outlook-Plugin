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
    "Content-Type":  "application/json",
  };
  const body = {
    givenName:       info.name,
    emailAddresses: [{ address: info.email, name: info.name }],
    businessPhones:  [ info.phone ],
    companyName:     info.organization,
    homeAddress:     { postalCode: info.postcode }
  };
  const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
    method:  "POST",
    headers,
    body:    JSON.stringify(body)
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`createContact failed: ${res.status} ${text}`);
  }
}

/**
 * A node in our folder‐tree:
 */
export interface FolderNode {
  id:        string;
  name:      string;
  children:  FolderNode[];
  pathIds:   string[];
  pathNames: string[];
  path:      string;    // convenience: pathNames.join(" / ")
}

/**
 * Look up the site & drive IDs for our SharePoint library.
 */
export async function getSiteAndDrive(token: string): Promise<{ siteId: string; driveId: string }> {
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type":   "application/json",
  };

  // 1) Fetch the site by fixed path
  const siteRes = await fetch(
    "https://graph.microsoft.com/v1.0/sites/synergiacapital.sharepoint.com:/sites/Data",
    { headers }
  );
  if (!siteRes.ok) {
    throw new Error(`getSiteAndDrive: site lookup failed ${siteRes.status}`);
  }
  const siteJson = await siteRes.json();
  const siteId   = siteJson.id;

  // 2) Fetch the **default** drive for that site
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
 * Recursively fetch up to 2 levels under “Prospects”, **using path-based URLs**.
 */
async function loadProspectsSubtree(
  driveId:   string,
  pathIds:   string[],
  pathNames: string[],
  headers:   Record<string,string>,
  depth:     number
): Promise<FolderNode> {
  // Depth guard
  if (depth >= 2) {
    const id   = pathIds[pathIds.length - 1];
    const name = pathNames[pathNames.length - 1];
    return { id, name, children: [], pathIds, pathNames, path: pathNames.join(" / ") };
  }

  // Build a relative path under the drive root, skipping the first segment ("Shared Documents")
  const relSegments = pathNames.slice(1).map(encodeURIComponent);
  const relPath     = relSegments.join("/");

  // Path‐based children fetch
  const childrenRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${relPath}:/children?$select=id,name,folder`,
    { headers }
  );
  if (!childrenRes.ok) {
    throw new Error(`loadProspectsSubtree: children fetch failed ${childrenRes.status}`);
  }
  const childrenJson = await childrenRes.json();

  const children: FolderNode[] = [];
  for (const child of childrenJson.value.filter((i: any) => i.folder)) {
    const newIds   = [...pathIds, child.id];
    const newNames = [...pathNames, child.name];
    const node     = await loadProspectsSubtree(
      driveId, newIds, newNames, headers, depth + 1
    );
    children.push(node);
  }

  return {
    id:        pathIds[pathIds.length - 1],
    name:      pathNames[pathNames.length - 1],
    children,
    pathIds,
    pathNames,
    path:      pathNames.join(" / ")
  };
}

/**
 * Build the full tree for folder‐suggestions:
 *  • Fetch root‐level folders of the drive
 *  • If we see “Shared”, drill into it (its children)
 *  • Under “Prospects”, recurse two levels deep via loadProspectsSubtree
 */
export async function getDriveTree(token: string, driveId: string): Promise<FolderNode[]> {
  const headers = { Authorization: `Bearer ${token}` };

  // 1) List root children
  const rootRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder`,
    { headers }
  );
  if (!rootRes.ok) {
    throw new Error(`getDriveTree: root fetch failed ${rootRes.status}`);
  }
  const rootJson = await rootRes.json();

  const nodes: FolderNode[] = [];
  // 2) Look for the special “Shared” container
  for (const item of rootJson.value.filter((i: any) => i.folder)) {
    if (item.name === "Shared") {
      // drill into Shared’s children
      const sharedRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/children?$select=id,name,folder`,
        { headers }
      );
      if (!sharedRes.ok) {
        throw new Error(`getDriveTree: Shared children fetch failed ${sharedRes.status}`);
      }
      const sharedJson = await sharedRes.json();

      for (const sharedChild of sharedJson.value.filter((i: any) => i.folder)) {
        const baseIds   = ["root", item.id, sharedChild.id];
        const baseNames = ["Shared Documents", item.name, sharedChild.name];

        if (sharedChild.name === "Prospects") {
          const subtree = await loadProspectsSubtree(
            driveId, baseIds, baseNames, headers, 0
          );
          nodes.push(subtree);
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
  }
  return nodes;
}


