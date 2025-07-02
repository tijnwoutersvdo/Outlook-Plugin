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
 * A node in our folder‐tree for suggestions.
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

/**
 * Recursively fetch up to 2 levels under “Prospects”, using path-based URLs.
 */
async function loadProspectsSubtree(
  driveId:   string,
  pathIds:   string[],
  pathNames: string[],
  headers:   Record<string,string>,
  depth:     number
): Promise<FolderNode> {
  // Stop after 2 levels
  if (depth >= 2) {
    const id   = pathIds[pathIds.length - 1];
    const name = pathNames[pathNames.length - 1];
    return { id, name, children: [], pathIds, pathNames, path: pathNames.join(" / ") };
  }

  // Build a relative path under the drive root, skipping the first segment ("Shared Documents")
  const relSegments = pathNames.slice(1).map(encodeURIComponent);
  const relPath     = relSegments.join("/");

  // Path-based children fetch
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${relPath}:/children?$select=id,name,folder`;
  const childrenRes = await fetch(url, { headers });
  if (!childrenRes.ok) {
    const errText = await childrenRes.text();
    throw new Error(`loadProspectsSubtree: children fetch failed ${childrenRes.status}: ${errText}`);
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
 * Build the full tree for folder-suggestions:
 *  • List root children of the drive.
 *  • If a folder named “Shared” exists, drill into it and use its children.
 *  • Under “Prospects”, recurse two levels deep via loadProspectsSubtree.
 */
export async function getDriveTree(token: string, driveId: string): Promise<FolderNode[]> {
  const headers = { Authorization: `Bearer ${token}` };

  // 1) Get root-level child items
  const rootRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder`,
    { headers }
  );
  if (!rootRes.ok) {
    throw new Error(`getDriveTree: root fetch failed ${rootRes.status}`);
  }
  const rootJson = await rootRes.json();
  console.log("🔍 Root folders:", rootJson.value.map((i: any) => i.name));

  const nodes: FolderNode[] = [];

  // 2) Look for the special “Shared” container
  for (const item of rootJson.value.filter((i: any) => i.folder)) {
    if (item.name === "Shared") {
      // Drill into Shared’s children
      const sharedRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/children?$select=id,name,folder`,
        { headers }
      );
      if (!sharedRes.ok) {
        throw new Error(`getDriveTree: Shared children fetch failed ${sharedRes.status}`);
      }
      const sharedJson = await sharedRes.json();
      console.log("🔍 Shared → children:", sharedJson.value.map((i: any) => i.name));

      for (const sharedChild of sharedJson.value.filter((i: any) => i.folder)) {
        const baseIds   = ["root", item.id, sharedChild.id];
        const baseNames = ["Shared Documents", item.name, sharedChild.name];

        if (sharedChild.name === "Prospects") {
          const subtree = await loadProspectsSubtree(
            driveId, baseIds, baseNames, headers, 0
          );
          nodes.push(subtree);


        // ← New SCF* folders: pull in SCF III/IV/V/VI → Participaties → children
        } else if (/^SCF /.test(sharedChild.name)) {
          // fetch children of SCF N
          const scfRes = await fetch(
            `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${sharedChild.id}/children?$select=id,name,folder`,
            { headers }
          );
          const scfJson = await scfRes.json();
          // find the “Participaties” folder
          const part = scfJson.value.find((i: any) => i.folder && i.name === "Participaties");
          if (part) {
            // fetch its immediate children
            const childrenRes = await fetch(
              `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${part.id}/children?$select=id,name,folder`,
              { headers }
            );
            const childrenJson = await childrenRes.json();
            // build a node representing SCF → Participaties
            nodes.push({
              id:        part.id,
              name:      "Participaties",
              children:  childrenJson.value.filter((i: any) => i.folder).map((c: any) => ({
                id:        c.id,
                name:      c.name,
                children:  [],
                pathIds:   [...baseIds, part.id, c.id],
                pathNames: [...baseNames, "Participaties", c.name],
                path:      [...baseNames, "Participaties", c.name].join(" / ")
              })),
              pathIds:   [...baseIds, part.id],
              pathNames: [...baseNames, "Participaties"],
              path:      [...baseNames, "Participaties"].join(" / ")
            });
          }


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



