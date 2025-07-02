// File: src/graph.ts

// ——————————————————————————————————————————————————————————————————————————————————
// Site & Drive lookup
// ——————————————————————————————————————————————————————————————————————————————————
export interface SiteAndDrive {
  siteId:  string;
  driveId: string;
}

/**
 * Given a Graph access token, returns the SharePoint siteId and
 * the driveId for 'Shared Documents' (the default Documents library).
 */
export async function getSiteAndDrive(token: string): Promise<SiteAndDrive> {
  const headers = { Authorization: `Bearer ${token}` };

  // 1) Lookup the root SharePoint site
  const siteRes = await fetch("https://graph.microsoft.com/v1.0/sites/root", { headers });
  if (!siteRes.ok) {
    throw new Error(`getSiteAndDrive: site lookup failed ${siteRes.status}`);
  }
  const siteJson = await siteRes.json();
  const siteId   = siteJson.id;

  // 2) Enumerate document libraries (drives) under that site
  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers }
  );
  if (!drivesRes.ok) {
    throw new Error(`getSiteAndDrive: drives lookup failed ${drivesRes.status}`);
  }
  const drivesJson = await drivesRes.json();

  // 3) Find the Documents library ("Shared Documents")
  const docDrive = drivesJson.value.find((d: any) =>
    d.name === "Documents" || d.name === "Shared Documents"
  );
  if (!docDrive) {
    throw new Error("getSiteAndDrive: 'Shared Documents' drive not found");
  }

  return { siteId, driveId: docDrive.id };
}

// ——————————————————————————————————————————————————————————————————————————————————
// Folder tree and FolderNode
// ——————————————————————————————————————————————————————————————————————————————————
export interface FolderNode {
  id:         string;
  name:       string;
  children:   FolderNode[];
  pathIds:    string[];
  pathNames:  string[];
  path:       string;
}

/**
 * Fetch only the first-level children of the drive’s root,
 * then for “Prospects” fetch two levels deep so your suggestion
 * algorithm has enough data.
 */
export async function getDriveTree(token: string, driveId: string): Promise<FolderNode[]> {
  const headers = { Authorization: `Bearer ${token}` };

  // ── Fetch the root children (only folders) ───────────────────────────
  const rootRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder`,
    { headers }
  );
  if (!rootRes.ok) {
    throw new Error(`getDriveTree: root fetch failed ${rootRes.status}`);
  }
  const rootJson = await rootRes.json();

  // Map each top-level folder into a FolderNode
  const nodes: FolderNode[] = [];
  for (const item of rootJson.value.filter((i: any) => i.folder)) {
    const baseIds   = ["root", item.id];
    const baseNames = ["Shared Documents", item.name];
    const basePath  = baseNames.join(" / ");

    if (item.name === "Prospects") {
      // Under Prospects: load two levels deep
      const prospectsNode = await loadProspectsSubtree(
        item.id,
        baseIds,
        baseNames,
        headers,
        /*depth=*/0
      );
      nodes.push(prospectsNode);
    } else {
      // Other folders: no children
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
 * Recursively fetches children under Prospects up to 2 levels deep.
 */
async function loadProspectsSubtree(
  itemId:    string,
  pathIds:   string[],
  pathNames: string[],
  headers:   Record<string,string>,
  depth:     number
): Promise<FolderNode> {
  // Stop at two levels
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

  // Fetch this folder’s children
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${headers.Authorization!.split(" ")[1]}/items/${itemId}/children?$select=id,name,folder`,
    { headers }
  );
  if (!res.ok) {
    throw new Error(`loadProspectsSubtree: fetch failed ${res.status}`);
  }
  const json    = await res.json();
  const folders = json.value.filter((i: any) => i.folder);

  // Recurse one level deeper
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
