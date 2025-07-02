// File: src/taskpane/components/App.tsx

import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { makeStyles } from "@fluentui/react-components";
import { getGraphToken } from "../authConfig";
import { getSiteAndDrive, getDriveTree, FolderNode } from "../graph";
import { getAttachments, IAttachment } from "../taskpane";
import { ContactForm } from "./ContactForm";

const useStyles = makeStyles({
  root:        { minHeight: "100vh", padding: "16px", fontFamily: "Segoe UI, sans-serif" },
  section:     { marginBottom: "16px" },
  breadcrumb:  { display: "flex", gap: "4px", alignItems: "center", marginBottom: "8px" },
  crumbButton: { background: "none", border: "none", color: "#0067B8", cursor: "pointer", padding: 0, fontSize: "1rem" },
  list:        { listStyleType: "none", padding: 0, margin: 0 },
  item:        { cursor: "pointer", padding: "4px 8px", borderRadius: "4px", marginBottom: "4px", backgroundColor: "#f3f2f1" },
  input:       { marginRight: "8px", padding: "4px 8px", fontSize: "1rem" },
  suggestion:  { marginBottom: "8px", backgroundColor: "#e7f3ff", padding: "8px", borderRadius: "4px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" },
  dismiss:     { marginLeft: "8px", cursor: "pointer", fontWeight: "bold" },
  button:      { padding: "8px 16px", fontSize: "1rem", cursor: "pointer", marginTop: "8px" },
  error:       { color: "red", marginBottom: "16px" },
  success:     { color: "green", marginBottom: "16px" },
});

const App: React.FC = () => {
  console.log("🚀 App component rendered");
  const styles = useStyles();

  // If mode=contact, render the ContactForm
  const params = new URLSearchParams(window.location.search);
  if (params.get("mode") === "contact") {
    return <ContactForm />;
  }

  // State for File Saver UI
  const [isSignedIn, setIsSignedIn]           = useState(false);
  const [error, setError]                     = useState<string | null>(null);
  const [success, setSuccess]                 = useState<string | null>(null);
  const [graphToken, setGraphToken]           = useState<string | null>(null);
  const [siteId, setSiteId]                   = useState<string | null>(null);
  const [driveId, setDriveId]                 = useState<string | null>(null);
  const [path, setPath]                       = useState<{ id: string; name: string }[]>([
    { id: "root", name: "Shared Documents" }
  ]);
  const [folders, setFolders]                 = useState<{ id: string; name: string }[]>([]);
  const [attachments, setAttachments]         = useState<IAttachment[]>([]);
  const [selectedIds, setSelectedIds]         = useState<string[]>([]);
  const [newFolderName, setNewFolderName]     = useState<string>("");
  const [suggestion, setSuggestion]           = useState<FolderNode | null>(null);
  const [treeLoaded, setTreeLoaded] = useState(false);
  const treeRef = useRef<FolderNode[]>([]);
  

  // Load subfolders under a parent
  const loadSubfolders = async (token: string, drive: string, parentId: string) => {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive}/items/${parentId}/children`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Mappen laden faalde: ${res.status}`);
    const data = await res.json();
    const subs = data.value
      .filter((item: any) => item.folder)
      .map((item: any) => ({ id: item.id, name: item.name }));
    setFolders(subs);
  };

  // Fetch complete folder tree (2 levels) and cache in treeRef

  // Sign in and initialize SharePoint context
  const signInAndLoad = async () => {
    setError(null);
    setSuccess(null);
    try {
      const token = await getGraphToken();
      setGraphToken(token);
      const ids = await getSiteAndDrive(token);
      setSiteId(ids.siteId);
      setDriveId(ids.driveId);
      treeRef.current = await getDriveTree(token, ids.driveId);
      console.log("Drive tree fetched, nodes:", treeRef.current.length);
      setTreeLoaded(true);
      await loadSubfolders(token, ids.driveId, "root");
      setIsSignedIn(true);
    } catch (e: any) {
      console.error(e);
      setError("Inloggen of laden mislukt: " + e.message);
    }
  };

  // Toggle selection of an attachment
  const toggleSelect = (id: string) => {
    setSelectedIds(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
    );
  };

  const acceptSuggestion = (node: FolderNode) => {
    const newPath = node.pathIds.map((id, idx) => ({
      id,
      name: node.pathNames[idx],
    }));
    setPath(newPath);
    if (graphToken && driveId) {
      loadSubfolders(graphToken, driveId, node.id);
    }
    setSuggestion(null);
  };

  // Upload selected attachments to SharePoint, optionally in new folder
  const uploadToSharePoint = async () => {
    setError(null);
    setSuccess(null);
    if (!graphToken || !driveId) {
      setError("Niet ingelogd");
      return;
    }

    let parentId = path[path.length - 1].id;
    let targetFolderId = parentId;

    if (newFolderName.trim()) {
      // Create new folder or get existing
      const createRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentId}/children`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${graphToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ name: newFolderName.trim(), folder: {} })
        }
      );
      if (createRes.ok) {
        const folder = await createRes.json();
        targetFolderId = folder.id;
      } else if (createRes.status === 409) {
        // Folder exists: find it
        const listRes = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentId}/children`,
          { headers: { Authorization: `Bearer ${graphToken}` } }
        );
        const listData = await listRes.json();
        const existing = listData.value.find(
          (i: any) => i.folder && i.name === newFolderName.trim()
        );
        if (!existing) {
          setError("Map bestaat maar niet gevonden");
          return;
        }
        targetFolderId = existing.id;
      } else {
        setError(`Mapcreatie faalde: ${createRes.status}`);
        return;
      }
    }

    try {
      for (const att of attachments.filter(a =>
        selectedIds.includes(a.id)
      )) {
        const result: any = await new Promise(res =>
          Office.context.mailbox.item.getAttachmentContentAsync(att.id, {}, res)
        );
        if (
          result.status === Office.AsyncResultStatus.Succeeded &&
          result.value.format === "base64"
        ) {
          const arrayBuffer = Uint8Array.from(
            atob(result.value.content),
            c => c.charCodeAt(0)
          );
          await fetch(
            `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${targetFolderId}:/${att.name}:/content`,
            {
              method: "PUT",
              headers: {
                Authorization: `Bearer ${graphToken}`,
                "Content-Type": "application/octet-stream"
              },
              body: arrayBuffer
            }
          );
        }
      }
      setSuccess(
        `Bijlagen succesvol geüpload naar “${
          newFolderName.trim() || "huidige map"
        }”.`
      );
      setNewFolderName("");
      await loadSubfolders(graphToken, driveId, parentId);
    } catch (e: any) {
      console.error(e);
      setError("Upload mislukt: " + e.message);
    }
  };

  // On mount: load attachments
  useEffect(() => {
  Office.onReady(() => {
    console.log("🛠️  File-saver pane Office.onReady fired");

    // Immediately invoked async function to load attachments
    (async () => {
      try {
        const atts = await getAttachments();
        console.log("📎 Attachments loaded:", atts.length);
        setAttachments(atts);

        // Auto-select all non-image attachments
        const initial = atts
          .filter(a => !a.name.toLowerCase().includes("image"))
          .map(a => a.id);
        setSelectedIds(initial);
      } catch (e: any) {
        console.error("❌ Error loading attachments:", e);
        setError("Kon bijlagen niet laden: " + e.message);
      }
    })();
  });
}, []);


 // Suggestion effect
useEffect(() => {
  console.log("🏷️  Suggestion effect 🔄", {
    attachments: attachments.length,
    selected:    selectedIds.length,
    driveId,
    treeLoaded,
    treeSize:    treeRef.current.length
  });

  if (!treeLoaded) {
    console.log("❌ Tree not loaded, skipping suggestion");
    setSuggestion(null);
    return;
  }

  const firstAtt = attachments.find(a => selectedIds.includes(a.id));
  if (!firstAtt) {
    console.log("❌ No attachment selected");
    setSuggestion(null);
    return;
  }

  const filename = firstAtt.name.toLowerCase();
  console.log("🔎 Matching for filename:", filename);

  // Break filename into tokens
  const tokens = filename.split(/[^a-z0-9]+/).filter(Boolean);
  console.log("🔑 Tokens:", tokens);

  // Find the Prospects root in the tree
  const prospects = treeRef.current.find(n => n.name === "Prospects");
  if (!prospects) {
    console.log("❌ 'Prospects' folder not found in tree");
    setSuggestion(null);
    return;
  }

  // Build all contiguous substrings of tokens
  const substrings: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    let accum = tokens[i];
    substrings.push(accum);
    for (let j = i + 1; j < tokens.length; j++) {
      accum += " " + tokens[j];
      substrings.push(accum);
    }
  }

  let bestMatch: FolderNode | null = null;
  let bestLen = 0;

  // Recursive search only within Prospects subtree
  const searchNode = (node: FolderNode) => {
    const fullPath = node.pathNames.join("/").toLowerCase();
    for (const sub of substrings) {
      if (fullPath.includes(sub) && sub.length > bestLen) {
        bestLen = sub.length;
        bestMatch = node;
      }
    }
    node.children.forEach(searchNode);
  };
  searchNode(prospects);

  if (bestMatch) {
    console.log("✅ Best match:", bestMatch.pathNames.join("/"), "(length:", bestLen, ")");
    setSuggestion(bestMatch);
  } else {
    console.log("🔄 No match found; falling back to Prospects");
    setSuggestion(prospects);
  }
}, [attachments, selectedIds, driveId, treeLoaded]);

  return (
    <div className={styles.root}>
      {error && <div className={styles.error}>{error}</div>}
      {success && <div className={styles.success}>{success}</div>}

      {!isSignedIn ? (
        <button className={styles.button} onClick={signInAndLoad}>
          Sign in to SharePoint
        </button>
      ) : (
        <>
          {suggestion && (
            <div
              className={styles.suggestion}
              onClick={() => acceptSuggestion(suggestion)}
            >
              Save file here? <strong>{suggestion.pathNames.join('/')}</strong>
              <span
                className={styles.dismiss}
                onClick={() => setSuggestion(null)}
              >
                ×
              </span>
            </div>
          )}
          <div className={styles.breadcrumb}>
            {path.map((crumb, idx) => (
              <React.Fragment key={crumb.id}>
                <button
                  className={styles.crumbButton}
                  onClick={() => {
                    const newPath = path.slice(0, idx + 1);
                    setPath(newPath);
                    if (graphToken && driveId) {
                      loadSubfolders(graphToken, driveId, newPath[newPath.length - 1].id);
                    }
                  }}
                >
                  {crumb.name}
                </button>
                {idx < path.length - 1 && <span>&gt;</span>}
              </React.Fragment>
            ))}
          </div>

          <ul className={styles.list}>
            {folders.map(f => (
              <li
                key={f.id}
                className={styles.item}
                onClick={() => {
                  setPath(prev => [...prev, f]);
                  if (graphToken && driveId) {
                    loadSubfolders(graphToken, driveId, f.id);
                  }
                }}
              >
                📂 {f.name}
              </li>
            ))}
          </ul>

          <div className={styles.section}>
            <input
              className={styles.input}
              placeholder="Nieuwe mapnaam (optioneel)"
              value={newFolderName}
              onChange={e => setNewFolderName(e.target.value)}
            />
            <button className={styles.button} onClick={uploadToSharePoint}>
              Upload naar SharePoint
            </button>
          </div>

          <div className={styles.section}>
            <ul className={styles.list}>
              {attachments.map(a => (
                <li key={a.id} className={styles.item}>
                  <label>
                    <input
                      type="checkbox"
                      checked={selectedIds.includes(a.id)}
                      onChange={() => toggleSelect(a.id)}
                    />{" "}
                    {a.name} ({Math.round(a.size / 1024)} KB)
                  </label>
                </li>
              ))}
            </ul>
          </div>
        </>
      )}
    </div>
  );
};

export default App
