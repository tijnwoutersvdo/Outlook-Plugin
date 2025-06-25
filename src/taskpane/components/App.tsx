// File: src/taskpane/components/App.tsx

import * as React from "react";
import { useEffect, useState } from "react";
import { findBestMatch } from "string-similarity";
import { makeStyles } from "@fluentui/react-components";
import { getGraphToken } from "../authConfig";
import { getSiteAndDrive } from "../graph";
import { getAttachments, IAttachment } from "../taskpane";
import { ContactForm } from "./ContactForm";

async function computeFolderSuggestions(
  token: string,
  driveId: string,
  parentId: string,
  fileNames: string[],
  mailSubject: string
): Promise<{ name: string; path: string[] }[]> {
  const headers = { Authorization: `Bearer ${token}` };

  const collect = async (
    id: string,
    current: string[],
    depth: number
  ): Promise<{ id: string; name: string; path: string[] }[]> => {
    if (depth > 5) return [];
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${id}/children`,
      { headers }
    );
    if (!res.ok) throw new Error(`Mappen laden faalde: ${res.status}`);
    const data = await res.json();
    const folders = data.value
      .filter((it: any) => it.folder)
      .map((it: any) => ({
        id: it.id,
        name: it.name,
        path: [...current, it.id],
      }));
    const all = [...folders];
    for (const f of folders) {
      all.push(...(await collect(f.id, f.path, depth + 1)));
    }
    return all;
  };

  const allFolders = await collect(parentId, [parentId], 1);

  const queries = [
    ...fileNames.map(n => n.replace(/\.[^/.]+$/, "").toLowerCase()),
    mailSubject.toLowerCase(),
  ];

  const rated = allFolders.map(f => {
    const { bestMatch } = findBestMatch(f.name.toLowerCase(), queries);
    return { folder: f, rating: bestMatch.rating };
  });

  return rated
    .filter(r => r.rating >= 0.3)
    .sort((a, b) => b.rating - a.rating)
    .slice(0, 2)
    .map(r => ({ name: r.folder.name, path: r.folder.path }));
}

const useStyles = makeStyles({
  root:        { minHeight: "100vh", padding: "16px", fontFamily: "Segoe UI, sans-serif" },
  section:     { marginBottom: "16px" },
  breadcrumb:  { display: "flex", gap: "4px", alignItems: "center", marginBottom: "8px" },
  crumbButton: { background: "none", border: "none", color: "#0067B8", cursor: "pointer", padding: 0, fontSize: "1rem" },
  list:        { listStyleType: "none", padding: 0, margin: 0 },
  item:        { cursor: "pointer", padding: "4px 8px", borderRadius: "4px", marginBottom: "4px", backgroundColor: "#f3f2f1" },
  input:       { marginRight: "8px", padding: "4px 8px", fontSize: "1rem" },
  button:      { padding: "8px 16px", fontSize: "1rem", cursor: "pointer", marginTop: "8px" },
  error:       { color: "red", marginBottom: "16px" },
  success:     { color: "green", marginBottom: "16px" },
});

const App: React.FC = () => {
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
  const [mailSubject, setMailSubject]         = useState<string>("");
  const [suggestions, setSuggestions] = useState<{ name: string; path: string[] }[]>([]);

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

  const navigateToPathIds = async (pathIds: string[]) => {
    if (!graphToken || !driveId) return;
    try {
      const headers = { Authorization: `Bearer ${graphToken}` };
      const newPath: { id: string; name: string }[] = [];
      for (const id of pathIds) {
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${id}`,
          { headers }
        );
        if (!res.ok) throw new Error(`Map ophalen faalde: ${res.status}`);
        const data = await res.json();
        newPath.push({ id: data.id, name: data.name });
      }
      setPath(newPath);
      await loadSubfolders(graphToken, driveId, pathIds[pathIds.length - 1]);
    } catch (e: any) {
      setError("Navigeren mislukt: " + e.message);
    }
  };

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
    Office.onReady().then(async () => {
      try {
        const atts = await getAttachments();
        setAttachments(atts);
        const initial = atts
          .filter(a => !a.name.toLowerCase().includes("image"))
          .map(a => a.id);
        setSelectedIds(initial);
        const item: any = Office.context.mailbox.item;
        setMailSubject(item?.subject || "");
      } catch (e: any) {
        setError("Kon bijlagen niet laden: " + e.message);
      }
    });
  }, []);

  useEffect(() => {
    if (!graphToken || !driveId || attachments.length === 0) return;
    const currentId = path[path.length - 1].id;
    const names = attachments.map(a => a.name);
    computeFolderSuggestions(
      graphToken,
      driveId,
      currentId,
      names,
      mailSubject
    )
      .then(setSuggestions)
      .catch(e => console.error(e));
  }, [graphToken, driveId, path, attachments, mailSubject]);

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
          {suggestions.length > 0 && (
            <div className={styles.section}>
              <strong>Suggestie:</strong>{" "}
              {suggestions.map(s => (
                <button
                  key={s.path.join("-")}
                  onClick={() => navigateToPathIds(s.path)}
                >
                  {s.name}
                </button>
              ))}
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

export default App;

