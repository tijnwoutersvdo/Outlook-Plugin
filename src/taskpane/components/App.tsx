import * as React from "react";
import { useEffect, useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import { getGraphToken } from "../authConfig";
import { getSiteAndDrive } from "../graph";
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
  button:      { padding: "8px 16px", fontSize: "1rem", cursor: "pointer", marginTop: "8px" },
  error:       { color: "red", marginBottom: "16px" },
  success:     { color: "green", marginBottom: "16px" },
});

const App: React.FC = () => {
  const styles = useStyles();

  // if mode=contact, delegate to your ContactForm
  const params = new URLSearchParams(window.location.search);
  if (params.get("mode") === "contact") {
    return <ContactForm />;
  }

  // ─ State for File-Saver pane ──────────────────────────────────────────────
  const [isSignedIn, setIsSignedIn]       = useState(false);
  const [error, setError]                 = useState<string | null>(null);
  const [success, setSuccess]             = useState<string | null>(null);
  const [graphToken, setGraphToken]       = useState<string | null>(null);
  const [siteId, setSiteId]               = useState<string | null>(null);
  const [driveId, setDriveId]             = useState<string | null>(null);
  const [path, setPath]                   = useState<{ id: string; name: string }[]>([
    { id: "root", name: "Shared Documents" }
  ]);
  const [folders, setFolders]             = useState<{ id: string; name: string }[]>([]);
  const [attachments, setAttachments]     = useState<IAttachment[]>([]);
  const [selectedIds, setSelectedIds]     = useState<string[]>([]);
  const [newFolderName, setNewFolderName] = useState<string>("");

  // load children for a given folder
  const loadSubfolders = async (token: string, drive: string, parentId: string) => {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive}/items/${parentId}/children`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Mappen laden faalde: ${res.status}`);
    const data = await res.json();
    const subs = data.value
      .filter((i: any) => i.folder)
      .map((i: any) => ({ id: i.id, name: i.name }));
    setFolders(subs);
  };

  // sign in and prime pane
  const signInAndLoad = async () => {
    setError(null);
    setSuccess(null);
    try {
      const token = await getGraphToken();
      setGraphToken(token);

      const ids = await getSiteAndDrive(token);
      setSiteId(ids.siteId);
      setDriveId(ids.driveId);

      // load the root children
      await loadSubfolders(token, ids.driveId, "root");
      setIsSignedIn(true);
    } catch (e: any) {
      console.error(e);
      setError("Inloggen of laden mislukt: " + e.message);
    }
  };

  const toggleSelect = (id: string) => {
    setSelectedIds(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
    );
  };

  // upload attachments (identical to your original)
  const uploadToSharePoint = async () => {
    setError(null);
    setSuccess(null);
    if (!graphToken || !driveId) {
      setError("Niet ingelogd");
      return;
    }

    let parentId       = path[path.length - 1].id;
    let targetFolderId = parentId;

    if (newFolderName.trim()) {
      // create or find existing folder
      const createRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentId}/children`,
        {
          method:  "POST",
          headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
          body:    JSON.stringify({ name: newFolderName.trim(), folder: {} })
        }
      );
      if (createRes.ok) {
        targetFolderId = (await createRes.json()).id;
      } else if (createRes.status === 409) {
        const listRes   = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentId}/children`,
          { headers: { Authorization: `Bearer ${graphToken}` } }
        );
        const existing  = (await listRes.json()).value.find((i: any) =>
          i.folder && i.name === newFolderName.trim()
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
      for (const att of attachments.filter(a => selectedIds.includes(a.id))) {
        const result: any = await new Promise(res =>
          Office.context.mailbox.item.getAttachmentContentAsync(att.id, {}, res)
        );
        if (
          result.status === Office.AsyncResultStatus.Succeeded &&
          result.value.format === "base64"
        ) {
          const arrayBuffer = Uint8Array.from(
            atob(result.value.content), c => c.charCodeAt(0)
          );
          await fetch(
            `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${targetFolderId}:/${att.name}:/content`,
            {
              method:  "PUT",
              headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/octet-stream" },
              body:    arrayBuffer
            }
          );
        }
      }
      setSuccess(
        `Bijlagen succesvol geüpload naar “${newFolderName.trim() || "huidige map"}”.`
      );
      setNewFolderName("");
      await loadSubfolders(graphToken, driveId, parentId);
    } catch (e: any) {
      console.error(e);
      setError("Upload mislukt: " + e.message);
    }
  };

  // on mount: grab your attachments
  useEffect(() => {
    Office.onReady().then(async () => {
      try {
        const atts = await getAttachments();
        setAttachments(atts);
        const initial = atts
          .filter(a => !a.name.toLowerCase().includes("image"))
          .map(a => a.id);
        setSelectedIds(initial);
      } catch (e: any) {
        setError("Kon bijlagen niet laden: " + e.message);
      }
    });
  }, []);

  return (
    <div className={styles.root}>
      {error   && <div className={styles.error}>{error}</div>}
      {success && <div className={styles.success}>{success}</div>}

      {!isSignedIn
        ? <button className={styles.button} onClick={signInAndLoad}>
            Sign in to SharePoint
          </button>
        : <>
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
      }
    </div>
  );
};

export default App;
