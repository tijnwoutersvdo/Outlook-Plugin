import * as React from "react";
import { useState, useEffect } from "react";
import { makeStyles } from "@fluentui/react-components";
import JSZip from "jszip";
import { saveAs } from "file-saver";

interface Attachment {
  id: string;
  name: string;
  size: number;
}

const useStyles = makeStyles({
  root: { minHeight: "100vh", padding: "16px", fontFamily: "Segoe UI, sans-serif" },
  folderSection: { marginBottom: "16px" },
  folderInput: { marginRight: "8px", padding: "4px 8px", fontSize: "1rem" },
  attachmentsSection: { marginBottom: "16px" },
  attachmentList: { listStyleType: "none", padding: 0 },
  attachmentItem: { display: "flex", alignItems: "center", marginBottom: "8px" },
  checkboxLabel: { display: "flex", alignItems: "center", cursor: "pointer" },
  checkbox: { marginRight: "8px" },
  downloadButton: { padding: "8px 16px", fontSize: "1rem", cursor: "pointer" },
  error: { color: "red", marginBottom: "16px" }
});

const App: React.FC = () => {
  // Path naar je logo, zorg dat logo.png in public/assets staat
  

  const styles = useStyles();
  const [folderName, setFolderName] = useState<string>("");
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);

  useEffect(() => {
    Office.onReady().then(() => {
      const item = Office.context.mailbox.item as any;
      if (item.attachments?.length) {
        const atts: Attachment[] = item.attachments.map((att: any) => ({ id: att.id, name: att.name, size: att.size }));
        setAttachments(atts);
        setSelectedIds(atts.map(a => a.id));
      }
    }).catch(e => setError(`Kon add-in niet initialiseren: ${e instanceof Error ? e.message : e}`));
  }, []);

  const toggleSelect = (id: string) => {
    setSelectedIds(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
  };

  const downloadSelected = async () => {
    for (const att of attachments.filter(a => selectedIds.includes(a.id))) {
      const result: any = await new Promise(res => Office.context.mailbox.item.getAttachmentContentAsync(att.id, {}, res));
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.format === 'base64') {
        const buffer = Uint8Array.from(atob(result.value.content), c => c.charCodeAt(0)).buffer;
        const blob = new Blob([buffer]);
        saveAs(blob, att.name);
      } else {
        console.error('Fout bij ophalen:', result.error);
      }
    }
  };

  const downloadAsZip = async () => {
    if (!folderName) { setError('Voer eerst een mapnaam in.'); return; }
    const zip = new JSZip();
    const folder = zip.folder(folderName)!;
    for (const att of attachments.filter(a => selectedIds.includes(a.id))) {
      const result: any = await new Promise(res => Office.context.mailbox.item.getAttachmentContentAsync(att.id, {}, res));
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.format === 'base64') {
        const buffer = Uint8Array.from(atob(result.value.content), c => c.charCodeAt(0)).buffer;
        folder.file(att.name, buffer);
      }
    }
    const blob = await zip.generateAsync({ type: 'blob' });
    saveAs(blob, `${folderName}.zip`);
  };

  return (
    <div className={styles.root}>
      {/* Logo bovenaan */}
      
      {error && <div className={styles.error}>{error}</div>}
      <div className={styles.folderSection}>
        <input
          className={styles.folderInput}
          type="text"
          aria-label="Mapnaam"
          placeholder="Mapnaam (e.g. Project A)"
          value={folderName}
          onChange={e => setFolderName(e.target.value)}
        />
      </div>
      <div className={styles.attachmentsSection}>
        <ul className={styles.attachmentList}>
          {attachments.map(att => (
            <li key={att.id} className={styles.attachmentItem}>
              <label className={styles.checkboxLabel}>
                <input
                  className={styles.checkbox}
                  type="checkbox"
                  checked={selectedIds.includes(att.id)}
                  onChange={() => toggleSelect(att.id)}
                />
                {att.name} ({Math.round(att.size / 1024)} KB)
              </label>
            </li>
          ))}
        </ul>
        {attachments.length > 0 && (
          <>
            <button
              className={styles.downloadButton}
              onClick={downloadSelected}
              disabled={selectedIds.length === 0}
            >
              Download geselecteerde bijlagen
            </button>
            <button
              className={styles.downloadButton}
              onClick={downloadAsZip}
              disabled={selectedIds.length === 0 || !folderName}
            >
              Download als map
            </button>
          </>
        )}
      </div>
    </div>
  );
};

export default App;

