// src/taskpane/storage.ts
// src/taskpane/storage.ts

/** Interface voor opslag-mechanismen */
export interface IStorage {
  /**
   * Sla één bijlage op.
   * @param name Naam van het bestand (inclusief extensie).
   * @param content De raw bytes van het bestand.
   */
  saveAttachment(name: string, content: ArrayBuffer): Promise<void>;
}

/**
 * Eenvoudige opslag-implementatie die een browser-download triggert.
 * Dynamisch importeert file-saver en gebruikt de juiste export.
 */
export class LocalFileSaver implements IStorage {
  async saveAttachment(name: string, content: ArrayBuffer): Promise<void> {
    try {
      // Dynamisch importeren van file-saver
      const module = await import("file-saver");
      // Gebruik named export of default export
      const save = (module as any).saveAs || (module as any).default;
      const blob = new Blob([content]);
      save(blob, name);
    } catch (e) {
      console.error("Fout bij LocalFileSaver saveAttachment:", e);
      throw e;
    }
  }
}

/**
 * Storage-implementatie via de File System Access API
 * Maakt bestanden in een gekozen map.
 */
export class FileSystemStorage implements IStorage {
  constructor(private dirHandle: FileSystemDirectoryHandle) {}

  async saveAttachment(name: string, content: ArrayBuffer): Promise<void> {
    const fileHandle = await this.dirHandle.getFileHandle(name, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(content);
    await writable.close();
  }
}
