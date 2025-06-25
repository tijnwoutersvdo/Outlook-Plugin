/* global Office console */

/**
 * Interface voor e-mailbijlagen
 */
export interface IAttachment {
  id: string;
  name: string;
  size: number;
}

/**
 * Haalt attachments uit de huidige mail en retourneert ze als array.
 */
export async function getAttachments(): Promise<IAttachment[]> {
  const item = Office.context.mailbox.item as any;
  if (!item.attachments || !item.attachments.length) {
    return [];
  }
  return item.attachments.map((att: any) => ({
    id: att.id,
    name: att.name,
    size: att.size,
  }));
}