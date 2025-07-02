/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

/* global Office, OfficeRuntime */

interface ContactInfo {
  name: string;
  email: string;
  phone: string;
  organization: string;
  postcode: string;
}

let pending: { isNew: boolean; id?: string; info: ContactInfo } | null = null;

Office.onReady(() => {
  Office.actions.associate("onMessageOpenHandler", onMessageOpenHandler);
  Office.actions.associate("addContactYes", addContactYes);
  Office.actions.associate("addContactNo", addContactNo);
});

export async function onMessageOpenHandler(event: any) {
  const item = Office.context.mailbox.item;
  const senderName = item.from?.displayName || "";
  const senderEmail = item.from?.emailAddress || "";

  const rawSeed = senderEmail.split("@")[1]?.split(".")[0] || "";
  const organization = rawSeed
    ? rawSeed.charAt(0).toUpperCase() + rawSeed.slice(1).toLowerCase()
    : "";

  const body = await new Promise<string>((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (res) => {
      resolve(res.value || "");
    });
  });

  const sig = extractSignatureBlock(body, senderName);
  const info = parseSignature(sig, senderName, senderEmail, organization);

  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false,
    });
    const headers = {
      Authorization: `Bearer ${token}`,
    };
    const filter = `displayName eq '${senderName.replace("'", "''")}'`;
    const existingRes = await fetch(
      "https://graph.microsoft.com/v1.0/me/contacts?$filter=" +
        encodeURIComponent(filter),
      { headers }
    );
    const existingJson = await existingRes.json();
    const existing = existingJson.value?.[0];

    const existingEmail = existing?.emailAddresses?.[0]?.address || "";
    const existingPhone = existing?.businessPhones?.[0] || "";
    const existingCompany = existing?.companyName || "";

    const needsUpdate =
      existing &&
      (existingEmail !== info.email ||
        existingPhone !== info.phone ||
        existingCompany !== info.organization);

    if (!existing || needsUpdate) {
      pending = {
        isNew: !existing,
        id: existing?.id,
        info: { ...info },
      };

      const message = existing
        ? `Update information for "${info.name}"?`
        : `Add "${info.name}" to your contacts?`;

      item.notificationMessages.addAsync("contactPrompt", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message,
        icon: "Icon.80x80",
        persistent: true,
        actions: [
          { action: "addContactYes", title: "Yes" },
          { action: "addContactNo", title: "No" },
        ],
      });
    }
  } catch (err) {
    console.error(err);
  } finally {
    event.completed();
  }
}

export async function addContactYes(event: any) {
  if (!pending) {
    event.completed();
    return;
  }
  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false,
    });
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const body = {
      givenName: pending.info.name,
      emailAddresses: [{ address: pending.info.email, name: pending.info.name }],
      businessPhones: [pending.info.phone],
      companyName: pending.info.organization,
      homeAddress: { postalCode: pending.info.postcode },
    };
    if (pending.isNew) {
      await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
        method: "POST",
        headers,
        body: JSON.stringify(body),
      });
    } else if (pending.id) {
      await fetch(`https://graph.microsoft.com/v1.0/me/contacts/${pending.id}`, {
        method: "PATCH",
        headers,
        body: JSON.stringify(body),
      });
    }
  } catch (err) {
    console.error(err);
  } finally {
    Office.context.mailbox.item.notificationMessages.removeAsync("contactPrompt");
    pending = null;
    event.completed();
  }
}

export function addContactNo(event: any) {
  Office.context.mailbox.item.notificationMessages.removeAsync("contactPrompt");
  pending = null;
  event.completed();
}

function extractSignatureBlock(body: string, senderName: string): string {
  if (senderName) {
    const idx = body.indexOf(senderName);
    if (idx >= 0) return body.substring(idx).trim();
  }
  const parts = body.split(/\r?\n\s*\r?\n/);
  return (parts.length > 1 ? parts.pop() : body).trim();
}

function parseSignature(
  sig: string,
  senderName: string,
  senderEmail: string,
  organization: string
): ContactInfo {
  const lines = sig
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l);

  let name = "";
  if (senderName && sig.toLowerCase().includes(senderName.toLowerCase())) {
    name = senderName;
  } else {
    for (const line of lines) {
      if (line.includes(senderEmail)) continue;
      if (/\+?\d[\d\-\u2013()\s]{5,}\d/.test(line)) continue;
      if (/https?:\/\//i.test(line)) continue;
      if (/www\./i.test(line)) continue;
      if (/^[\+\d]/.test(line)) continue;
      name = line;
      break;
    }
  }

  const email = senderEmail;

  const phoneRegex = /(\+?\d[\d\-\u2013()\s]{5,}\d)/g;
  const matches = [] as string[];
  let m: RegExpExecArray | null;
  while ((m = phoneRegex.exec(sig))) matches.push(m[1]);
  const phone = matches.length
    ? matches.reduce((a, b) => (a.length >= b.length ? a : b))
    : "";

  const postcodeRegex = /\b\d{4}\s?[A-Za-z]{2}\b/;
  let postcode = "";
  for (const line of lines) {
    const mp = line.match(postcodeRegex);
    if (mp) {
      postcode = mp[0];
      break;
    }
  }

  return { name, email, phone, organization, postcode };
}

