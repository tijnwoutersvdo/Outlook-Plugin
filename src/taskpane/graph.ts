/**
 * Haalt via Microsoft Graph de siteId en driveId op voor de SharePoint-site.
 * @param token Bearer access token voor Microsoft Graph
 * @returns Promise met { siteId, driveId }
 */
export async function getSiteAndDrive(token: string): Promise<{ siteId: string; driveId: string }> {
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  // 1) Site ophalen via pad '/sites/Data'
  const siteRes = await fetch(
    "https://graph.microsoft.com/v1.0/sites/synergiacapital.sharepoint.com:/sites/Data",
    { headers }
  );
  if (!siteRes.ok) {
    throw new Error(`Fout bij ophalen site: ${siteRes.status} ${siteRes.statusText}`);
  }
  const siteData = await siteRes.json();
  const siteId: string = siteData.id;

  // 2) Drive (documentbibliotheek) ophalen voor deze site
  const driveRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`,
    { headers }
  );
  if (!driveRes.ok) {
    throw new Error(`Fout bij ophalen drive: ${driveRes.status} ${driveRes.statusText}`);
  }
  const driveData = await driveRes.json();
  const driveId: string = driveData.id;

  return { siteId, driveId };
}

export interface ContactInfo {
  name:         string;
  email:        string;
  phone:        string;
  organization: string;
  postcode:     string;
}

export async function createContact(token: string, info: ContactInfo): Promise<void> {
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type":  "application/json"
  };

  const body = {
    givenName:       info.name,
    emailAddresses: [{ address: info.email, name: info.name }],
    businessPhones:  [ info.phone ],
    companyName:     info.organization,
    homeAddress:    { postalCode: info.postcode }
  };

  const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
    method:  "POST",
    headers,
    body:    JSON.stringify(body)
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Contacts.Create faalde: ${res.status} ${text}`);
  }
}
