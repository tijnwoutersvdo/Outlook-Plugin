// File: src/taskpane/components/ContactForm.tsx

import React, { useEffect, useState } from "react";
import { Button, Input, Text } from "@fluentui/react-components";
import { createContact } from "../graph";
import { extractSignatureBlock, parseSignature } from "./signature";
import { getGraphToken } from "../authConfig";

export function ContactForm() {
  // parsed info from signature
  const [info, setInfo] = useState({
    name: "",
    email: "",
    phone: "",
    organization: "",
    postcode: "",
  });

  // track where we are in the "check" lifecycle
  const [status, setStatus] = useState<
    "checking" | "not-found" | "exists-unchanged" | "exists-changed" | "idle"
  >("checking");
  const [existingId, setExistingId] = useState<string | null>(null);

  const [loading, setLoading] = useState(true);   // true while signature parsing
  const [saving, setSaving]     = useState(false); // true while POST/PATCH in flight
  const [statusMessage, setStatusMessage] = useState<
    { type: "success" | "error"; text: string } | null
  >(null);

  // 1) On load, grab the signature + parse into `info`
  useEffect(() => {
    Office.onReady(() => {
      const item = Office.context.mailbox.item;
      const senderName  = item.from?.displayName  || "";
      const senderEmail = item.from?.emailAddress || "";
      // derive org from email domain
      const rawSeed = senderEmail.split("@")[1]?.split(".")[0] || "";
      const organization = rawSeed
        ? rawSeed.charAt(0).toUpperCase() + rawSeed.slice(1).toLowerCase()
        : "";

      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const sig    = extractSignatureBlock(result.value, senderName);
          const parsed = parseSignature(sig, senderName, senderEmail, organization);
          setInfo(parsed);
        } else {
          console.error(result.error);
          setStatusMessage({ type: "error", text: "Error fetching email body." });
        }
        setLoading(false);
      });
    });
  }, []);

  // 2) Once `info` is set (loading === false), lookup in Graph
  useEffect(() => {
    if (loading) return; // wait for signature parse

    (async () => {
      setStatus("checking");
      try {
        const token = await getGraphToken();
        const filter = `emailAddresses/any(a:a/address eq '${info.email}')`;
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/me/contacts?$filter=${encodeURIComponent(filter)}`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        const json     = await res.json();
        const existing = json.value?.[0];
        if (!existing) {
          setStatus("not-found");
        } else {
          setExistingId(existing.id);
          const sameEmail = existing.emailAddresses[0]?.address === info.email;
          const samePhone = existing.businessPhones[0]       === info.phone;
          const sameOrg   = existing.companyName             === info.organization;
          setStatus(
            sameEmail && samePhone && sameOrg
              ? "exists-unchanged"
              : "exists-changed"
          );
        }
      } catch (e: any) {
        console.error(e);
        setStatusMessage({ type: "error", text: "Error checking existing contacts." });
        setStatus("idle");
      }
    })();
  }, [loading, info.email, info.phone, info.organization]);

  // update local state when form inputs change
  const updateField =
    (key: keyof typeof info) =>
    (e: React.ChangeEvent<HTMLInputElement>) =>
      setInfo({ ...info, [key]: e.target.value });

  // 3a) Create a fresh contact
  const onSave = async () => {
    setSaving(true);
    setStatusMessage(null);
    try {
      const token = await getGraphToken();
      await createContact(token, info);
      setStatusMessage({ type: "success", text: "Contact opgeslagen!" });
      setStatus("idle");
    } catch (e: any) {
      console.error(e);
      setStatusMessage({ type: "error", text: "Opslaan mislukt: " + e.message });
    } finally {
      setSaving(false);
    }
  };

  // 3b) Patch an existing contact
  const updateExisting = async () => {
    if (!existingId) return;
    setSaving(true);
    setStatusMessage(null);
    try {
      const token = await getGraphToken();
      await fetch(
        `https://graph.microsoft.com/v1.0/me/contacts/${existingId}`,
        {
          method:  "PATCH",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            givenName:      info.name,
            emailAddresses: [{ address: info.email, name: info.name }],
            businessPhones: [info.phone],
            companyName:    info.organization,
            homeAddress:    { postalCode: info.postcode },
          }),
        }
      );
      setStatusMessage({ type: "success", text: "Contact geüpdatet!" });
      setStatus("idle");
    } catch (e: any) {
      console.error(e);
      setStatusMessage({ type: "error", text: "Update mislukt: " + e.message });
    } finally {
      setSaving(false);
    }
  };

  // ── Early exits ──────────────────────────────────────────────────────
  if (loading) {
    return <Text>Loading signature…</Text>;
  }
  if (status === "checking") {
    return <Text>Checking your contacts…</Text>;
  }

  // ── Main UI ─────────────────────────────────────────────────────────
  return (
    <div style={{ padding: 20, maxWidth: 400, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>Contact toevoegen</h2>

      {/* Scenario banners */}
      {status === "exists-unchanged" && (
        <Text style={{ marginBottom: 16 }}>
          “{info.name}” is already a contact—no new information.
        </Text>
      )}

      {status === "exists-changed" && (
        <div style={{ marginBottom: 16 }}>
          <Text>“{info.name}” is already a contact. Update their information?</Text>
          <Button onClick={updateExisting} disabled={saving} style={{ marginRight: 8 }}>
            {saving ? "Updating…" : "Yes, update"}
          </Button>
          <Button onClick={() => setStatus("idle")}>No, thanks</Button>
        </div>
      )}

      {status === "not-found" && (
        <Text style={{ marginBottom: 16 }}>
          “{info.name}” is not yet a contact. Adjust fields and click “Add.”
        </Text>
      )}

      {/* Success / Error messages */}
      {statusMessage && (
        <Text
          style={{
            marginBottom: 16,
            color: statusMessage.type === "success" ? "green" : "red",
          }}
        >
          {statusMessage.text}
        </Text>
      )}

      {/* Form fields */}
      <label>
        Naam
        <br />
        <Input value={info.name} onChange={updateField("name")} />
      </label>
      <br />
      <br />

      <label>
        E-mail
        <br />
        <Input value={info.email} onChange={updateField("email")} />
      </label>
      <br />
      <br />

      <label>
        Telefoon
        <br />
        <Input value={info.phone} onChange={updateField("phone")} />
      </label>
      <br />
      <br />

      <label>
        Organisatie
        <br />
        <Input value={info.organization} onChange={updateField("organization")} />
      </label>
      <br />
      <br />

      <label>
        Postcode
        <br />
        <Input value={info.postcode} onChange={updateField("postcode")} />
      </label>
      <br />
      <br />

      {/* Add button: only enabled if new or after dismissing update prompt */}
      <Button
        appearance="primary"
        onClick={onSave}
        disabled={saving || status === "exists-unchanged"}
      >
        {saving ? "Opslaan…" : "Add to contacts"}
      </Button>
    </div>
  );
}

