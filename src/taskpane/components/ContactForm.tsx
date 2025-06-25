// File: src/taskpane/components/ContactForm.tsx

import React, { useEffect, useState } from "react";
import { Button, Input, Text } from "@fluentui/react-components";
import { createContact } from "../graph";
import { extractSignatureBlock, parseSignature } from "./signature";
import { getGraphToken } from "../authConfig";

export function ContactForm() {
  const [info, setInfo] = useState({ name: "", email: "", phone: "", organization: "", postcode: "" });
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);

  useEffect(() => {
    Office.onReady(() => {
      const item = Office.context.mailbox.item;
      const senderName = item.from?.displayName || "";
      const senderEmail = item.from?.emailAddress || "";
      const rawSeed = senderEmail.split("@")[1]?.split(".")[0] || "";
      const organization = rawSeed
        ? rawSeed.charAt(0).toUpperCase() + rawSeed.slice(1).toLowerCase()
        : "";

      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const sig = extractSignatureBlock(result.value, senderName);
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

  const updateField = (key: keyof typeof info) => (e: React.ChangeEvent<HTMLInputElement>) =>
    setInfo({ ...info, [key]: e.target.value });

  const onSave = async () => {
    setSaving(true);
    setStatusMessage(null);
    try {
      const token = await getGraphToken();
      await createContact(token, info);
      setStatusMessage({ type: "success", text: "Contact opgeslagen!" });
    } catch (e: any) {
      console.error(e);
      setStatusMessage({ type: "error", text: "Opslaan mislukt: " + e.message });
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return <Text>Loading signature…</Text>;
  }

  return (
    <div style={{ padding: 20, maxWidth: 400, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>Contact toevoegen</h2>

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

      <label>
        Naam<br />
        <Input value={info.name} onChange={updateField("name")} />
      </label>
      <br />
      <br />

      <label>
        E-mail<br />
        <Input value={info.email} onChange={updateField("email")} />
      </label>
      <br />
      <br />

      <label>
        Telefoon<br />
        <Input value={info.phone} onChange={updateField("phone")} />
      </label>
      <br />
      <br />

      <label>
        Organisatie<br />
        <Input value={info.organization} onChange={updateField("organization")} />
      </label>
      <br />
      <br />

      <label>
        Postcode<br />
        <Input value={info.postcode} onChange={updateField("postcode")} />
      </label>
      <br />
      <br />

      <Button appearance="primary" onClick={onSave} disabled={saving}>
        {saving ? "Opslaan…" : "Add to contacts"}
      </Button>
    </div>
  );
}
