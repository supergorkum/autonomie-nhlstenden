// netlify/functions/scheduled-backup.mjs
// Automatische backup om 12:00 en 18:00 elke dag (UTC+2 = 10:00 en 16:00 UTC)
import { getStore } from "@netlify/blobs";

export default async () => {
  try {
    const store = getStore("nhl-assessment");

    // Haal huidige data op
    const data = await store.get("apps", { type: "json" });
    if (!data || !Array.isArray(data) || data.length === 0) {
      console.log("Geen data om te backuppen.");
      return new Response("No data", { status: 200 });
    }

    // Sla backup op met timestamp als key: backup:2026-06-09T12:00
    const now = new Date();
    const key = `backup:${now.toISOString().slice(0, 16).replace("T", "T")}`; // bijv. backup:2026-06-09T10:00
    await store.setJSON(key, {
      timestamp: now.toISOString(),
      apps: data,
      count: data.length,
    });
    console.log(`Backup opgeslagen: ${key} (${data.length} apps)`);

    // Verwijder backups ouder dan 7 dagen
    const cutoff = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const { blobs } = await store.list({ prefix: "backup:" });
    let deleted = 0;
    for (const blob of blobs) {
      const keyDate = blob.key.replace("backup:", "");
      const blobDate = new Date(keyDate);
      if (!isNaN(blobDate) && blobDate < cutoff) {
        await store.delete(blob.key);
        deleted++;
        console.log(`Oude backup verwijderd: ${blob.key}`);
      }
    }
    if (deleted > 0) console.log(`${deleted} oude backup(s) verwijderd.`);

    return new Response(JSON.stringify({ ok: true, key, count: data.length, deleted }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error("Backup fout:", err);
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
};

// Elke dag om 08:00, 11:00 en 16:00 UTC = 10:00, 13:00 en 18:00 NL zomertijd
export const config = {
  schedule: "0 8,11,16 * * *",
};
