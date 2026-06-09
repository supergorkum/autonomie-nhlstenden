// netlify/functions/restore-backup.mjs
import { getStore } from "@netlify/blobs";

export default async (request) => {
  const expectedToken = process.env.APP_API_TOKEN;
  const receivedToken = request.headers.get("x-api-token");

  if (!expectedToken || !receivedToken || receivedToken !== expectedToken) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), { status: 401 });
  }

  if (request.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405 });
  }

  try {
    const { key, previewOnly } = await request.json();
    if (!key || !key.startsWith("backup:")) {
      return new Response(JSON.stringify({ error: "Ongeldige backup key" }), { status: 400 });
    }

    const store = getStore("nhl-assessment");
    const backup = await store.get(key, { type: "json" });
    if (!backup || !Array.isArray(backup.apps)) {
      return new Response(JSON.stringify({ error: "Backup niet gevonden" }), { status: 404 });
    }

    // previewOnly: stuur alleen de data terug zonder te herstellen
    if (previewOnly) {
      return new Response(JSON.stringify({
        apps: backup.apps,
        count: backup.apps.length,
        timestamp: backup.timestamp,
      }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }

    // Herstel de apps
    await store.setJSON("apps", backup.apps);

    return new Response(JSON.stringify({
      ok: true,
      count: backup.apps.length,
      timestamp: backup.timestamp,
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
};

export const config = { path: "/api/restore-backup" };
