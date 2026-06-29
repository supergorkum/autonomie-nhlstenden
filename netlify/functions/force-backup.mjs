// netlify/functions/force-backup.mjs
// Handmatig een backup forceren vanuit de Beheeromgeving
import { getStore } from "@netlify/blobs";

export default async (request) => {
  const expectedToken = process.env.APP_API_TOKEN;
  const receivedToken = request.headers.get("x-api-token");

  if (!expectedToken || !receivedToken || receivedToken !== expectedToken) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), { status: 401 });
  }

  try {
    const store = getStore("nhl-assessment");
    const data = await store.get("apps", { type: "json" });
    if (!data || !Array.isArray(data) || data.length === 0) {
      return new Response(JSON.stringify({ error: "Geen data om te backuppen" }), { status: 400 });
    }

    const now = new Date();
    const key = `backup:${now.toISOString().slice(0, 16)}`;
    await store.setJSON(key, {
      timestamp: now.toISOString(),
      apps: data,
      count: data.length,
    });

    return new Response(JSON.stringify({ ok: true, key, count: data.length }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
};

export const config = { path: "/api/force-backup" };
