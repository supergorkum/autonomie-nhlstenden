// netlify/functions/list-backups.mjs
import { getStore } from "@netlify/blobs";

export default async (request) => {
  const expectedToken = process.env.APP_API_TOKEN;
  const receivedToken = request.headers.get("x-api-token");

  if (!expectedToken || !receivedToken || receivedToken !== expectedToken) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), { status: 401 });
  }

  try {
    const store = getStore("nhl-assessment");
    const { blobs } = await store.list({ prefix: "backup:" });

    // Sorteer op datum, nieuwste eerst
    const backups = blobs
      .map(b => ({
        key: b.key,
        timestamp: b.key.replace("backup:", ""),
      }))
      .sort((a, b) => b.timestamp.localeCompare(a.timestamp))
      .slice(0, 14); // max 14 tonen (7 dagen × 2 per dag)

    return new Response(JSON.stringify(backups), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
};

export const config = { path: "/api/list-backups" };
