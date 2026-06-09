// netlify/functions/save-export-timestamp.mjs
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
    const { timestamp, type } = await request.json();
    const store = getStore("nhl-assessment");
    await store.setJSON("last-export", { timestamp, type });
    return new Response(JSON.stringify({ ok: true }), { status: 200 });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
};

export const config = { path: "/api/save-export-timestamp" };
