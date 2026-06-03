// netlify/functions/save-data.mjs
import { getStore } from "@netlify/blobs";

export default async (request) => {
  const expectedToken = process.env.APP_API_TOKEN;
  const receivedToken = request.headers.get("x-api-token");

  if (!expectedToken) {
    return new Response(JSON.stringify({ error: "Server misconfigured: APP_API_TOKEN not set" }), { status: 500 });
  }
  if (!receivedToken || receivedToken !== expectedToken) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), { status: 401 });
  }

  if (request.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405 });
  }

  try {
    const data = await request.json();
    if (!Array.isArray(data)) {
      return new Response(JSON.stringify({ error: "Invalid payload: expected array" }), { status: 400 });
    }
    const store = getStore("nhl-assessment");
    await store.setJSON("apps", data);
    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
};

export const config = { path: "/api/save-data" };
