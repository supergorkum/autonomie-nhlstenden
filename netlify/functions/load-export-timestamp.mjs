// netlify/functions/load-export-timestamp.mjs
import { getStore } from "@netlify/blobs";

export default async (request) => {
  const expectedToken = process.env.APP_API_TOKEN;
  const receivedToken = request.headers.get("x-api-token");
  if (!expectedToken || !receivedToken || receivedToken !== expectedToken) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), { status: 401 });
  }
  try {
    const store = getStore("nhl-assessment");
    const data = await store.get("last-export", { type: "json" });
    return new Response(JSON.stringify(data ?? null), { status: 200 });
  } catch (err) {
    return new Response(JSON.stringify(null), { status: 200 });
  }
};

export const config = { path: "/api/load-export-timestamp" };
