// netlify/functions/load-data.js
const { getStore } = require("@netlify/blobs");

exports.handler = async (event) => {
  // ── Server-side tokencheck ────────────────────────────────
  const expectedToken = process.env.APP_API_TOKEN;
  const receivedToken = event.headers["x-api-token"];

  if (!expectedToken) {
    return { statusCode: 500, body: JSON.stringify({ error: "Server misconfigured: APP_API_TOKEN not set" }) };
  }
  if (!receivedToken || receivedToken !== expectedToken) {
    return { statusCode: 401, body: JSON.stringify({ error: "Unauthorized" }) };
  }

  // ── Alleen GET accepteren ─────────────────────────────────
  if (event.httpMethod !== "GET") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const store = getStore("nhl-sov-data");
    const data  = await store.get("apps", { type: "json" });
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data ?? []),
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
