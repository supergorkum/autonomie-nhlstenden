// netlify/functions/save-data.js
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

  // ── Alleen POST accepteren ────────────────────────────────
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const data = JSON.parse(event.body);
    if (!Array.isArray(data)) {
      return { statusCode: 400, body: JSON.stringify({ error: "Invalid payload: expected array" }) };
    }
    const store = getStore("nhl-sov-data");
    await store.setJSON("apps", data);
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: true }),
    };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
