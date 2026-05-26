import { getStore } from "@netlify/blobs";

export default async (req, context) => {
  try {
    const body  = await req.json();
    const store = getStore("nhl-assessment");
    await store.setJSON("apps", body);
    return Response.json({ ok: true });
  } catch (err) {
    console.error("save-data error:", err);
    return Response.json({ ok: false, error: err.message }, { status: 500 });
  }
};

export const config = { path: "/api/save-data" };
