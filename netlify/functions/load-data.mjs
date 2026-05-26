import { getStore } from "@netlify/blobs";

export default async (req, context) => {
  try {
    const store = getStore("nhl-assessment");
    const data  = await store.get("apps", { type: "json" });
    return Response.json(data || []);
  } catch (err) {
    console.error("load-data error:", err);
    return Response.json([]);
  }
};

export const config = { path: "/api/load-data" };
