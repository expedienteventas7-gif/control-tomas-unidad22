import { getStore } from "@netlify/blobs";
import { validateToken, corsHeaders, getBearerToken } from "./_lib/session.js";

let _store = null;
function getUnidadStore() {
  if (!_store) {
    const siteID = process.env.BLOBS_SITE_ID;
    const token = process.env.BLOBS_TOKEN;
    const config = siteID && token
      ? { name: "tomas-unidad", siteID, token, consistency: "strong" }
      : { name: "tomas-unidad", consistency: "strong" };
    _store = getStore(config);
  }
  return _store;
}

function json(statusCode, body) {
  return { statusCode, headers: corsHeaders(), body: JSON.stringify(body) };
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: corsHeaders(), body: "" };

  try {
    const store = getUnidadStore();
    const token = getBearerToken(event);
    const session = await validateToken(token);
    if (!session) return json(401, { error: "no autenticado" });

    const params = event.queryStringParameters || {};
    const key = params.key;

    if (event.httpMethod === "GET") {
      if (params.list) {
        const prefix = params.prefix || "";
        const { blobs } = await store.list({ prefix });
        return json(200, { keys: blobs.map((b) => b.key) });
      }
      if (!key) return json(400, { error: "falta 'key'" });
      const raw = await store.get(key);
      if (raw === null) return json(404, { error: "no existe" });
      return json(200, { key, value: raw });
    }

    if (event.httpMethod === "POST") {
      const body = JSON.parse(event.body || "{}");
      if (!body.key) return json(400, { error: "falta 'key'" });
      await store.set(body.key, body.value);
      return json(200, { key: body.key, value: body.value });
    }

    if (event.httpMethod === "DELETE") {
      if (!key) return json(400, { error: "falta 'key'" });
      await store.delete(key);
      return json(200, { key, deleted: true });
    }

    return json(405, { error: "método no soportado" });
  } catch (e) {
    return json(500, { error: String((e && e.message) || e) });
  }
}
