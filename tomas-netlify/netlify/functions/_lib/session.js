import { getStore } from "@netlify/blobs";

let _sessionsStore = null;
let _usersStore = null;

// En deploys manuales (arrastrar carpeta/zip), Netlify a veces no inyecta el
// contexto automático que Blobs necesita. Como respaldo, si existen estas
// variables de entorno (configuradas en Site settings > Environment variables),
// las usamos explícitamente.
function blobsConfig(name) {
  const siteID = process.env.BLOBS_SITE_ID;
  const token = process.env.BLOBS_TOKEN;
  if (siteID && token) {
    return { name, siteID, token, consistency: "strong" };
  }
  return { name, consistency: "strong" };
}

function getSessionsStore() {
  if (!_sessionsStore) _sessionsStore = getStore(blobsConfig("tomas-auth-sessions"));
  return _sessionsStore;
}

function getUsersStore() {
  if (!_usersStore) _usersStore = getStore(blobsConfig("tomas-auth-users"));
  return _usersStore;
}

export const sessionsStore = {
  get(...args) { return getSessionsStore().get(...args); },
  set(...args) { return getSessionsStore().set(...args); },
  delete(...args) { return getSessionsStore().delete(...args); },
};

export const usersStore = {
  get(...args) { return getUsersStore().get(...args); },
  set(...args) { return getUsersStore().set(...args); },
  delete(...args) { return getUsersStore().delete(...args); },
};

const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 días

export async function validateToken(token) {
  if (!token) return null;
  const raw = await sessionsStore.get(token);
  if (raw === null) return null;
  const session = JSON.parse(raw);
  if (Date.now() - session.createdAt > SESSION_TTL_MS) {
    await sessionsStore.delete(token);
    return null;
  }
  return session;
}

export function corsHeaders() {
  return {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
  };
}

export function getBearerToken(event) {
  const raw = (event.headers && (event.headers.authorization || event.headers.Authorization)) || "";
  return raw.replace(/^Bearer\s+/i, "").trim();
}
