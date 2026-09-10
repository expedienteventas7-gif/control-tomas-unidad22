// Imita exactamente la API window.storage que usa la app (get/set/delete/list),
// pero en vez de guardar dentro de Claude, guarda en Netlify Blobs a través
// de la función /netlify/functions/storage. Cada petición manda el token de
// sesión del usuario que inició sesión (ver Auth.jsx).

const API = "/api/storage";
const TOKEN_KEY = "tomas-auth-token";

function authHeaders() {
  const token = localStorage.getItem(TOKEN_KEY);
  return token ? { Authorization: `Bearer ${token}` } : {};
}

async function parseOrThrow(res) {
  if (res.status === 401) {
    // La sesión ya no es válida: se cierra y se manda a la pantalla de login.
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem("tomas-auth-email");
    window.location.reload();
    throw new Error("unauthenticated");
  }
  if (res.status === 404) {
    throw new Error("not_found");
  }
  if (!res.ok) {
    throw new Error("storage_error_" + res.status);
  }
  return res.json();
}

export const storage = {
  async get(key) {
    const res = await fetch(`${API}?key=${encodeURIComponent(key)}`, { headers: authHeaders() });
    return parseOrThrow(res);
  },
  async set(key, value) {
    const res = await fetch(API, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ key, value }),
    });
    return parseOrThrow(res);
  },
  async delete(key) {
    const res = await fetch(`${API}?key=${encodeURIComponent(key)}`, { method: "DELETE", headers: authHeaders() });
    return parseOrThrow(res);
  },
  async list(prefix) {
    const res = await fetch(`${API}?list=1&prefix=${encodeURIComponent(prefix || "")}`, { headers: authHeaders() });
    return parseOrThrow(res);
  },
};

if (typeof window !== "undefined") {
  window.storage = storage;
}
