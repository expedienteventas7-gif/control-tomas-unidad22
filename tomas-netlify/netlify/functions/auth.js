import bcrypt from "bcryptjs";
import crypto from "crypto";
import { sessionsStore, usersStore, validateToken, corsHeaders } from "./_lib/session.js";

// Usuarios iniciales: se crean solos la primera vez que alguien intenta
// iniciar sesión, con esta contraseña temporal. Cámbiala aquí si quieres
// otra antes de publicar, o cambia la contraseña real desde la app en el
// primer ingreso (que es obligatorio).
const TEMP_PASSWORD = "Chesa2026!";
const SEED_EMAILS = ["expediente.ventas7@chesa.mx", "expediente.ventas5@chesa.mx"];

function json(statusCode, body) {
  return { statusCode, headers: corsHeaders(), body: JSON.stringify(body) };
}

async function ensureSeedUsers() {
  for (const email of SEED_EMAILS) {
    const existing = await usersStore.get(email);
    if (existing === null) {
      const passwordHash = await bcrypt.hash(TEMP_PASSWORD, 10);
      await usersStore.set(email, JSON.stringify({ passwordHash, mustChangePassword: true }));
    }
  }
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: corsHeaders(), body: "" };
  if (event.httpMethod !== "POST") return json(405, { error: "método no soportado" });

  let body;
  try {
    body = JSON.parse(event.body || "{}");
  } catch (e) {
    return json(400, { error: "cuerpo de la petición inválido" });
  }

  const action = body.action;

  try {
    await ensureSeedUsers();

    if (action === "login") {
      const email = (body.email || "").trim().toLowerCase();
      const password = body.password || "";
      if (!email || !password) return json(400, { error: "correo y contraseña son obligatorios" });

      const raw = await usersStore.get(email);
      if (raw === null) return json(401, { error: "usuario o contraseña incorrectos" });
      const user = JSON.parse(raw);
      const okPass = await bcrypt.compare(password, user.passwordHash);
      if (!okPass) return json(401, { error: "usuario o contraseña incorrectos" });

      const token = crypto.randomUUID();
      await sessionsStore.set(token, JSON.stringify({ email, createdAt: Date.now() }));
      return json(200, { token, email, mustChangePassword: !!user.mustChangePassword });
    }

    if (action === "verify") {
      const session = await validateToken(body.token);
      if (!session) return json(401, { error: "sesión inválida o expirada" });
      const raw = await usersStore.get(session.email);
      const user = raw ? JSON.parse(raw) : {};
      return json(200, { email: session.email, mustChangePassword: !!user.mustChangePassword });
    }

    if (action === "change-password") {
      const session = await validateToken(body.token);
      if (!session) return json(401, { error: "sesión inválida o expirada" });
      const currentPassword = body.currentPassword || "";
      const newPassword = body.newPassword || "";
      if (newPassword.length < 8) return json(400, { error: "la nueva contraseña debe tener al menos 8 caracteres" });

      const raw = await usersStore.get(session.email);
      if (raw === null) return json(404, { error: "usuario no encontrado" });
      const user = JSON.parse(raw);
      const okPass = await bcrypt.compare(currentPassword, user.passwordHash);
      if (!okPass) return json(401, { error: "la contraseña actual no es correcta" });

      const passwordHash = await bcrypt.hash(newPassword, 10);
      await usersStore.set(session.email, JSON.stringify({ passwordHash, mustChangePassword: false }));
      return json(200, { ok: true });
    }

    if (action === "logout") {
      if (body.token) await sessionsStore.delete(body.token);
      return json(200, { ok: true });
    }

    return json(400, { error: "acción no reconocida" });
  } catch (err) {
    // En vez de dejar que la función truene con un 502 sin explicación,
    // devolvemos el mensaje de error real para poder diagnosticarlo desde el navegador.
    return json(500, { error: "Error interno: " + (err && err.message ? err.message : String(err)) });
  }
}
