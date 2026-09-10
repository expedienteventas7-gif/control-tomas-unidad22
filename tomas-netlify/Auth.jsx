import React, { useState, useEffect, useCallback } from "react";

const C = {
  bg: "#EEF0EE",
  surface: "#FFFFFF",
  ink: "#1B231F",
  inkMuted: "#5B6660",
  line: "#DBDFD9",
  primary: "#2B5C4E",
  danger: "#A23E3E",
  dangerSoft: "#F3E1DF",
};

const FONT = "'Space Grotesk', sans-serif";
const FONT_BODY = "'IBM Plex Sans', sans-serif";

const TOKEN_KEY = "tomas-auth-token";
const EMAIL_KEY = "tomas-auth-email";

async function callAuth(payload) {
  const res = await fetch("/api/auth", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Error de autenticación");
  return data;
}

const inputStyle = {
  fontFamily: FONT_BODY, fontSize: 14, padding: "10px 12px", borderRadius: 4,
  border: `1px solid ${C.line}`, outline: "none", width: "100%", boxSizing: "border-box",
};

function Card({ children }) {
  return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: FONT_BODY }}>
      <div style={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, padding: 32, width: "100%", maxWidth: 380, display: "flex", flexDirection: "column", gap: 16 }}>
        {children}
      </div>
    </div>
  );
}

function LoginScreen({ onSuccess }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const data = await callAuth({ action: "login", email, password });
      localStorage.setItem(TOKEN_KEY, data.token);
      localStorage.setItem(EMAIL_KEY, data.email);
      onSuccess(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card>
      <div>
        <div style={{ fontFamily: FONT, fontSize: 20, color: C.ink }}>Control de tomas de unidad</div>
        <div style={{ fontFamily: FONT_BODY, fontSize: 13, color: C.inkMuted }}>Inicia sesión para continuar</div>
      </div>
      <form onSubmit={submit} style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, color: C.inkMuted, fontFamily: FONT }}>Correo</span>
          <input style={inputStyle} type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoFocus autoComplete="username" />
        </label>
        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, color: C.inkMuted, fontFamily: FONT }}>Contraseña</span>
          <input style={inputStyle} type="password" required value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" />
        </label>
        {error && <div style={{ background: C.dangerSoft, color: C.danger, fontSize: 13, padding: "8px 10px", borderRadius: 4 }}>{error}</div>}
        <button type="submit" disabled={loading} style={{
          fontFamily: FONT, fontSize: 14, padding: "10px 14px", borderRadius: 4, cursor: "pointer",
          background: C.primary, color: "#fff", border: "none", opacity: loading ? 0.6 : 1,
        }}>{loading ? "Entrando…" : "Entrar"}</button>
      </form>
    </Card>
  );
}

function ChangePasswordScreen({ email, onDone }) {
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    if (newPassword !== confirm) { setError("Las contraseñas nuevas no coinciden"); return; }
    if (newPassword.length < 8) { setError("La nueva contraseña debe tener al menos 8 caracteres"); return; }
    setLoading(true);
    try {
      const token = localStorage.getItem(TOKEN_KEY);
      await callAuth({ action: "change-password", token, currentPassword, newPassword });
      onDone();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card>
      <div>
        <div style={{ fontFamily: FONT, fontSize: 20, color: C.ink }}>Crea tu contraseña</div>
        <div style={{ fontFamily: FONT_BODY, fontSize: 13, color: C.inkMuted }}>{email} — es tu primer ingreso, cambia la contraseña temporal por una tuya.</div>
      </div>
      <form onSubmit={submit} style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, color: C.inkMuted, fontFamily: FONT }}>Contraseña temporal</span>
          <input style={inputStyle} type="password" required value={currentPassword} onChange={(e) => setCurrentPassword(e.target.value)} autoFocus autoComplete="current-password" />
        </label>
        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, color: C.inkMuted, fontFamily: FONT }}>Nueva contraseña (mín. 8 caracteres)</span>
          <input style={inputStyle} type="password" required value={newPassword} onChange={(e) => setNewPassword(e.target.value)} autoComplete="new-password" />
        </label>
        <label style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <span style={{ fontSize: 12, color: C.inkMuted, fontFamily: FONT }}>Confirma la nueva contraseña</span>
          <input style={inputStyle} type="password" required value={confirm} onChange={(e) => setConfirm(e.target.value)} autoComplete="new-password" />
        </label>
        {error && <div style={{ background: C.dangerSoft, color: C.danger, fontSize: 13, padding: "8px 10px", borderRadius: 4 }}>{error}</div>}
        <button type="submit" disabled={loading} style={{
          fontFamily: FONT, fontSize: 14, padding: "10px 14px", borderRadius: 4, cursor: "pointer",
          background: C.primary, color: "#fff", border: "none", opacity: loading ? 0.6 : 1,
        }}>{loading ? "Guardando…" : "Guardar y entrar"}</button>
      </form>
    </Card>
  );
}

export default function Auth({ children }) {
  const [status, setStatus] = useState("loading"); // loading | login | changePassword | ready
  const [email, setEmail] = useState("");

  const checkSession = useCallback(async () => {
    const token = localStorage.getItem(TOKEN_KEY);
    if (!token) { setStatus("login"); return; }
    try {
      const data = await callAuth({ action: "verify", token });
      setEmail(data.email);
      setStatus(data.mustChangePassword ? "changePassword" : "ready");
    } catch (err) {
      localStorage.removeItem(TOKEN_KEY);
      localStorage.removeItem(EMAIL_KEY);
      setStatus("login");
    }
  }, []);

  useEffect(() => { checkSession(); }, [checkSession]);

  const handleLoginSuccess = (data) => {
    setEmail(data.email);
    setStatus(data.mustChangePassword ? "changePassword" : "ready");
  };

  const logout = async () => {
    const token = localStorage.getItem(TOKEN_KEY);
    try { await callAuth({ action: "logout", token }); } catch (e) { /* noop */ }
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(EMAIL_KEY);
    setStatus("login");
  };

  if (status === "loading") {
    return <div style={{ padding: 40, fontFamily: FONT_BODY, color: C.inkMuted }}>Cargando…</div>;
  }
  if (status === "login") {
    return <LoginScreen onSuccess={handleLoginSuccess} />;
  }
  if (status === "changePassword") {
    return <ChangePasswordScreen email={email} onDone={() => setStatus("ready")} />;
  }

  return (
    <div style={{ position: "relative" }}>
      <button
        onClick={logout}
        title="Cerrar sesión"
        style={{
          position: "fixed", top: 10, right: 12, zIndex: 100,
          fontFamily: FONT, fontSize: 11.5, padding: "6px 10px", borderRadius: 4, cursor: "pointer",
          background: C.surface, color: C.inkMuted, border: `1px solid ${C.line}`,
        }}
      >{email} · salir</button>
      {children}
    </div>
  );
}
