# Control de tomas de unidad — versión web (Netlify)

Esta es la misma app que usabas dentro de Claude, convertida en un proyecto
real que puedes publicar en internet con tu propia dirección. Los datos ya
no se guardan dentro de Claude: se guardan en **Netlify Blobs**, un
almacén de datos incluido gratis con tu sitio de Netlify.

Ahora también tiene **inicio de sesión**. Nadie puede ver ni guardar tomas
sin entrar primero con correo y contraseña.

## Usuarios y contraseña temporal

Ya vienen precargados estos dos correos, con la misma contraseña temporal
para los dos:

| Correo | Contraseña temporal |
|---|---|
| expediente.ventas7@chesa.mx | `Chesa2026!` |
| expediente.ventas5@chesa.mx | `Chesa2026!` |

**La primera vez que cada quien entre, la app le va a pedir cambiar esa
contraseña por una propia** (mínimo 8 caracteres) — no se puede pasar de
esa pantalla sin cambiarla. Los usuarios se crean solos, en automático, la
primera vez que alguien intenta iniciar sesión en el sitio ya publicado
(no tienes que hacer nada extra para crearlos).

Si en algún momento quieres cambiar la contraseña temporal antes de
publicar (por ejemplo, si vas a mandarla por un medio inseguro), edita la
constante `TEMP_PASSWORD` en `netlify/functions/auth.js` antes de subir el
proyecto.

## Opción A — sin GitHub, arrastrando la carpeta (la más rápida)

1. Entra a [app.netlify.com](https://app.netlify.com) y da clic en **"Add new site" → "Deploy manually"**.
2. En tu computadora, dentro de esta carpeta, corre:
   ```
   npm install
   npm run build
   ```
   Esto crea una carpeta `dist/`.
3. Arrastra la carpeta `dist/` a la ventana de Netlify donde dice "Drag and drop your site output folder here".
4. Netlify te da una URL tipo `algo-al-azar.netlify.app`. Ya funciona — pero
   **antes de usarla, sigue el paso "Activar las funciones" más abajo**,
   porque el deploy manual por arrastre a veces no sube la carpeta
   `netlify/functions` sola.

   Si al usar la app ves el error "no se pudo guardar", es justo por eso:
   ve a tu sitio en Netlify → pestaña **"Functions"** y confirma que
   aparece la función `storage`. Si no aparece, usa la Opción B.

## Opción B — con GitHub (recomendada, se actualiza sola)

1. Crea una cuenta en [github.com](https://github.com) si no tienes.
2. Crea un repositorio nuevo (puede ser privado) y sube esta carpeta completa:
   ```
   git init
   git add .
   git commit -m "Control de tomas de unidad"
   git branch -M main
   git remote add origin https://github.com/TU-USUARIO/TU-REPO.git
   git push -u origin main
   ```
3. En Netlify: **"Add new site" → "Import an existing project"** → elige
   GitHub → selecciona tu repositorio.
4. Netlify detecta solo la configuración (todo está en `netlify.toml`):
   - Build command: `npm run build`
   - Publish directory: `dist`
   - Functions directory: `netlify/functions`
5. Dale "Deploy site". Con este método, la función `storage` sí se activa
   automáticamente porque Netlify construye el proyecto completo desde el
   código fuente (no solo la carpeta `dist`).
6. Cada vez que subas un cambio a GitHub (`git push`), Netlify vuelve a
   publicar solo.

## Probarla en tu computadora antes de publicar (opcional)

```
npm install -g netlify-cli
npm install
netlify dev
```

Esto levanta la app Y las funciones juntas en `http://localhost:8888`,
exactamente como se comportará ya publicada.

## Compartir con tu compañero(a)

Nada más mándale la URL que te dio Netlify (o tu dominio si conectas uno
propio). Como los datos viven en Netlify Blobs y no en el navegador de
cada quien, los dos van a ver la misma información — igual que cuando la
usaban dentro de Claude.

## Estructura del proyecto

- `src/App.jsx` — la app completa (dashboard, registro, formularios).
- `src/Auth.jsx` — pantallas de login y de cambio de contraseña obligatorio.
- `src/storage.js` — imita el `window.storage` de Claude, pero habla con
  la función de datos, mandando el token de sesión en cada petición.
- `netlify/functions/auth.js` — login, cambio de contraseña, verificación
  de sesión y cierre de sesión. Aquí se crean los dos usuarios la primera
  vez que se usan.
- `netlify/functions/storage.js` — guarda y lee las tomas en Netlify
  Blobs; ahora exige una sesión válida para responder.
- `netlify/functions/_lib/session.js` — código compartido entre las dos
  funciones de arriba para validar el token de sesión.
- `netlify.toml` — le dice a Netlify cómo construir y dónde están las
  funciones.

## Importante sobre este login

Este login es sencillo pero real: las contraseñas se guardan encriptadas
(no en texto plano) y cada sesión expira sola a los 30 días. No tiene
recuperación de contraseña por correo ni roles de permisos — si alguien
olvida su contraseña nueva, hay que restablecerla a mano borrando su
registro en el almacén `tomas-auth-users` desde el panel de Netlify
Blobs, para que se vuelva a crear con la temporal.

**Recuerda: la Opción A (arrastrar la carpeta) a veces no sube las
funciones.** Si después de publicar el login no funciona, revisa en tu
sitio de Netlify → pestaña "Functions" que aparezcan `auth` y `storage`.
Si no aparecen, usa la Opción B (conectada a GitHub).
