// ─────────────────────────────────────────────────────────────
//  OPCIÓN A (recomendada) · GOOGLE SHEETS como base de datos
//  Pega aquí la URL /exec de la aplicación web de Apps Script
//  (ver README → "Conectar Google Sheets").
// ─────────────────────────────────────────────────────────────
export const SHEETS_API_URL = "https://script.google.com/macros/s/AKfycbwqSchmsiffcOySI0NG3akcD3Di6eYpW02CXc8AM-eRlWhsB0AJvWcNcMclbpdYrDkO/exec";
// Ejemplo: "https://script.google.com/macros/s/AKfycb.../exec"

// ─────────────────────────────────────────────────────────────
//  OPCIÓN B · Configuración Firebase (sólo si NO usas Google Sheets) · Estadística Operacional de Naves HTP
//  1. Crea un proyecto en https://console.firebase.google.com
//  2. Agrega una "App web" (</>) y copia aquí el objeto firebaseConfig
//  3. Activa Firestore Database y Authentication → Google
//  Mientras esto quede en null, la página funciona en MODO LOCAL
//  (los datos se guardan sólo en este navegador).
// ─────────────────────────────────────────────────────────────
export const firebaseConfig = null;
/* Ejemplo:
export const firebaseConfig = {
  apiKey: "AIza...",
  authDomain: "htp-estadisticas.firebaseapp.com",
  projectId: "htp-estadisticas",
  storageBucket: "htp-estadisticas.appspot.com",
  messagingSenderId: "000000000000",
  appId: "1:000000000000:web:xxxxxxxx"
};
*/

// Correos autorizados para ingresar / editar servicios (debe coincidir con firestore.rules)
// Si la lista queda vacía, cualquier usuario que inicie sesión con Google puede editar.
export const EDITORES = ["jvilla@htpchile.cl"];
