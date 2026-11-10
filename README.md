# Estadística Operacional de Naves · HTP

Dashboard web de Huachipato Terminal Portuario para registrar servicios de naves y analizar transferencia, rates, uso de muelle y detenciones.

## Qué incluye
- **Resumen ejecutivo**: toneladas por mes (descarga/embarque + N° naves), mix por proceso, top productos, naves por sitio/operador, comparativo interanual acumulado.
- **Productividad y rates**: rate alcanzado vs contratado por nave, cumplimiento comercial mensual, tons/hr por producto, ranking de naves más productivas y bajo contrato.
- **Muelle y detenciones**: horas de muelle por sitio, detenciones por causa, distribución del tiempo en muelle por nave.
- **Clientes y productos**: participación por cliente, tipo de carga, matriz cliente × producto.
- **Registro de servicios**: tabla con búsqueda y orden; botón **＋ Ingresar servicio** con formulario completo y cálculo automático de uso de muelle, tiempo de operación, rates, % cumplimiento y tons/hr.
- Importar el Excel `Estadisticas_naves` (hoja *Estadística Operacional Naves*) y exportar a Excel.

## Conectar Google Sheets (base de datos)
1. Abre **sheets.new** y en **Archivo → Importar → Subir** carga `Plantilla_Google_Sheets_Naves.xlsx` (opción *Reemplazar hoja de cálculo*). La pestaña debe llamarse **Naves** y la columna A **ID**.
2. En la planilla: **Extensiones → Apps Script**. Borra lo que aparece, pega `google-apps-script/Code.gs` y **cambia `CLAVE_EDICION`** por tu clave.
3. **Implementar → Nueva implementación → ⚙ Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
   → Autoriza los permisos y copia la URL que termina en `/exec`.
4. Pega esa URL en `config.js` → `SHEETS_API_URL`.
5. Listo: la página lee la planilla (se actualiza sola cada 2 min o con **↻ Actualizar**). El formulario **＋ Ingresar servicio** escribe en la planilla y pide la clave de edición la primera vez.

Pueden seguir llenando datos directo en la planilla: respeta los encabezados de la fila 1 y deja una ID única en la columna A (cualquier texto). Las columnas de indicadores las recalcula la página.

> Si editas `Code.gs`, vuelve a **Implementar → Gestionar implementaciones → Editar → Nueva versión** para que el cambio se publique.

## Alternativa: Firebase (sólo si no usas Google Sheets)
1. https://console.firebase.google.com → **Agregar proyecto** (ej. `htp-estadisticas`).
2. **Compilación → Firestore Database → Crear base de datos** (modo producción, región `southamerica-east1`).
3. **Compilación → Authentication → Comenzar → Google** (habilitar).
4. **Configuración del proyecto → Tus apps → Web (</>)** → copia el objeto `firebaseConfig` en `config.js`.
5. **Firestore → Reglas** → pega el contenido de `firestore.rules` y publica. Agrega ahí (y en `EDITORES` de `config.js`) los correos que pueden ingresar servicios.
6. **Authentication → Configuración → Dominios autorizados** → agrega `jvillapena123-cell.github.io`.
7. Abre la página, inicia sesión y presiona **Cargar los servicios del Excel inicial** (sólo la primera vez).

Sin `SHEETS_API_URL` ni Firebase, la página funciona en **modo local** (guarda en el navegador).

## Publicación
- **GitHub Pages**: Settings → Pages → Branch `main` / root.
- **Firebase Hosting** (opcional): `npm i -g firebase-tools && firebase login && firebase init hosting && firebase deploy`.
