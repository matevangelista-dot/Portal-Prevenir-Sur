# Protocolo de Iluminación

App web para cargar desde el celular, en el lugar, las mediciones de iluminación según la **Res. SRT 84/12** y el **Dec. 351/79 Anexo IV**, y descargar el protocolo en Excel.

- Grilla por local con índice K = (L·A)/[Hm·(L+A)] y puntos mínimos (x+2)²
- Teclado numérico en pantalla para cargar lecturas mientras se mide
- E media, E mínima, uniformidad (E mín ≥ E med/2) y cumplimiento en vivo
- Varias empresas y mediciones por usuario, con login
- Funciona sin señal: guarda en el teléfono y sincroniza al recuperar conexión
- Importa el Excel del protocolo del año anterior (empresa, puntos y valores para comparar)
- Exporta el protocolo en Excel (formato tradicional o Res. 84/12 de 3 hojas)

Es un sitio estático (HTML + JS, sin compilación). Los datos se guardan en **Supabase** (plan gratuito) y el sitio se publica en **Vercel**.

---

## Puesta en marcha (una sola vez)

### 1. Crear la base de datos en Supabase

1. Entrá a <https://supabase.com> → **Start your project** → ingresá con tu cuenta de GitHub.
2. **New project**: nombre `protocolo-iluminacion`, elegí una contraseña de base de datos (guardala) y la región **South America (São Paulo)** → **Create new project**. Tarda un par de minutos.
3. Menú izquierdo → **SQL Editor** → **New query** → pegá todo el contenido de [`schema.sql`](schema.sql) → **Run**. Tiene que decir *Success*.
4. Menú izquierdo → **Project Settings** → **API**. Copiá:
   - **Project URL** (algo como `https://abcd1234.supabase.co`)
   - **anon public** key (una clave larga que empieza con `eyJ...`)
5. En GitHub abrí el archivo [`config.js`](config.js) → ícono del lápiz → pegá los dos valores entre las comillas → **Commit changes**.

> La *anon key* es pública por diseño. La seguridad la dan las reglas del `schema.sql`: cada usuario solo puede leer y modificar sus propias mediciones.

### 2. Publicar en Vercel

1. Entrá a <https://vercel.com> → ingresá con GitHub.
2. **Add New… → Project** → buscá `protocolo-iluminacion` → **Import**.
3. *Framework Preset*: **Other**. No hace falta tocar nada más → **Deploy**.
4. Te queda una dirección del tipo `https://protocolo-iluminacion.vercel.app`.

### 3. Avisarle a Supabase la dirección del sitio

Para que los emails de confirmación y de "olvidé mi contraseña" lleven al sitio correcto:

Supabase → **Authentication** → **URL Configuration** → en **Site URL** poné la dirección de Vercel (`https://protocolo-iluminacion.vercel.app`) → **Save**.

Listo. Cada vez que se modifique el código en GitHub, Vercel publica la nueva versión sola.

---

## Uso

1. Abrí la dirección de Vercel en el celular → **Crear cuenta nueva** → confirmá el email → **Ingresar**.
2. En Chrome (Android): menú ⋮ → **Agregar a pantalla de inicio**. En iPhone (Safari): botón compartir → **Agregar a inicio**. Queda como una app y abre sin conexión.
3. **Importar protocolo anterior (.xlsx)** para partir de la medición del año pasado, o **Nueva medición** para empezar de cero.
4. Pestaña **Medir**: tocá el punto, cargá largo/ancho/altura de montaje, y tipeá cada lectura con el teclado en pantalla → **Guardar lectura**.
5. Pestaña **Análisis** → **Generar conclusiones automáticamente** y revisá el texto.
6. Pestaña **Exportar** → **Descargar Excel**.

Para compartirla con otros profesionales alcanza con pasarles la dirección: cada uno crea su cuenta y ve solo sus mediciones.

### Sin conexión en la planta

Si la app ya se abrió una vez con internet, funciona sin señal: cada lectura se guarda en el teléfono y arriba se ve **📱 sin conexión**. Al recuperar señal se sube sola (**☁ sincronizado**). No cierres sesión mientras haya cambios sin subir.

---

## Estructura

```
index.html            página
styles.css            estilos
app.js                lógica (cálculos, pantallas, sincronización, Excel)
config.js             URL y clave pública de Supabase
sw.js                 funcionamiento sin conexión
manifest.webmanifest  instalación como app
schema.sql            tabla y reglas de seguridad
vendor/               librerías (Supabase JS 2.45.4, xlsx-js-style 1.2.0)
```

Al publicar cambios en `app.js` o `styles.css`, subí el número de `VERSION` en `sw.js` para que los teléfonos tomen la versión nueva.

Si `config.js` queda vacío, la app funciona en **modo local** (sin login, datos solo en el dispositivo), útil para probar.

## Criterios técnicos

- **Valor requerido**: Tabla 1 del Anexo IV (Dec. 351/79) según la dificultad de la tarea visual, o un valor definido por el profesional.
- **Supera máximo**: opcional, se marca como observación cuando la E media supera el máximo del rango de la Tabla 1 (riesgo de deslumbramiento).
- **Iluminación localizada**: control de la relación localizada/general según la tabla del Anexo IV.
- **Calibración**: aviso si el certificado tiene más de 12 meses a la fecha de medición.
