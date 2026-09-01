# Dineritos — puesta en marcha

App de finanzas personales que **lee y escribe directamente en tu Excel de OneDrive**.
No hay servidor propio ni intermediarios: solo tu navegador y Microsoft.

Son cuatro pasos y se hacen una vez. Cuenta unos 20 minutos la primera vez.

---

## 1. Alojar la app

Necesita estar en una dirección `https://`, porque el login de Microsoft no funciona de otra
forma. Lo más fácil es GitHub Pages, que es gratis.

1. Entra en [github.com](https://github.com) (crea cuenta si no tienes).
2. **New repository** → nombre `dineritos` → **Public** → *Create repository*.
   > Público solo significa que se ve el código, que no tiene ningún secreto dentro.
   > **Tus datos no están aquí**: viven en tu navegador y en tu OneDrive.
   > Si prefieres repositorio privado, usa Cloudflare Pages en vez de GitHub Pages.
3. **Add file → Upload files** y sube los seis archivos de esta carpeta:
   `index.html`, `app.js`, `manifest.webmanifest`, `sw.js`, `icon-192.png`, `icon-512.png`.
   (El `LEEME.md` da igual.) → *Commit changes*.
4. **Settings → Pages** → *Source*: **Deploy from a branch** → *Branch*: `main` / `(root)` → **Save**.
5. Espera un par de minutos. Tu dirección será:

   ```
   https://TUUSUARIO.github.io/dineritos/
   ```

   **Apúntala con la barra final incluida.** La necesitas exacta en el paso 2.

---

## 2. Registrar la app en Microsoft

Esto es lo que le da permiso para tocar tu OneDrive.

1. Ve a [portal.azure.com](https://portal.azure.com) y entra con **tu cuenta personal de
   Microsoft** (la misma donde vas a poner el Excel).
2. Busca arriba **Microsoft Entra ID** → en el menú izquierdo, **Registros de aplicaciones**
   (*App registrations*) → **Nuevo registro**.
3. Rellena:
   - **Nombre**: `Dineritos`
   - **Tipos de cuenta compatibles**: **Solo cuentas personales de Microsoft**
   - **URI de redirección**: elige la plataforma **Aplicación de una sola página (SPA)** y pega
     tu dirección del paso 1, tal cual, con la barra final:
     `https://TUUSUARIO.github.io/dineritos/`
4. **Registrar**.
5. En la pantalla *Información general*, copia el **Id. de aplicación (cliente)**. Son 36
   caracteres con guiones, del estilo `a1b2c3d4-1111-2222-3333-abcdef123456`.
6. **No crees ningún secreto de cliente.** Las apps de navegador no lo usan (van con PKCE), y
   por eso el Id. de aplicación no es información sensible.
7. En **Permisos de API**, añade de *Microsoft Graph → Permisos delegados*: **`Files.ReadWrite`**.
   Con cuenta personal no hace falta que lo apruebe ningún administrador: te lo preguntará a ti
   la primera vez que entres.

> Si prefieres el permiso mínimo posible, puedes usar `Files.ReadWrite.AppFolder` en lugar de
> `Files.ReadWrite`, pero entonces el Excel tiene que vivir dentro de la carpeta que Microsoft
> crea para la app, y el buscador de archivos de la app no lo encontrará. Con `Files.ReadWrite`
> puedes tener el libro donde quieras.

---

## 3. Poner el Excel en tu OneDrive personal

1. Entra en [onedrive.live.com](https://onedrive.live.com) con tu cuenta personal.
2. Sube `Dineritos Pro.xlsx`.

Puedes dejar el original en la carpeta de trabajo como archivo histórico: la app no lo tocará.

---

## 4. Primer arranque

1. Abre tu dirección (`https://TUUSUARIO.github.io/dineritos/`) en el móvil.
2. Pestaña **Ajustes** → pega el **Id. de aplicación** → **Entrar con Microsoft** → acepta los
   permisos que te pida.
3. Busca `Dineritos` y pulsa **Es este** en el libro correcto. La app lo lee y ya tienes todo.
4. En el navegador del móvil: **Compartir → Añadir a pantalla de inicio**. Queda como una app
   más y arranca a pantalla completa.

Repite el paso 2-3 en el portátil si quieres usarla también allí. Cada dispositivo guarda su
propia sesión; el Excel es lo que los mantiene sincronizados.

---

## El uso diario

- **Anotar un gasto**: pestaña Mes → `＋ Anotar gasto` → concepto, categoría, importe → *Anotar*.
- **Llevarlo al Excel**: botón **Guardar** de la barra de arriba.
- Las primeras veces, usa **Ver qué se va a escribir** antes de guardar: te enseña rango por
  rango lo que va a tocar, sin enviar nada.
- **Cierra el libro en el ordenador antes de guardar**, o Excel se quejará de coautoría.

---

## Cosas que conviene saber

- **Cómo escribe**: celda a celda, con la API de Excel de Microsoft Graph. No reescribe el
  archivo, así que tus 5 gráficos, el formato condicional y las validaciones no se tocan, y los
  totales se recalculan solos en el servidor.
- **Nunca toca las columnas D ni F** de las hojas mensuales: ahí viven las fórmulas del plan y
  de la desviación. Solo escribe en B, C, E, G y H.
- **La sesión de Microsoft caduca cada 24 h** en apps de navegador. Normalmente se renueva sola
  sin que lo notes; algún día te pedirá entrar otra vez. No se pierde nada: lo que tengas sin
  guardar sigue en el dispositivo.
- **En cada hoja mensual caben 19 apuntes.** Si te pasas, la app te dice cuáles no ha escrito.
- **Sin conexión** la app abre y puedes anotar; guarda en el móvil y sube al Excel cuando vuelvas
  a tener red.
- **Si añades un mes** en la app que no tiene hoja en el Excel, te avisa: la hoja hay que crearla
  en Excel.
- **Si el Excel ha cambiado** por otro lado desde que la app lo leyó, al guardar te avisa y
  decides tú: releer o escribir encima.

## Actualizar la app

Sube de nuevo `app.js` e `index.html` al repositorio y refresca. El *service worker* se
actualiza solo en la siguiente visita.

## Copias de seguridad

- **Ajustes → Copia de seguridad (JSON)** te baja todo el estado.
- El script `sincronizar_excel.py` de la carpeta de arriba sigue funcionando como vía alternativa
  desde el ordenador, si algún día el login te da problemas.

## Si algo va mal

| Síntoma | Causa casi siempre |
|---|---|
| «La vuelta del login no cuadra» | La URI de redirección registrada no coincide *exactamente*. Repasa la barra final. |
| El login abre y vuelve sin entrar | El registro no es de tipo **SPA**, o no es «solo cuentas personales». |
| «No he podido descargar el libro» | Falta el permiso `Files.ReadWrite`, o el libro se movió. Prueba *Elegir otro libro*. |
| No aparece «Añadir a pantalla de inicio» | La dirección no es `https://`, o falta `manifest.webmanifest`. |
| Excel se queja al guardar | Tienes el libro abierto en el ordenador. Ciérralo. |
