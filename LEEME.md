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

Esto es lo que le da permiso para tocar tu OneDrive. Lo único que necesitas de aquí es un
**Id. de aplicación (Client ID)**.

> ### Antes de empezar: necesitas un «directorio»
> Si entras en el portal de Azure con tu cuenta personal sin más, te sale este error:
> *«AADSTS16000: User account from identity provider 'live.com' does not exist in tenant
> 'Microsoft Services'»*.
> No es culpa tuya: a las cuentas personales Microsoft las mete en un tenant compartido que **no
> tiene directorio**, y sin directorio no se puede registrar nada. Hay que resolver eso primero,
> y tienes dos caminos.

### Camino A — Tu propio directorio gratis (recomendado)

Es el que te deja completamente independiente: la app es tuya y no depende de nadie.

1. Ve a [azure.microsoft.com/es-es/free](https://azure.microsoft.com/es-es/free) y regístrate con
   **tu cuenta personal de Microsoft**.
2. Te pedirá teléfono y una tarjeta. **Es solo verificación de identidad, no se cobra nada**: los
   registros de aplicaciones entran en el nivel gratuito de Entra ID y son gratis siempre. Si te
   incomoda, mira el camino B.
3. Al terminar tendrás tu propio directorio y serás su administrador.
4. Entra en [entra.microsoft.com](https://entra.microsoft.com) → **Aplicaciones** → **Registros de
   aplicaciones** → **Nuevo registro**.
5. Rellena:
   - **Nombre**: `Dineritos`
   - **Tipos de cuenta compatibles**: **Solo cuentas personales de Microsoft**
   - **URI de redirección**: plataforma **Aplicación de una sola página (SPA)** y tu dirección del
     paso 1, tal cual, con la barra final: `https://TUUSUARIO.github.io/dineritos/`
6. **Registrar** → copia el **Id. de aplicación (cliente)**.

### Camino B — Registrarla en el directorio de Exus (sin tarjeta)

Funciona porque **el directorio donde se registra la app no tiene que ser el de quien la usa**:
registras la app en Exus, pero luego entras en ella con tu cuenta personal y lee tu OneDrive
personal. Los datos no pasan por Exus en ningún momento.

1. Entra en [entra.microsoft.com](https://entra.microsoft.com) con tu **cuenta de trabajo**.
2. **Aplicaciones** → **Registros de aplicaciones** → **Nuevo registro**.
3. Igual que arriba, pero en **Tipos de cuenta compatibles** elige
   **«Cuentas en cualquier directorio de organización y cuentas personales de Microsoft»**.
4. Copia el **Id. de aplicación (cliente)**.

Dos avisos honestos sobre este camino:

- Si Exus tiene desactivado que los usuarios registren aplicaciones, el botón te dará error. Lo
  sabrás en diez segundos.
- El registro es un objeto del directorio de tu empresa: IT lo puede ver y borrar, y desaparece el
  día que te vayas. Además, montar una app personal en el directorio corporativo puede chocar con
  la política interna. Si te da reparo, ve al camino A.

### En los dos casos

- **No crees ningún secreto de cliente.** Las apps de navegador no lo usan (van con PKCE), y por
  eso el Id. de aplicación no es información sensible.
- En **Permisos de API**, añade de *Microsoft Graph → Permisos delegados*: **`Files.ReadWrite`**.
  Con cuenta personal no hace falta que lo apruebe ningún administrador: el permiso lo das tú
  sobre tu propio OneDrive la primera vez que entres. (Con eso basta: la app solo usa lectura y
  escritura de archivos, no la API de Excel.)

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
3. Se abre tu OneDrive. Navega por las carpetas (o pulsa **Recientes**) hasta encontrar
   `Dineritos Pro.xlsx` y pulsa **Es este**. La app lo lee y ya tienes todo.
   > El buscador por nombre también está, pero es menos fiable: el índice de OneDrive tarda un
   > rato en ver un archivo recién subido. Si buscando no aparece, navega por carpetas o mira
   > en Recientes, que van directos y siempre lo encuentran.
4. En el navegador del móvil: **Compartir → Añadir a pantalla de inicio**. Queda como una app
   más y arranca a pantalla completa.

Repite el paso 2-3 en el portátil si quieres usarla también allí. Cada dispositivo guarda su
propia sesión; el Excel es lo que los mantiene sincronizados.

---

## Revisarlo desde el ordenador del trabajo

El libro vive en tu OneDrive personal, y la app entra con esa cuenta. Para poder abrirlo desde el
ordenador de la oficina tienes dos caminos.

**A. Compartirlo con tu cuenta profesional** (recomendado)

1. En [onedrive.live.com](https://onedrive.live.com), botón derecho sobre `Dineritos Pro.xlsx` →
   **Compartir**.
2. Escribe tu dirección de trabajo y —esto es lo importante— pon el permiso en **«Puede ver»**,
   no en «Puede editar».
3. Te llega un correo al buzón del trabajo con el enlace. Guárdalo en favoritos y ya lo abres
   cuando quieras, en Excel para la web.

Por qué en solo lectura: la app sube el libro entero cada vez que guardas. Si desde el trabajo lo
editaras a la vez, uno de los dos cambios se perdería. En solo lectura eso no puede pasar, y para
revisar es lo que necesitas. Si algún día tienes que editar desde el trabajo, cambias el permiso
un momento y luego lo vuelves a dejar en «Puede ver».

**B. Entrar con tu cuenta personal en el navegador del trabajo**

Sin compartir nada: abres onedrive.live.com, inicias sesión con la cuenta personal y ahí está.
Más simple, pero deja una sesión personal en un equipo de la empresa, que en algunos sitios no
está bien visto.

**Tres cosas que conviene saber**

- La app **no se entera** de con quién compartes el libro: entra con tu cuenta personal y usa su
  propio OneDrive. Compartir no le afecta en nada.
- Si el libro cambia por otro lado entre que la app lo leyó y guardas, la app lo detecta y te
  pregunta en vez de pisarlo.
- Algunas empresas bloquean en su red los dominios de OneDrive personal, y el correo con la
  invitación queda en tu buzón de trabajo (o sea, visible para la empresa). Nada de eso rompe
  nada, pero mejor saberlo antes.

---

## El uso diario

- **Anotar un gasto**: pestaña Mes → `＋ Anotar gasto` → concepto, categoría, importe → *Anotar*.
- **Llevarlo al Excel**: botón **Guardar** de la barra de arriba.
- Las primeras veces, usa **Ver qué se va a escribir** antes de guardar: te enseña rango por
  rango lo que va a tocar, sin enviar nada.
- **Cierra el libro en el ordenador antes de guardar**, o Excel se quejará de coautoría.

---

## Cosas que conviene saber

- **Cómo escribe**: baja el libro, cambia dentro solo las celdas que toca y lo vuelve a subir.
  Todo lo demás se copia byte a byte. Comprobado sobre tu libro real: de sus 45 componentes
  internos, 41 quedan idénticos y solo cambian las hojas afectadas; los 9 componentes de los
  gráficos y dibujos, el formato condicional (69 reglas) y las validaciones (19) siguen intactos.
  El libro queda marcado para que Excel recalcule los totales al abrirlo.
  > No se usa la API de Excel de Graph (`/workbook/...`) porque Microsoft **no la soporta en
  > OneDrive personal**, solo en OneDrive de empresa.
- **Si algo saliera mal**, OneDrive guarda la versión anterior: botón derecho sobre el archivo →
  *Historial de versiones*.
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

Sube de nuevo **los tres**: `app.js`, `index.html` y `sw.js`. El `index.html` importa porque
llama al script con la versión en la dirección (`app.js?v=afa5ff3`), y eso es lo que impide que
un navegador te sirva una copia vieja.

**Para saber qué versión tiene cada dispositivo**, mira el final de la pantalla: pone
`Versión` y siete caracteres. Si el móvil y el ordenador muestran códigos distintos, uno de los dos
se ha quedado atrás: cierra la app del todo y vuelve a abrirla. En Android, si se resiste:
Chrome → ⋮ → Configuración → Configuración de sitios → Todos los sitios → tu dirección →
*Borrar y restablecer*.

## Copias de seguridad

- **Ajustes → Copia de seguridad (JSON)** te baja todo el estado.
- El script `sincronizar_excel.py` de la carpeta de arriba sigue funcionando como vía alternativa
  desde el ordenador, si algún día el login te da problemas.

## Si algo va mal

| Síntoma | Causa casi siempre |
|---|---|
| «La vuelta del login no cuadra» | La URI de redirección registrada no coincide *exactamente*. Repasa la barra final. |
| El login abre y vuelve sin entrar | El registro no es de tipo **SPA**, o los tipos de cuenta no admiten cuentas personales. |
| `AADSTS16000 ... tenant 'Microsoft Services'` al entrar en el portal | Tu cuenta personal no tiene directorio. Es el paso 2: camino A o camino B. |
| «No he podido descargar el libro» | Falta el permiso `Files.ReadWrite`, o el libro se movió. Prueba *Elegir otro libro*. |
| No aparece «Añadir a pantalla de inicio» | La dirección no es `https://`, o falta `manifest.webmanifest`. |
| En el móvil se comporta como una versión antigua | Compara el `Versión` del pie en los dos dispositivos. Cierra la app y reábrela. |
| «Se ha perdido el hilo del login» | La vuelta del login aterrizó en otro sitio. Haz el primer login desde el navegador, no desde el icono de la pantalla de inicio. |
| El buscador no encuentra tu Excel | El índice de OneDrive va con retraso. Usa **Recientes** o navega por carpetas. |
| Excel se queja al guardar | Tienes el libro abierto en el ordenador. Ciérralo. |
