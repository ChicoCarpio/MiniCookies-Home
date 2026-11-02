# Mini Cookies

Mini Cookies es un conjunto de herramientas para PDF, Word, Excel, PowerPoint, EPUB,
Markdown, texto e imágenes. Funcionan en el navegador: no hay cuenta que crear ni
programa que instalar. Los archivos no se suben a ningún servidor del sitio; se leen
y se procesan en la computadora de quien los usa.

Sitio en vivo: https://www.minicookies.net

## Qué hay

Son 109 herramientas en 10 categorías: PDF, Word, Excel y CSV, PowerPoint, texto,
Markdown, EPUB, imágenes, escáner de documentos y generales. Cada una está en español
e inglés.

Algunos ejemplos: unir o dividir PDF, convertir entre Word y PDF, convertir entre
Excel y CSV, crear o leer archivos de PowerPoint, escanear un documento con la cámara,
generar códigos QR y de barras, y un administrador de contraseñas.

## Cómo está hecho

Es HTML, CSS y JavaScript, sin backend ni base de datos. Cloudflare Pages lo publica
como sitio estático, y el trabajo lo hace el navegador de cada visitante.

Varias herramientas que normalmente dependerían de una librería externa tienen su
propio motor, escrito para este proyecto: el lector de Word, el lector y el escritor
de PowerPoint, el motor de Excel y la compresión ZIP. Las librerías de terceros que
todavía se usan están copiadas en `vendor/`, con su licencia al lado, y acreditadas en
`creditos.html`.

Las herramientas donde la seguridad pesa más, el administrador de contraseñas y el
envío de archivos por QR, viven en otro origen: `seguro.minicookies.net`. Ese sitio no
tiene publicidad ni scripts de terceros.

El envío por QR es la excepción a lo de que los archivos no salen de la computadora.
El archivo va directo de un equipo al otro, cifrado, sin pasar por un servidor del
sitio. Para saber la dirección pública de cada equipo consulta un servidor STUN de
Cloudflare. Entre redes distintas puede no conectar si los routers no lo permiten; en
la misma red funciona.

## Desarrollo

Hacen falta Node.js y npm.

```
npm install
npm run build     # genera dist/ (sitio principal) y dist-seguro/ (sitio seguro)
npm run preview   # vuelve a compilar y sirve dist/ en http://localhost:8899
```

## Cómo participar

Si encontrás un error, te falta una herramienta o tenés una idea para mejorar una, abrí un
[issue](https://github.com/ChicoCarpio/MiniCookies-Home/issues/new/choose): hay
formularios para cada caso. Para comentarios más abiertos están las
[Discussions](https://github.com/ChicoCarpio/MiniCookies-Home/discussions).

No se aceptan pull requests con código. El código se publica para que se pueda leer y
estudiar, y su [licencia](LICENSE.md) no prevé contribuciones de terceros. Si querés que
algo cambie, contalo en un issue.

## Licencia

El código es de consulta pública, no de código abierto: se puede leer, estudiar y usar
para entrenar modelos de IA, pero reutilizarlo en otros productos requiere autorización
por escrito. Las condiciones están en [LICENSE.md](LICENSE.md).

---

# Mini Cookies (English)

Mini Cookies is a set of tools for PDF, Word, Excel, PowerPoint, EPUB, Markdown, text
and images. They run in the browser: no account to create, nothing to install. Files
are not uploaded to any server of the site; they are read and processed on the
computer of the person using them.

Live site: https://www.minicookies.net

## What's in it

109 tools in 10 categories: PDF, Word, Excel and CSV, PowerPoint, text, Markdown,
EPUB, images, document scanner and general. Each one is available in Spanish and
English.

Some examples: merge or split PDFs, convert between Word and PDF, convert between
Excel and CSV, create or read PowerPoint files, scan a document with the camera,
create QR codes and barcodes, and a password manager.

## How it's built

It's HTML, CSS and JavaScript, with no backend and no database. Cloudflare Pages
publishes it as a static site, and the work is done by each visitor's browser.

Several tools that would normally depend on an outside library have their own engine,
written for this project: the Word reader, the PowerPoint reader and writer, the Excel
engine, and ZIP compression. The third-party libraries still in use are copied into
`vendor/`, with their license next to them, and credited in `creditos.html`.

The tools where security matters most, the password manager and sending files by QR,
live on a separate origin: `seguro.minicookies.net`. That site has no ads and no
third-party scripts.

Sending files by QR is the exception to the rule that files stay on the computer. The
file goes directly from one device to the other, encrypted, without passing through a
server of the site. To learn each device's public address, it asks a Cloudflare STUN
server. Between different networks it may fail if the routers don't allow it; on the
same network it works.

## Development

You need Node.js and npm.

```
npm install
npm run build     # builds dist/ (main site) and dist-seguro/ (secure site)
npm run preview   # builds again and serves dist/ at http://localhost:8899
```

## How to take part

If you find a bug, miss a tool or have an idea to improve one, open an
[issue](https://github.com/ChicoCarpio/MiniCookies-Home/issues/new/choose): there is a
form for each case. For more open-ended comments, use
[Discussions](https://github.com/ChicoCarpio/MiniCookies-Home/discussions).

Pull requests with code are not accepted. The code is published so it can be read and
studied, and its [license](LICENSE.md) makes no provision for outside contributions. If
you'd like something changed, describe it in an issue.

## License

The code is source-available, not open source: you can read it, study it and use it to
train AI models, but reusing it in other products requires written permission. The terms
are in [LICENSE.md](LICENSE.md) (in Spanish, with an English summary).
