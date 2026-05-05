# Notas, Asunciones y Advertencias ÔÇö Hispano-AO

> *Esta documentaci├│n fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Metodolog├¡a de an├ílisis

Esta documentaci├│n fue generada mediante an├ílisis directo de:
- Estructura de archivos del repositorio
- Archivos de proyecto `.vbp` (Client.vbp, SERVER.VBP)
- Archivos de configuraci├│n `.ini` (Server.ini, Configuracion.ini)
- C├│digo fuente VB6 seleccionado (m├│dulos cr├¡ticos: Declares.bas, FileIO.bas, modNuevoTimer.bas, SistemaCombate.bas, TCP.bas, Application.bas, clsClanPretoriano.cls, modCentinela.bas)
- Listados de directorios y tama├▒os de archivos

**No se ejecut├│ el c├│digo** ÔÇö toda la documentaci├│n es producto de an├ílisis est├ítico.

---

## Asunciones y elementos inferidos

Las siguientes afirmaciones est├ín marcadas como inferidas porque no se verificaron directamente en el c├│digo fuente, sino que se deducen de la estructura del proyecto, nomenclatura, o patrones comunes en AO:

### Protocolo de red
- Se asume que el protocolo usa un **byte de opcode** al inicio de cada paquete, basado en el patr├│n est├índar de AO y la existencia de parsers estructurados en `Protocol.bas` y `ProtocolCmdParse.bas`. **No verificado directamente** en el c├│digo de los parsers (archivo de 903 KB no analizado en su totalidad).
- El cifrado en `clsCripto.cls` (~52 KB) podr├¡a aplicarse al handshake inicial o a datos sensibles (contrase├▒as). **No verificado** su uso exacto en el flujo de red.

### Motor gr├ífico
- Se asume que `TileEngine.bas` usa exclusivamente GDI/DIBSections basado en la presencia de `cDIBSection.cls` y la ausencia de referencias DirectX para gr├íficos en el `.vbp`. **Probable pero no exhaustivamente verificado**.

### Sistema de atributos de personaje
- Los atributos del personaje (Fuerza, Agilidad, etc.) se infieren del sistema de creaci├│n (`frmCrearPersonaje`) y de la l├│gica de combate. No se analiz├│ en detalle la estructura `UserList()` completa en `Declares.bas`.

### `Autoupdate.exe`
- Se asume que gestiona parches autom├íticos basado en su nombre, tama├▒o (~36 KB), y la presencia de `MSINET.OCX` y `Unzip32.dll`. **No verificado** ÔÇö no hay c├│digo fuente disponible para este binario.

### `server_fotos.exe`
- Se asume que toma capturas del estado del mapa basado en su nombre. **No verificado** ÔÇö no hay c├│digo fuente disponible.

### World Editor
- El editor no tiene c├│digo fuente en el repositorio. Toda la documentaci├│n del editor es **inferida** del ejecutable y los datos asociados.

---

## Cosas no verificadas

| Elemento | Estado |
|---|---|
| Contenido completo de `Protocol.bas` del servidor (903 KB) | No analizado en su totalidad |
| Contenido completo de `Modulo_UsUaRiOs.bas` (150 KB) | No analizado en su totalidad |
| Contenido completo de `Trabajo.bas` (128 KB) | No analizado en su totalidad |
| Estructura exacta de archivos `.charfile` | No analizada |
| Formato binario exacto de archivos `.map` e `.inf` | No documentado aqu├¡ |
| Sistema de canje de puntos (`frmCanjes`) ÔÇö l├│gica de servidor | No analizado |
| Sistema de matrimonio (`Configuracion.ini ÔåÆ [Matrimonio]`) | No analizado |
| L├│gica completa de `clsCripto.cls` | No analizada |
| URLs del sistema de actualizaci├│n autom├ítica | No encontradas en el repositorio |
| Contenido de `client/Screenshots/` | No listado |
| Contenido de `client/Graficos/` | No listado |
| Contenido exacto de `client/MIDI/` y `client/MP3/` | No listado |

---

## Posibles problemas y deuda t├®cnica

### C├│digo

1. **M├│dulo `Protocol.bas` del servidor (903 KB):** Un ├║nico archivo de 903 KB es dif├¡cil de mantener. Sugiere que el protocolo nunca fue refactorizado en m├│dulos m├ís peque├▒os.

2. **Nombres de m├│dulos inconsistentes:** El m├│dulo `TCP.bas` del servidor contiene l├│gica de apariencia de personajes (cabezas por raza/g├®nero), no solo networking. La separaci├│n de responsabilidades es deficiente.

3. **VB6 single-threaded:** VB6 es inherentemente de un solo hilo (`MaxNumberOfThreads=1`). El servidor maneja hasta 550 usuarios concurrentes en un solo hilo mediante un event loop de timers. Esto es una limitaci├│n arquitectural fundamental que impide escalar en hardware moderno.

4. **Archivos `.log` en CODIGO/:** Se encontraron archivos `.log` dentro del directorio de c├│digo fuente (`frmCargando.log`, `frmMain.log`, `frmOpciones.log`, `frmPanelGm.log`). Estos parecen ser artefactos del IDE de VB6 y no deber├¡an estar versionados.

5. **Archivo `VB7.tmp` en client/ y `VB8BF7.tmp` en server/:** Archivos temporales del IDE VB6 que no deber├¡an estar en el repositorio.

6. **`clsGrapchicalInventory.cls`:** El nombre tiene un typo ("Grapchical" en lugar de "Graphical"). Este typo est├í presente tanto en el nombre del archivo como en las referencias del proyecto y es heredado del c├│digo base original.

### Datos

7. **`ItemsShop.ini`:** Existe en `server/` pero no se analiz├│ su estructura. Probablemente define una tienda de ├¡tems de pago o premium (contexto: "Compra tus personajes" en mensajes del servidor).

8. **`apuestas.dat` (44 bytes):** Muy peque├▒o, probablemente un sistema de apuestas incompleto o vac├¡o.

9. **`consultas.dat` (254 bytes) y `ConsultasPopulares.cls`:** Sistema de consultas frecuentes de estado del servidor, posiblemente para una interfaz web externa. No verificado.

10. **Mapas del cliente vs servidor:** El cliente tiene **94 mapas** (Mapa1ÔÇôMapa100, con gaps) y el servidor tiene **93 mapas** con tres archivos cada uno. Hay discrepancias entre los n├║meros de mapa disponibles en cada lado. **No se verific├│** cu├íl es el set can├│nico.

11. **`Mapa100.map` en cliente:** Existe en el cliente pero no hay correspondiente `Mapa100.dat/.inf` en el servidor. Podr├¡a ser un mapa incompleto o de prueba.

### Seguridad

12. **Contrase├▒as en texto plano:** No se puede confirmar si las contrase├▒as est├ín hasheadas en los archivos de personaje sin analizar `FileIO.bas` en detalle. El sistema de `clsCripto.cls` existe pero su uso exacto para contrase├▒as no est├í verificado.

13. **`BanIps.dat` (229 bytes):** Lista de IPs baneadas muy peque├▒a, posiblemente desactualizada.

14. **Protocolo sin TLS:** Toda la comunicaci├│n es en texto/binario plano. Esto es esperado para la ├®poca del desarrollo (~2003-2015) pero es un riesgo si se despliega en redes no confiables.

---

## Estado real del proyecto

### L├¡nea de tiempo del c├│digo (inferida de comentarios en cabeceras)

| Per├¡odo | Actividad |
|---|---|
| ~2002-2003 | C├│digo base original de Pablo Ignacio M├írquez (AO 0.12.2) |
| ~2003-2008 | Contribuciones de la comunidad (ImperiumAO, AlkonAO, otros forks) |
| ~2008-2012 | Integraci├│n de sistemas Pretorianos, Centinela, Retos, Partys |
| ~2012-2015 | Ajustes de Miqueas (fechas verificadas en comentarios: 2015) |
| ~2015+ | Estado aparentemente estable ÔÇö sin evidencia de cambios posteriores en el c├│digo analizado |

### Snapshot vs proyecto activo

El repositorio parece ser un **snapshot** del c├│digo en un momento espec├¡fico, no un proyecto con historial de commits visible. El directorio `.git` existe pero no se analiz├│ el historial de commits. Los directorios `CVS/` dentro de `client/CODIGO/` y `server/Codigo/` sugieren que el proyecto us├│ **CVS** como control de versiones en su ├®poca activa.

---

## Partes incompletas o probablemente rotas

1. **`asd.ini` y `cuerpos.ini` en `client/INIT/`:** Archivos de 34 bytes con nombres gen├®ricos/temporales. Probablemente artefactos de desarrollo.

2. **`messages.txt` en `client/INIT/`:** Solo 53 bytes. Archivo de mensajes del sistema muy peque├▒o, posiblemente incompleto.

3. **Sistema de foro (`frmForo`, `modForum.bas`):** El archivo `frmForo.frx` tiene solo 6 bytes. Probablemente el formulario del foro no tiene recursos gr├íficos propios y usa la UI base de VB6.

4. **`clsEstadisticasIPC.cls`:** El nombre sugiere estad├¡sticas de comunicaci├│n inter-proceso (IPC). No est├í claro qu├® proceso externo se comunica con el servidor.

5. **`ConsultasPopulares.cls`:** Podr├¡a ser para una API web o monitor externo del servidor. No verificado.

---

## Recomendaciones para preservaci├│n

1. **Documentar el formato de mapa:** Los archivos `.map`, `.dat` e `.inf` son binarios propietarios. Sin documentaci├│n del formato, los mapas no pueden ser editados sin el World Editor original.

2. **Preservar los ejecutables binarios:** `WorldEditor.exe`, `Autoupdate.exe`, `server_fotos.exe` y `libreria de render.exe` no tienen c├│digo fuente disponible en este repositorio. Son los m├ís vulnerables a perderse.

3. **Conservar OCX y DLL:** Las versiones exactas de `CSWSK32.OCX`, `RICHTX32.OCX` y otras OCX incluidas son cr├¡ticas para la compilaci├│n y ejecuci├│n. Versiones m├ís nuevas pueden no ser compatibles.

4. **Registrar hashes MD5 de los ejecutables:** El servidor verifica el MD5 de `Hispano Online.exe`. Los hashes `6c27fd12ae7e76247d1a3fb278321240` (cliente normal) y `2c5c6253c3fb6c81617f315e0f7d589f` (cliente AlphaBlending) deber├¡an preservarse.

5. **Datos de `server/Dat/`:** Los archivos `NPCs.dat` y `obj.dat` (~195 KB cada uno) contienen toda la definici├│n del contenido del juego. Son irreemplazables sin las herramientas de edici├│n originales.
