# Notas, Asunciones y Advertencias — Hispano-AO

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Metodología de análisis

Esta documentación fue generada mediante análisis directo de:
- Estructura de archivos del repositorio
- Archivos de proyecto `.vbp` (Client.vbp, SERVER.VBP)
- Archivos de configuración `.ini` (Server.ini, Configuracion.ini)
- Código fuente VB6 seleccionado (módulos críticos: Declares.bas, FileIO.bas, modNuevoTimer.bas, SistemaCombate.bas, TCP.bas, Application.bas, clsClanPretoriano.cls, modCentinela.bas)
- Listados de directorios y tamaños de archivos

**No se ejecutó el código** — toda la documentación es producto de análisis estático.

---

## Asunciones y elementos inferidos

Las siguientes afirmaciones están marcadas como inferidas porque no se verificaron directamente en el código fuente, sino que se deducen de la estructura del proyecto, nomenclatura, o patrones comunes en AO:

### Protocolo de red
- Se asume que el protocolo usa un **byte de opcode** al inicio de cada paquete, basado en el patrón estándar de AO y la existencia de parsers estructurados en `Protocol.bas` y `ProtocolCmdParse.bas`. **No verificado directamente** en el código de los parsers (archivo de 903 KB no analizado en su totalidad).
- El cifrado en `clsCripto.cls` (~52 KB) podría aplicarse al handshake inicial o a datos sensibles (contraseñas). **No verificado** su uso exacto en el flujo de red.

### Motor gráfico
- Se asume que `TileEngine.bas` usa exclusivamente GDI/DIBSections basado en la presencia de `cDIBSection.cls` y la ausencia de referencias DirectX para gráficos en el `.vbp`. **Probable pero no exhaustivamente verificado**.

### Sistema de atributos de personaje
- Los atributos del personaje (Fuerza, Agilidad, etc.) se infieren del sistema de creación (`frmCrearPersonaje`) y de la lógica de combate. No se analizó en detalle la estructura `UserList()` completa en `Declares.bas`.

### `Autoupdate.exe`
- Se asume que gestiona parches automáticos basado en su nombre, tamaño (~36 KB), y la presencia de `MSINET.OCX` y `Unzip32.dll`. **No verificado** — no hay código fuente disponible para este binario.

### `server_fotos.exe`
- Se asume que toma capturas del estado del mapa basado en su nombre. **No verificado** — no hay código fuente disponible.

### World Editor
- El editor no tiene código fuente en el repositorio. Toda la documentación del editor es **inferida** del ejecutable y los datos asociados.

---

## Cosas no verificadas

| Elemento | Estado |
|---|---|
| Contenido completo de `Protocol.bas` del servidor (903 KB) | No analizado en su totalidad |
| Contenido completo de `Modulo_UsUaRiOs.bas` (150 KB) | No analizado en su totalidad |
| Contenido completo de `Trabajo.bas` (128 KB) | No analizado en su totalidad |
| Estructura exacta de archivos `.charfile` | No analizada |
| Formato binario exacto de archivos `.map` e `.inf` | No documentado aquí |
| Sistema de canje de puntos (`frmCanjes`) — lógica de servidor | No analizado |
| Sistema de matrimonio (`Configuracion.ini → [Matrimonio]`) | No analizado |
| Lógica completa de `clsCripto.cls` | No analizada |
| URLs del sistema de actualización automática | No encontradas en el repositorio |
| Contenido de `client/Screenshots/` | No listado |
| Contenido de `client/Graficos/` | No listado |
| Contenido exacto de `client/MIDI/` y `client/MP3/` | No listado |

---

## Posibles problemas y deuda técnica

### Código

1. **Módulo `Protocol.bas` del servidor (903 KB):** Un único archivo de 903 KB es difícil de mantener. Sugiere que el protocolo nunca fue refactorizado en módulos más pequeños.

2. **Nombres de módulos inconsistentes:** El módulo `TCP.bas` del servidor contiene lógica de apariencia de personajes (cabezas por raza/género), no solo networking. La separación de responsabilidades es deficiente.

3. **VB6 single-threaded:** VB6 es inherentemente de un solo hilo (`MaxNumberOfThreads=1`). El servidor maneja hasta 550 usuarios concurrentes en un solo hilo mediante un event loop de timers. Esto es una limitación arquitectural fundamental que impide escalar en hardware moderno.

4. **Archivos `.log` en CODIGO/:** Se encontraron archivos `.log` dentro del directorio de código fuente (`frmCargando.log`, `frmMain.log`, `frmOpciones.log`, `frmPanelGm.log`). Estos parecen ser artefactos del IDE de VB6 y no deberían estar versionados.

5. **Archivo `VB7.tmp` en client/ y `VB8BF7.tmp` en server/:** Archivos temporales del IDE VB6 que no deberían estar en el repositorio.

6. **`clsGrapchicalInventory.cls`:** El nombre tiene un typo ("Grapchical" en lugar de "Graphical"). Este typo está presente tanto en el nombre del archivo como en las referencias del proyecto y es heredado del código base original.

### Datos

7. **`ItemsShop.ini`:** Existe en `server/` pero no se analizó su estructura. Probablemente define una tienda de ítems de pago o premium (contexto: "Compra tus personajes" en mensajes del servidor).

8. **`apuestas.dat` (44 bytes):** Muy pequeño, probablemente un sistema de apuestas incompleto o vacío.

9. **`consultas.dat` (254 bytes) y `ConsultasPopulares.cls`:** Sistema de consultas frecuentes de estado del servidor, posiblemente para una interfaz web externa. No verificado.

10. **Mapas del cliente vs servidor:** El cliente tiene **94 mapas** (Mapa1–Mapa100, con gaps) y el servidor tiene **93 mapas** con tres archivos cada uno. Hay discrepancias entre los números de mapa disponibles en cada lado. **No se verificó** cuál es el set canónico.

11. **`Mapa100.map` en cliente:** Existe en el cliente pero no hay correspondiente `Mapa100.dat/.inf` en el servidor. Podría ser un mapa incompleto o de prueba.

### Seguridad

12. **Contraseñas en texto plano:** No se puede confirmar si las contraseñas están hasheadas en los archivos de personaje sin analizar `FileIO.bas` en detalle. El sistema de `clsCripto.cls` existe pero su uso exacto para contraseñas no está verificado.

13. **`BanIps.dat` (229 bytes):** Lista de IPs baneadas muy pequeña, posiblemente desactualizada.

14. **Protocolo sin TLS:** Toda la comunicación es en texto/binario plano. Esto es esperado para la época del desarrollo (~2003-2015) pero es un riesgo si se despliega en redes no confiables.

---

## Estado real del proyecto

### Línea de tiempo del código (inferida de comentarios en cabeceras)

| Período | Actividad |
|---|---|
| ~2002-2003 | Código base original de Pablo Ignacio Márquez (AO 0.12.2) |
| ~2003-2008 | Contribuciones de la comunidad (ImperiumAO, AlkonAO, otros forks) |
| ~2008-2012 | Integración de sistemas Pretorianos, Centinela, Retos, Partys |
| ~2012-2015 | Ajustes de Miqueas (fechas verificadas en comentarios: 2015) |
| ~2015+ | Estado aparentemente estable — sin evidencia de cambios posteriores en el código analizado |

### Snapshot vs proyecto activo

El repositorio parece ser un **snapshot** del código en un momento específico, no un proyecto con historial de commits visible. El directorio `.git` existe pero no se analizó el historial de commits. Los directorios `CVS/` dentro de `client/CODIGO/` y `server/Codigo/` sugieren que el proyecto usó **CVS** como control de versiones en su época activa.

---

## Partes incompletas o probablemente rotas

1. **`asd.ini` y `cuerpos.ini` en `client/INIT/`:** Archivos de 34 bytes con nombres genéricos/temporales. Probablemente artefactos de desarrollo.

2. **`messages.txt` en `client/INIT/`:** Solo 53 bytes. Archivo de mensajes del sistema muy pequeño, posiblemente incompleto.

3. **Sistema de foro (`frmForo`, `modForum.bas`):** El archivo `frmForo.frx` tiene solo 6 bytes. Probablemente el formulario del foro no tiene recursos gráficos propios y usa la UI base de VB6.

4. **`clsEstadisticasIPC.cls`:** El nombre sugiere estadísticas de comunicación inter-proceso (IPC). No está claro qué proceso externo se comunica con el servidor.

5. **`ConsultasPopulares.cls`:** Podría ser para una API web o monitor externo del servidor. No verificado.

---

## Recomendaciones para preservación

1. **Documentar el formato de mapa:** Los archivos `.map`, `.dat` e `.inf` son binarios propietarios. Sin documentación del formato, los mapas no pueden ser editados sin el World Editor original.

2. **Preservar los ejecutables binarios:** `WorldEditor.exe`, `Autoupdate.exe`, `server_fotos.exe` y `libreria de render.exe` no tienen código fuente disponible en este repositorio. Son los más vulnerables a perderse.

3. **Conservar OCX y DLL:** Las versiones exactas de `CSWSK32.OCX`, `RICHTX32.OCX` y otras OCX incluidas son críticas para la compilación y ejecución. Versiones más nuevas pueden no ser compatibles.

4. **Registrar hashes MD5 de los ejecutables:** El servidor verifica el MD5 de `Hispano Online.exe`. Los hashes `6c27fd12ae7e76247d1a3fb278321240` (cliente normal) y `2c5c6253c3fb6c81617f315e0f7d589f` (cliente AlphaBlending) deberían preservarse.

5. **Datos de `server/Dat/`:** Los archivos `NPCs.dat` y `obj.dat` (~195 KB cada uno) contienen toda la definición del contenido del juego. Son irreemplazables sin las herramientas de edición originales.
