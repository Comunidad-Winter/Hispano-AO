# Hispano-AO

> **Proyecto de preservación.** Fork operacional de Argentum Online 0.13 en Visual Basic 6, desarrollado y mantenido por el servidor Hispano-AO. Incluye cliente, servidor y World Editor con datos del mundo completos.

![Captura de pantalla 2025-03-31 175622](https://github.com/user-attachments/assets/4ef43207-d961-44b7-9498-390c8db0441c)

---

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Qué es

**Hispano-AO** es un servidor privado de **Argentum Online** basado en el código fuente histórico de la versión **0.13** del juego. Argentum Online (AO) es un MMORPG 2D de origen argentino creado originalmente por Pablo Ignacio Márquez, licenciado bajo la **Affero General Public License**.

Este repositorio contiene el código fuente de tres componentes operacionales:

- **Cliente** (`Hispano Online.exe`) — interfaz gráfica del jugador.
- **Servidor** (`server9arreglado.exe`) — lógica de juego y networking.
- **World Editor** — herramienta de edición de mapas.

El proyecto está en estado de **preservación activa**: el código es funcional en términos históricos pero requiere entorno VB6 y dependencias de la época (Windows XP/7 era target principal). No es un proyecto de desarrollo activo con nuevas features de forma regular.

---

## Estructura del repositorio

```
Hispano-AO/
├── client/                 # Proyecto VB6 del cliente de juego
│   ├── CODIGO/             # Código fuente completo (.bas, .cls, .frm)
│   ├── Graficos/           # Sprites, gráficos de personajes y tiles
│   ├── Interfaces/         # Assets visuales de la UI (botones, ventanas .jpg)
│   ├── Mapas/              # Archivos de mapas del cliente (.map, 1–100)
│   ├── INIT/               # Datos de configuración local del cliente
│   ├── MIDI/               # Música MIDI del juego
│   ├── MP3/                # Música MP3
│   ├── WAV/                # Efectos de sonido (.wav)
│   ├── Screenshots/        # Capturas guardadas por el cliente
│   ├── EXTRAS/             # Assets adicionales (Npcs.dat, Obj.dat, HECHIZOS.dat)
│   ├── Client.vbp          # Archivo de proyecto Visual Basic 6
│   ├── Client.vbw          # Workspace VB6
│   ├── *.OCX / *.DLL       # Dependencias en tiempo de ejecución
│   └── Autoupdate.exe      # Binario de actualización automática del cliente
│
├── server/                 # Proyecto VB6 del servidor de juego
│   ├── Codigo/             # Código fuente completo del servidor
│   ├── Maps/               # Mapas del servidor (Mapa1–93: .map + .dat + .inf)
│   ├── Dat/                # Datos del mundo (NPCs, objetos, hechizos, etc.)
│   ├── Charfile/           # Archivos de personajes guardados (runtime)
│   ├── Guilds/             # Datos de clanes/guilds
│   ├── Logs/               # Logs de sesión del servidor (runtime)
│   ├── WorldBackup/        # Backups del estado del mundo
│   ├── Server.ini          # Configuración principal del servidor
│   ├── Configuracion.ini   # Parámetros del juego (niveles, rates, intervalos)
│   ├── ItemsShop.ini       # Configuración de la tienda de ítems
│   ├── SERVER.VBP          # Archivo de proyecto Visual Basic 6
│   └── server_fotos.exe    # Utilidad auxiliar (toma capturas del mapa)
│
└── WE HISPANO AO/          # World Editor (herramienta de edición de mapas)
    ├── Maps/               # Mapas extra y de prueba para el editor
    ├── Graficos/           # Gráficos de referencia del editor
    ├── Dat/                # Datos de referencia para el editor
    ├── INIT/               # Configuración del editor
    ├── MIDI/ / Wavs/       # Audio de referencia
    ├── WorldEditor.exe     # Ejecutable del editor de mapas
    ├── MAPA 101.map/.dat/.inf  # Mapa de prueba 101
    ├── MAPA TORNEO TEST.*  # Mapa de torneo para pruebas
    └── TESTEO.*            # Mapa de pruebas generales
```

---

## Componentes principales

### 1. Cliente (`client/`)

- **Lenguaje:** Visual Basic 6
- **Entry point:** `Sub Main` en `Application.bas`
- **Formulario principal:** `frmMain` (ventana de juego principal)
- **Ejecutable compilado:** `Hispano Online.exe`
- **Versión:** 0.13.8 (según `Client.vbp`, `MajorVer=0`, `MinorVer=13`, `RevisionVer=8`)

El cliente es responsable de renderizar el mundo, la interfaz del jugador, la comunicación TCP con el servidor y todos los diálogos del juego.

**Módulos de código principales:**

| Módulo | Descripción |
|---|---|
| `TileEngine.bas` (~89 KB) | Motor de renderizado 2D por tiles usando DIBSections |
| `Protocol.bas` (~420 KB) | Protocolo de comunicación cliente→servidor |
| `ProtocolCmdParse.bas` (~108 KB) | Parser de comandos recibidos del servidor |
| `General.bas` (~57 KB) | Lógica general de juego en cliente |
| `clsAudio.cls` (~41 KB) | Sistema de audio (WAV/MIDI/MP3 vía DirectX) |
| `clsByteQueue.cls` (~45 KB) | Cola de bytes para el protocolo TCP |
| `clsCripto.cls` (~52 KB) | Sistema de cifrado de datos |
| `clsGrapchicalInventory.cls` (~29 KB) | Inventario gráfico |
| `modScreenCapture.bas` (~26 KB) | Sistema de capturas de pantalla |
| `Declares.bas` (~31 KB) | Declaraciones globales y tipos de datos |

**Formularios/Ventanas:**

| Formulario | Descripción |
|---|---|
| `frmMain` | Ventana principal del juego |
| `frmConnect` | Pantalla de conexión al servidor |
| `frmCrearPersonaje` | Creación de personaje |
| `frmSkills3` | Panel de habilidades |
| `frmPanelGm` | Panel de Game Master |
| `frmGuildBrief/Details/Leader/Member` | Sistema de clanes/guilds |
| `frmComerciar` | Comercio con NPCs |
| `frmComerciarUsu` | Comercio entre usuarios |
| `frmHerrero / frmCarp` | Herrería y carpintería |
| `frmRetos` | Sistema de retos 1vs1 y 2vs2 |
| `frmCanjes` | Sistema de canje de puntos |
| `frmBancoObj` | Banco de objetos |
| `frmParty` | Sistema de grupos (party) |
| `FrmEstadisticas` | Estadísticas del personaje |

**Dependencias de runtime (OCX/DLL incluidas):**

| Archivo | Descripción |
|---|---|
| `dx7vb.dll` / `dx8vb.dll` | DirectX 7/8 para Visual Basic |
| `MSVBVM60.DLL` | Runtime VB6 |
| `MSCOMCTL.OCX` | Controles comunes de Microsoft |
| `CSWSK32.OCX` | Control de sockets TCP (Crescent Software) |
| `RICHTX32.OCX` | Control de Rich Text |
| `MSINET.OCX` | Control de Internet (WinSock) |
| `quartz.dll` | DirectShow (reproducción de audio) |
| `ijl11.dll` | Intel JPEG Library |
| `Unzip32.dll` | Descompresión ZIP (para actualizaciones) |

---

### 2. Servidor (`server/`)

- **Lenguaje:** Visual Basic 6
- **Entry point:** `Sub Main` en `frmMain.frm`
- **Formulario principal:** `frmMain` (panel de control del servidor)
- **Ejecutable compilado:** `server9arreglado.exe`
- **Versión:** 0.13.0 (según `SERVER.VBP`)
- **Puerto de escucha:** 7666 (TCP, configurable en `Server.ini`)
- **Máximo de usuarios:** 550 (configurable en `Server.ini`)
- **Condicional de compilación:** `UsarQueSocket = 1 : ConUpTime = 1`

**Módulos de código principales:**

| Módulo | Descripción |
|---|---|
| `Protocol.bas` (~903 KB) | Protocolo completo servidor (el módulo más grande) |
| `Modulo_UsUaRiOs.bas` (~150 KB) | Gestión de usuarios conectados |
| `Trabajo.bas` (~128 KB) | Sistema de trabajos (pesca, tala, minería, carpintería) |
| `SistemaCombate.bas` (~124 KB) | Motor de combate PvP y PvE |
| `FileIO.bas` (~122 KB) | Persistencia: carga/guardado de personajes, mapas, datos |
| `modHechizos.bas` (~127 KB) | Sistema de hechizos/magia |
| `GameLogic.bas` (~83 KB) | Lógica de juego general |
| `modGuilds.bas` (~85 KB) | Sistema de clanes |
| `AI_NPC.bas` (~62 KB) | Inteligencia artificial de NPCs |
| `ModFacciones.bas` (~41 KB) | Sistema de facciones (Real/Caos) |
| `TCP.bas` (~72 KB) | Capa de networking y gestión de conexiones |
| `clsClanPretoriano.cls` (~118 KB) | Sistema de Pretorianos (NPCs de facción) |
| `InvUsuario.bas` (~109 KB) | Gestión del inventario de usuarios |
| `Admin.bas` (~22 KB) | Comandos de administración GM |
| `PathFinding.bas` (~12 KB) | Algoritmo de pathfinding para NPCs |
| `SecurityIp.bas` (~14 KB) | Control y baneo de IPs |
| `modCentinela.bas` (~27 KB) | Sistema anti-macro/anti-trampa |
| `Mod_Retos1vs1.bas` (~26 KB) | Sistema de retos 1 vs 1 |
| `Mod_Retos2vs2.bas` (~39 KB) | Sistema de retos 2 vs 2 |
| `modBanco.bas` (~15 KB) | Sistema de banco de objetos |
| `mdlCOmercioConUsuario.bas` (~22 KB) | Comercio entre jugadores |
| `praetorians.bas` (~40 KB) | Lógica de NPCs Pretorianos |
| `modNuevoTimer.bas` (~19 KB) | Game loop basado en timers con intervalos configurables |

**Enumeraciones de datos del juego (verificadas en código):**

- **Clases:** Mago, Clérigo, Guerrero, Asesino, Ladrón, Bardo, Druida, Bandido, Paladín, Cazador, Trabajador, Pirata (12 clases)
- **Razas:** Humano, Elfo, Drow, Gnomo, Enano (5 razas)
- **Géneros:** Hombre, Mujer
- **Ciudades:** Ullathorpe, Nix, Banderbill, Lindos, Arghal, Arkhein, LastCity (7 ciudades)
- **Tipos de jugador:** User, Consejero, SemiDios, Dios, Admin, RoleMaster, ChaosCouncil, RoyalCouncil
- **Nivel máximo:** 47 (configurable en `Configuracion.ini`)
- **Rate de experiencia:** 8x | **Rate de oro:** 5x

---

### 3. World Editor (`WE HISPANO AO/`)

- **Ejecutable:** `WorldEditor.exe` (1 MB)
- **Dependencia:** `zlib.dll` (compresión), `libreria de render.exe`
- Permite crear y editar mapas del mundo en formato nativo AO (`.map` + `.dat` + `.inf`)
- Incluye mapas de prueba: Mapa 101, Mapa Torneo Test, TESTEO
- Los mapas exportados son compatibles directamente con el servidor

---

## Cómo funciona

### Arquitectura general

```
┌─────────────────────────────────────────────────────────┐
│                     CLIENTE VB6                         │
│  frmMain (UI) → TileEngine (render) → clsAudio (audio) │
│       ↕ TCP (CSWSK32.OCX / puerto 7666)                 │
│  clsByteQueue → Protocol.bas → ProtocolCmdParse.bas     │
└────────────────────────┬────────────────────────────────┘
                         │ TCP binario
┌────────────────────────▼────────────────────────────────┐
│                    SERVIDOR VB6                          │
│  frmMain (control) → modNuevoTimer (game loop 50ms)     │
│  TCP.bas + wsksock.bas (hasta 550 conexiones)           │
│  Protocol.bas → GameLogic / SistemaCombate / modHechizos│
│  FileIO.bas ↔ /Charfile (personajes) /Maps /Dat         │
└─────────────────────────────────────────────────────────┘
```

### Protocolo de red

- Comunicación **TCP binaria** sobre el puerto **7666**
- El cliente usa `CSWSK32.OCX` (Crescent WinSock) como capa de socket
- El servidor usa `wsksock.bas` + `wskapiAO.bas` para gestión de hasta 550 conexiones simultáneas
- Los datos se serializan/deserializan mediante `clsByteQueue` (ambos extremos)
- El servidor implementa **verificación MD5** del ejecutable del cliente (`Server.ini → [MD5Hush]`) para control de versiones

### Game loop

- El servidor ejecuta su ciclo principal mediante `modNuevoTimer.bas` con un intervalo base de **50 ms** (`IntervaloTimerExec`)
- Los intervalos de acción (ataque, hechizos, trabajo, movimiento NPC) son configurables en `Server.ini` y `Configuracion.ini`
- La IA de NPCs se ejecuta cada ~380 ms (`IntervaloNpcAI`)

### Mapas

Cada mapa se representa con tres archivos complementarios:

| Extensión | Contenido |
|---|---|
| `.map` | Geometría de tiles, capas, objetos estáticos (binario) |
| `.dat` | Metadatos del mapa (nombre, música, configuración) |
| `.inf` | Información adicional de tiles (triggers, spawns de NPCs, etc.) |

El mundo tiene **93 mapas** activos numerados (Mapa1–Mapa93). Los clientes solo reciben `.map`; el servidor carga las tres variantes.

### Datos del servidor (`server/Dat/`)

| Archivo | Contenido |
|---|---|
| `NPCs.dat` (~197 KB) | Definición completa de todos los NPCs del juego |
| `obj.dat` (~195 KB) | Definición de todos los objetos/ítems |
| `Hechizos.dat` (~54 KB) | Definición de todos los hechizos |
| `Balance.dat` | Modificadores de clase para el sistema de combate |
| `AreasStats.dat` | Estadísticas de áreas del mundo |
| `Pretorianos.dat` | Configuración de NPCs Pretorianos (facción) |
| `Invokar.dat` | Datos de invocación de NPCs |
| `BanIps.dat` | Lista de IPs baneadas |
| `NombresInvalidos.txt` | Lista de nombres de personaje prohibidos |
| `RECORDS.DAT` | Records globales del servidor |

---

## Instalación / Compilación / Ejecución

### Requisitos

- **Windows** (XP / 7 / 10 con compatibilidad) — sistema operativo objetivo original
- **Visual Basic 6** — IDE y compilador (Microsoft Visual Basic 6.0 SP6)
- Las DLL y OCX incluidas en `client/` deben estar registradas en el sistema (`regsvr32`)
- **DirectX 7 o superior** (incluido en sistemas Windows modernos)

### Compilar el cliente

1. Abrir `client/Client.vbp` en el IDE de VB6
2. Verificar referencias a OCX (CSWSK32, MSCOMCTL, COMCTL32, MSINET, RICHTX32)
3. Compilar → `File > Make Hispano Online.exe` (salida en `client/../`)
4. El ejecutable compilado se llama `Hispano Online.exe`

### Compilar el servidor

1. Abrir `server/SERVER.VBP` en el IDE de VB6
2. Verificar condición de compilación: `UsarQueSocket = 1 : ConUpTime = 1`
3. Compilar → `File > Make server9arreglado.exe` (salida en `server/../`)

### Configurar el servidor antes de ejecutar

1. Editar `server/Server.ini`:
   - `StartPort` → puerto TCP (por defecto 7666)
   - `MaxUsers` → límite de conexiones simultáneas
   - `Version` → versión de cliente aceptada
   - Sección `[Admines]` / `[Dioses]` / `[SemiDioses]` → nombres de GM
2. Editar `server/Configuracion.ini` para ajustar rates, niveles máximos e intervalos de combate
3. Crear el directorio `server/Charfile/` si no existe (almacena archivos de personajes)
4. Verificar que existan los directorios `server/Logs/` y `server/Guilds/`

### Ejecutar

1. Ejecutar `server9arreglado.exe` desde el directorio `server/`
2. Ejecutar `Hispano Online.exe` desde el directorio `client/../`
3. En la pantalla de conexión, apuntar a `localhost:7666` o la IP del servidor

---

## Estado del proyecto

| Aspecto | Estado |
|---|---|
| Tipo | Preservación / servidor privado histórico |
| Código fuente | Completo y disponible (cliente + servidor + editor) |
| Compilabilidad | Requiere VB6 — entorno obsoleto pero funcional |
| Datos del mundo | Completos (93 mapas, NPCs, objetos, hechizos) |
| Desarrollo activo | No verificado — el repositorio parece ser un snapshot histórico |
| Licencia del código base | Affero GPL (heredada del proyecto original AO) |

El código fuente lleva comentarios de múltiples desarrolladores a lo largo del tiempo, lo cual indica que fue un proyecto con contribuciones comunitarias activas. El historial de modificaciones visible en los comentarios abarca desde ~2003 hasta ~2015 (fechas de última modificación en cabeceras).

---

## Créditos y licencia

- **Concepto y código original de Argentum Online:** Pablo Ignacio Márquez (`morgolock@speedy.com.ar`)
- **Diseño del módulo de combate:** Gerardo Saiz
- **Contribuciones al servidor:** Juan Martín Sotuyo Dodero (maraxus), ZaMa, Miqueas, Barrin, El Oso (Mariano Barrou), y otros colaboradores anónimos mencionados en cabeceras de código
- **Fork Hispano-AO:** Equipo de Hispano-AO (`VersionCompanyName = "Hispano AO"` en `Client.vbp`)
- **Licencia:** [Affero General Public License v1+](http://www.affero.org/oagpl.html)

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Notas y advertencias

- Ver [`docs/notes.md`](docs/notes.md) para asunciones, cosas no verificadas y advertencias técnicas detalladas.
- Ver [`docs/architecture.md`](docs/architecture.md) para la arquitectura técnica en profundidad.
- Ver [`docs/components.md`](docs/components.md) para el detalle de cada módulo.
- Ver [`docs/build-and-run.md`](docs/build-and-run.md) para instrucciones extendidas de compilación.
