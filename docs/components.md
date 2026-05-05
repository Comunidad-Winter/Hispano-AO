# Componentes por M├│dulo ÔÇö Hispano-AO

> *Esta documentaci├│n fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Cliente (`client/CODIGO/`)

### M├│dulos de n├║cleo

#### `Application.bas`
- **Funci├│n:** M├│dulo de arranque de la aplicaci├│n (`Sub Main`)
- **Responsabilidades:** Detecta si la aplicaci├│n est├í activa en primer plano (`IsAppActive`). Entry point de la aplicaci├│n VB6.
- **Tama├▒o:** ~1.8 KB

#### `Declares.bas` (Mod_Declaraciones)
- **Funci├│n:** Declaraciones globales, tipos de datos y constantes del cliente
- **Contenido verificado:** Tipos `Canjeo`, instancias globales de clases (`Audio`, `Inventario`, `SurfaceDB`, etc.), constantes de intervalos de acci├│n, constantes de sonido
- **Tama├▒o:** ~31 KB

#### `General.bas` (Mod_General)
- **Funci├│n:** L├│gica general de juego del lado cliente (renderizado de personajes, movimiento, interacciones)
- **Tama├▒o:** ~57 KB (m├│dulo de l├│gica principal del cliente)

#### `GameIni.bas`
- **Funci├│n:** Carga de configuraci├│n inicial del juego desde archivos `.ini` y `.dat`
- **Tama├▒o:** ~2 KB

#### `PrevInstance.bas`
- **Funci├│n:** Previene m├║ltiples instancias simult├íneas del cliente
- **Tama├▒o:** ~4.3 KB

#### `Resolution.bas`
- **Funci├│n:** Gesti├│n de resoluci├│n de pantalla y modos de visualizaci├│n
- **Tama├▒o:** ~8.2 KB

#### `ModAreas.bas`
- **Funci├│n:** L├│gica de ├íreas del mapa (zonas seguras, zonas de combate, etc.)
- **Tama├▒o:** ~2.6 KB

---

### M├│dulos de red

#### `TCP.bas` (Mod_TCP)
- **Funci├│n:** Gesti├│n de la conexi├│n TCP con el servidor
- **Tecnolog├¡a:** `CSWSK32.OCX` (Crescent WinSock)
- **Tama├▒o:** ~2.4 KB

#### `Protocol.bas`
- **Funci├│n:** Serializaci├│n de todos los comandos del cliente ÔåÆ servidor
- **Tama├▒o:** ~420 KB ÔÇö el m├│dulo m├ís grande del cliente; contiene una funci├│n de env├¡o por cada tipo de acci├│n posible del jugador
- **Nota:** El tama├▒o refleja la amplitud del protocolo (movimiento, combate, comercio, guilds, hechizos, etc.)

#### `ProtocolCmdParse.bas`
- **Funci├│n:** Parser de todos los comandos recibidos del servidor
- **Tama├▒o:** ~108 KB ÔÇö deserializa cada tipo de respuesta del servidor y actualiza el estado del juego
- **Nota:** Complemento directo de `Protocol.bas` del servidor

#### `clsByteQueue.cls`
- **Funci├│n:** Cola FIFO de bytes para serializaci├│n binaria del protocolo
- **Uso:** Instanciada como `incomingData` y `outgoingData` en `Declares.bas`
- **Tama├▒o:** ~45 KB (compartida con el servidor: mismo archivo)

---

### Motor gr├ífico

#### `TileEngine.bas` (Mod_TileEngine)
- **Funci├│n:** Motor de renderizado 2D basado en tiles
- **Tecnolog├¡a:** GDI Windows con DIBSections (sin DirectX para gr├íficos)
- **Tama├▒o:** ~89 KB ÔÇö n├║cleo del renderizado del mundo
- **Responsabilidades:** Dibujado de capas de tiles, personajes, NPCs, efectos visuales, animaciones

#### `cDIBSection.cls`
- **Funci├│n:** Encapsula una secci├│n DIB (Device-Independent Bitmap) de Windows para renderizado eficiente
- **Tama├▒o:** ~19 KB

#### `clsSurfaceManager.cls`
- **Funci├│n:** Interfaz abstracta para manejo de superficies gr├íficas
- **Tama├▒o:** ~2.8 KB

#### `clsSurfaceManStatic.cls`
- **Funci├│n:** Gesti├│n de superficies est├íticas (tiles de terreno, objetos decorativos)
- **Tama├▒o:** ~13 KB

#### `clsSurfaceManDyn.cls`
- **Funci├│n:** Gesti├│n de superficies din├ímicas (personajes, NPCs en movimiento)
- **Tama├▒o:** ~16 KB

#### `Carteles.bas`
- **Funci├│n:** Sistema de carteles sobre las cabezas de los personajes (chat, nombres, da├▒o)
- **Tama├▒o:** ~3.3 KB

#### `Mod_rDamage.bas`
- **Funci├│n:** M├│dulo de renderizado de da├▒o visual (n├║meros flotantes de da├▒o recibido/causado)
- **Tama├▒o:** ~5.5 KB

---

### Sistema de audio

#### `clsAudio.cls`
- **Funci├│n:** Gesti├│n completa de audio del juego
- **Formatos soportados:** WAV (efectos), MIDI (m├║sica de fondo), MP3 (m├║sica alternativa)
- **Tecnolog├¡a:** DirectX (quartz.dll ÔÇö DirectShow), posiblemente dx8vb.dll
- **Tama├▒o:** ~41 KB

---

### Sistema de seguridad y cifrado

#### `clsCripto.cls`
- **Funci├│n:** Cifrado y descifrado de datos para comunicaci├│n segura
- **Tama├▒o:** ~52 KB ÔÇö m├│dulo de criptograf├¡a significativo
- **Nota:** El uso exacto en el protocolo de red requiere an├ílisis adicional del c├│digo

---

### Clases de interfaz y UI

#### `clsGrapchicalInventory.cls`
- **Funci├│n:** Inventario gr├ífico interactivo con drag-and-drop visual
- **Instancias en Declares.bas:** `Inventario`, `InvBanco(1)`, `InvComUsu`, `InvOroComUsu(2)`, `InvOfferComUsu(1)`, `InvComNpc`, `InvLingosHerreria(4)`, `InvMaderasCarpinteria(4)`
- **Tama├▒o:** ~29 KB

#### `clsGraphicalButton.cls`
- **Funci├│n:** Botones gr├íficos personalizados con estados (normal, hover, click)
- **Tama├▒o:** ~7.2 KB

#### `clsDialogs.cls`
- **Funci├│n:** Sistema de di├ílogos modales del juego
- **Tama├▒o:** ~15 KB

#### `clsCustomKeys.cls`
- **Funci├│n:** Sistema de teclas personalizables por el usuario
- **Tama├▒o:** ~16 KB

#### `clsCustomMessages.cls`
- **Funci├│n:** Mensajes personalizados r├ípidos configurables por el usuario
- **Tama├▒o:** ~8.9 KB

#### `clsGuildDlg.cls`
- **Funci├│n:** Di├ílogos espec├¡ficos del sistema de guilds/clanes
- **Tama├▒o:** ~6.5 KB

#### `clsFormMovementManager.cls`
- **Funci├│n:** Gesti├│n de movimiento y posicionamiento de formularios flotantes
- **Tama├▒o:** ~3.5 KB

#### `clsForum.cls`
- **Funci├│n:** Sistema de foro/noticias de guild integrado
- **Tama├▒o:** ~7.7 KB

#### `clsIniManager.cls`
- **Funci├│n:** Lectura y escritura de archivos INI de configuraci├│n
- **Tama├▒o:** ~21 KB (compartida con servidor)

---

### M├│dulos auxiliares

#### `MainTimer.cls` (clsTimer)
- **Funci├│n:** Timer principal del cliente para actualizaciones de estado y animaciones
- **Tama├▒o:** ~9.3 KB

#### `modConsole.bas`
- **Funci├│n:** Consola de depuraci├│n interna
- **Tama├▒o:** ~9.1 KB

#### `modScreenCapture.bas`
- **Funci├│n:** Sistema de capturas de pantalla del juego
- **Tama├▒o:** ~26 KB

---

### Formularios de juego (interfaces)

| Formulario | Prop├│sito | Tama├▒o |
|---|---|---|
| `frmMain` | Ventana principal del juego ÔÇö mapa, chat, stats | ~88 KB |
| `frmConnect` | Pantalla de conexi├│n/login | ~8 KB |
| `frmCrearPersonaje` | Creaci├│n de nuevo personaje | ~61 KB |
| `frmSkills3` | Panel de habilidades y skills | ~51 KB |
| `frmPanelGm` | Panel de administraci├│n GM | ~64 KB |
| `FrmEstadisticas` | Estad├¡sticas del personaje | ~29 KB |
| `frmMapa` | Minimapa del mundo | ~4.9 KB |
| `frmComerciar` | Comercio con NPCs | ~16 KB |
| `frmComerciarUsu` | Comercio entre usuarios | ~28 KB |
| `frmHerrero` | Herrer├¡a (forja de armas/armaduras) | ~32 KB |
| `frmCarp` | Carpinter├¡a (construcci├│n de objetos) | ~28 KB |
| `frmBancoObj` | Banco de objetos | ~15 KB |
| `frmParty` | Sistema de grupos (party) | ~14 KB |
| `frmGuildBrief` | Lista de clanes | ~17 KB |
| `frmGuildDetails` | Detalles de un clan | ~11 KB |
| `frmGuildLeader` | Panel de l├¡der de clan | ~17 KB |
| `frmGuildMember` | Vista de miembro de clan | ~6.9 KB |
| `frmGuildFoundation` | Fundaci├│n de nuevo clan | ~4.5 KB |
| `frmGuildAdm` | Administraci├│n de clan | ~5.8 KB |
| `frmGuildNews` | Noticias del clan | ~4.8 KB |
| `frmGuildURL` | URL del clan | ~4.5 KB |
| `frmSolicitud` | Solicitud de ingreso a clan | ~5.1 KB |
| `frmUserRequest` | Peticiones de usuario | ~4.4 KB |
| `frmPeaceProp` | Propuesta de paz entre clanes | ~5.6 KB |
| `frmRetos` | Retos PvP 1vs1 / 2vs2 | ~7.2 KB |
| `frmCanjes` | Canje de puntos por ├¡tems | ~7.3 KB |
| `frmOpciones` | Configuraci├│n de opciones | ~14 KB |
| `frmCustomKeys` | Configuraci├│n de teclas | ~22 KB |
| `frmMessageTxt` | Mensajes de texto flotantes | ~10 KB |
| `frmNewPassword` | Cambio de contrase├▒a | ~4.7 KB |
| `frmPasswd` | Formulario de contrase├▒a | ~7.8 KB |
| `frmMenu` | Men├║ principal | ~3.8 KB |
| `frmMSG` | Mensaje de sistema | ~6.8 KB |
| `frmMensaje` | Mensaje privado | ~4 KB |
| `frmForo` | Foro del clan | ~17 KB |
| `frmKeypad` | Teclado num├®rico | ~15 KB |
| `frmEntrenador` | NPC entrenador de mascotas | ~5.2 KB |
| `frmSpawnList` | Lista de spawns (GM) | ~3.4 KB |
| `frmEligeAlineacion` | Elecci├│n de alineaci├│n | ~5.8 KB |
| `frmCantidad` | Input de cantidad (gen├®rico) | ~5.2 KB |
| `frmCantidadDrop` | Input de cantidad (tirar objeto) | ~5.2 KB |
| `frmCargando` | Pantalla de carga | ~2 KB |
| `frmCharInfo` | Informaci├│n del personaje | ~14 KB |
| `frmCommet` | Comentarios/feedback | ~5.8 KB |
| `FrmControl` | Control interno (GM) | ~5.4 KB |
| `frmScreenshots` | Gesti├│n de capturas | ~0.9 KB |

---

## Servidor (`server/Codigo/`)

### M├│dulos de n├║cleo

#### `Declares.bas` (Declaraciones)
- **Funci├│n:** Todas las declaraciones globales, tipos, enumeraciones y constantes del servidor
- **Contenido verificado:** Enumeraciones `eClass`, `eRaza`, `eGenero`, `eCiudad`, `PlayerType`, `ePrivileges`, `eTrigger`, `FXIDs`, `iMinerales`; constantes de ├¡tems especiales (embarcaciones, armas m├ígicas)
- **Tama├▒o:** ~51 KB

#### `General.bas`
- **Funci├│n:** Funciones de utilidad general del servidor (asignaci├│n de cuerpos desnudos por raza/g├®nero, funciones auxiliares)
- **Tama├▒o:** ~64 KB

#### `GameLogic.bas` (Extra)
- **Funci├│n:** L├│gica de juego principal (eventos del mundo, clima, efectos de ├írea)
- **Tama├▒o:** ~83 KB

#### `FileIO.bas` (ES)
- **Funci├│n:** I/O completo de archivos ÔÇö persistencia de personajes, carga de datos de juego, lectura de configuraci├│n
- **Estructura interna verificada:** Tipo `ConfigHAO` que mapea `Configuracion.ini`; arrays `ExpForLvl()` para tabla de experiencia por nivel
- **Tama├▒o:** ~122 KB

#### `Matematicas.bas`
- **Funci├│n:** Funciones matem├íticas auxiliares (n├║meros aleatorios, estad├¡sticas)
- **Tama├▒o:** ~1.8 KB

---

### M├│dulos de red

#### `TCP.bas`
- **Funci├│n:** Gesti├│n de conexiones y operaciones de alto nivel sobre sockets
- **Tama├▒o:** ~72 KB (incluye l├│gica de cabezas/apariencias de personaje)

#### `wsksock.bas` (WSKSOCK)
- **Funci├│n:** Capa de socket de bajo nivel (abstracci├│n de WinSock API)
- **Tama├▒o:** ~41 KB

#### `wskapiAO.bas`
- **Funci├│n:** API de socket espec├¡fica para Argentum Online sobre `wsksock`
- **Tama├▒o:** ~22 KB

#### `Protocol.bas`
- **Funci├│n:** Protocolo completo servidor ÔÇö decodifica todos los paquetes entrantes de clientes y ejecuta acciones de juego
- **Tama├▒o:** ~903 KB ÔÇö el m├│dulo m├ís grande del proyecto

#### `modSendData.bas`
- **Funci├│n:** Env├¡o centralizado de datos a clientes (unicast, broadcast por ├írea/mapa)
- **Tama├▒o:** ~35 KB

#### `Queue.bas`
- **Funci├│n:** Cola gen├®rica de mensajes para el sistema de red
- **Tama├▒o:** ~2.6 KB

#### `clsByteQueue.cls` / `clsByteBuffer.cls`
- **Funci├│n:** Serializaci├│n binaria del protocolo (mismo mecanismo que el cliente)
- **Tama├▒os:** ~45 KB / ~6 KB

---

### Gesti├│n de usuarios

#### `Modulo_UsUaRiOs.bas` (UsUaRiOs)
- **Funci├│n:** Gesti├│n completa de usuarios conectados (conexi├│n, desconexi├│n, estado, posici├│n)
- **Tama├▒o:** ~150 KB ÔÇö segundo m├│dulo m├ís grande del servidor

#### `Characters.bas`
- **Funci├│n:** Operaciones sobre personajes (carga, guardado, creaci├│n)
- **Tama├▒o:** ~2.3 KB

#### `modUserRecords.bas`
- **Funci├│n:** Records y estad├¡sticas hist├│ricas de usuarios
- **Tama├▒o:** ~6.1 KB

#### `modPrivateMessages.bas`
- **Funci├│n:** Sistema de mensajes privados entre usuarios
- **Tama├▒o:** ~9.9 KB

---

### Sistema de combate

#### `SistemaCombate.bas`
- **Funci├│n:** Motor completo de combate PvP y PvE
- **Verificado:** Constantes `MAXDISTANCIAARCO = 18`, `MAXDISTANCIAMAGIA = 18`; modificadores de clase le├¡dos desde `Balance.dat` (desde 2008)
- **Autores:** Dise├▒o original por Pablo M├írquez; correcci├│n por Gerardo Saiz
- **Tama├▒o:** ~124 KB

#### `modHechizos.bas`
- **Funci├│n:** Sistema de hechizos (lanzamiento, efectos, cooldowns)
- **Hechizos especiales:** Apocalipsis (├¡ndice 25), Descarga el├®ctrica (├¡ndice 23)
- **Tama├▒o:** ~127 KB

#### `modNuevoTimer.bas`
- **Funci├│n:** Game loop con timers y gesti├│n de intervalos de acci├│n
- **Intervalos verificados:** 9 tipos configurables (ataque, flechas, hechizos, ├¡tems, pociones, combos)
- **Tama├▒o:** ~20 KB

---

### IA y NPCs

#### `AI_NPC.bas` (AI)
- **Funci├│n:** Inteligencia artificial de NPCs (movimiento, detecci├│n de enemigos, ataque)
- **Tama├▒o:** ~62 KB

#### `MODULO_NPCs.bas` (NPCs)
- **Funci├│n:** Gesti├│n de instancias de NPCs en el mundo (spawn, deathspawn, actualizaci├│n)
- **Tama├▒o:** ~50 KB

#### `Modulo_InventANDobj.bas` (InvNpc)
- **Funci├│n:** Inventario de NPCs y gesti├│n de objetos en el suelo
- **Tama├▒o:** ~14 KB

#### `PathFinding.bas`
- **Funci├│n:** Algoritmo de pathfinding para NPCs (probablemente A* o BFS)
- **Tama├▒o:** ~12 KB

#### `praetorians.bas` (PraetoriansCoopNPC)
- **Funci├│n:** L├│gica cooperativa de NPCs Pretorianos
- **Tama├▒o:** ~40 KB

#### `clsClanPretoriano.cls`
- **Funci├│n:** Clase completa de clan Pretoriano con IA propia
- **Autores:** Dise├▒o original por Mariano Barrou (El Oso); redise├▒o por ZaMa
- **Tama├▒o:** ~118 KB

---

### Sistemas de juego

#### `InvUsuario.bas`
- **Funci├│n:** Gesti├│n completa del inventario del usuario (a├▒adir, quitar, mover, usar objetos)
- **Tama├▒o:** ~109 KB

#### `Trabajo.bas`
- **Funci├│n:** Sistema de trabajos (pesca, tala, miner├¡a, carpinter├¡a, herrer├¡a)
- **Tama├▒o:** ~128 KB

#### `Comercio.bas` (modSistemaComercio)
- **Funci├│n:** Comercio de usuarios con NPCs comerciantes
- **Tama├▒o:** ~15 KB

#### `mdlCOmercioConUsuario.bas`
- **Funci├│n:** Comercio directo entre dos usuarios
- **Tama├▒o:** ~22 KB

#### `modBanco.bas`
- **Funci├│n:** Sistema de banco de objetos (dep├│sito y retiro)
- **Tama├▒o:** ~15 KB

#### `modGuilds.bas`
- **Funci├│n:** Sistema completo de clanes/guilds
- **Tama├▒o:** ~85 KB

#### `clsClan.cls`
- **Funci├│n:** Clase de clan con miembros, rangos, guerras, alianzas
- **Tama├▒o:** ~33 KB

#### `mdParty.bas`
- **Funci├│n:** Sistema de grupos de juego (party)
- **Tama├▒o:** ~24 KB

#### `clsParty.cls`
- **Funci├│n:** Clase de party con gesti├│n de miembros y distribuci├│n de experiencia
- **Tama├▒o:** ~23 KB

#### `ModFacciones.bas`
- **Funci├│n:** Sistema de facciones Real/Caos (estado del jugador, zonas de facci├│n)
- **Tama├▒o:** ~41 KB

#### `Mod_Retos1vs1.bas`
- **Funci├│n:** Sistema de retos 1 vs 1 en arena
- **Tama├▒o:** ~26 KB

#### `Mod_Retos2vs2.bas`
- **Funci├│n:** Sistema de retos 2 vs 2 en arena
- **Tama├▒o:** ~39 KB

#### `Mod_Cofres.bas`
- **Funci├│n:** Sistema de cofres (ba├║les con acceso controlado)
- **Tama├▒o:** ~5.4 KB

#### `modForum.bas`
- **Funci├│n:** Foro de guild (noticias, posts de miembros)
- **Tama├▒o:** ~17 KB

#### `ModAreas.bas`
- **Funci├│n:** Zonas especiales del mapa (seguras, de combate, de facci├│n)
- **Tama├▒o:** ~23 KB

#### `modInvisibles.bas`
- **Funci├│n:** Gesti├│n de personajes en estado de invisibilidad
- **Tama├▒o:** ~1.6 KB

#### `History.bas`
- **Funci├│n:** Historial de acciones del juego (log de eventos)
- **Tama├▒o:** ~5.5 KB

#### `Statistics.bas`
- **Funci├│n:** Sistema de estad├¡sticas del servidor (jugadores conectados, acciones, etc.)
- **Tama├▒o:** ~17 KB

---

### Administraci├│n y seguridad

#### `Admin.bas`
- **Funci├│n:** Comandos de administraci├│n para Game Masters (ban, kick, teleport, spawn, etc.)
- **Tama├▒o:** ~22 KB

#### `modCentinela.bas`
- **Funci├│n:** Sistema anti-macro ÔÇö desaf├¡a a usuarios que trabajan sin interacci├│n humana
- **Autores:** ImperiumAO (Barrin), Alkon AO (Juan Mart├¡n Sotuyo Dodero), ZaMa
- **Tama├▒o:** ~27 KB

#### `SecurityIp.bas`
- **Funci├│n:** Baneos por IP, control de multi-login, registro de IPs sospechosas
- **Tama├▒o:** ~14 KB

#### `clsAntiMassClon.cls`
- **Funci├│n:** Prevenci├│n de clonaci├│n masiva de personajes
- **Tama├▒o:** ~2.9 KB

#### `Acciones.bas`
- **Funci├│n:** Acciones del jugador (comandos de texto, emotes, interacciones)
- **Tama├▒o:** ~17 KB

#### `mod_DragAndDrop.bas`
- **Funci├│n:** Gesti├│n del sistema de arrastrar y soltar objetos en el mapa
- **Tama├▒o:** ~14 KB

---

### Clases de utilidad

#### `clsIniManager.cls` / `clsIniReader.cls`
- **Funci├│n:** Lectura y escritura de archivos INI de configuraci├│n
- **Tama├▒os:** ~21 KB / ~16 KB

#### `ModCola.cls` (cCola)
- **Funci├│n:** Cola gen├®rica de objetos
- **Tama├▒o:** ~6.7 KB

#### `cColaArray.cls`
- **Funci├│n:** Cola basada en array
- **Tama├▒o:** ~3.4 KB

#### `ConsultasPopulares.cls`
- **Funci├│n:** Cache de consultas frecuentes al servidor
- **Tama├▒o:** ~8 KB

#### `clsLimpiarMundo.cls`
- **Funci├│n:** Limpieza peri├│dica del mundo (objetos ca├¡dos, mapas espec├¡ficos)
- **Tama├▒o:** ~5 KB

#### `clsMapSoundManager.cls` (SoundMapInfo)
- **Funci├│n:** Gesti├│n de sonido por mapa (MIDI/WAV seg├║n zona)
- **Tama├▒o:** ~5.4 KB

#### `clsEstadisticasIPC.cls`
- **Funci├│n:** Estad├¡sticas de comunicaci├│n inter-proceso
- **Tama├▒o:** ~4.4 KB

#### `clsdicc.cls` (diccionario)
- **Funci├│n:** Diccionario gen├®rico para b├║squedas r├ípidas
- **Tama├▒o:** ~4.5 KB

#### `modHexaStrings.bas`
- **Funci├│n:** Conversi├│n de datos a representaci├│n hexadecimal (para depuraci├│n y logs)
- **Tama├▒o:** ~2.8 KB

#### `Modulo_SysTray.bas` (SysTray)
- **Funci├│n:** Integraci├│n con la bandeja del sistema de Windows (minimizar a systray)
- **Tama├▒o:** ~3.7 KB

---

### Formularios del servidor

| Formulario | Prop├│sito |
|---|---|
| `frmMain` | Panel de control principal del servidor (ventana de administraci├│n) |
| `frmServidor` | Consola de estado del servidor |
| `FrmInterv` | Monitor de intervenci├│n/moderaci├│n |
| `FrmStat` | Estad├¡sticas en tiempo real |
| `frmAdmin` | Panel de administraci├│n avanzado |
| `frmCargando` | Pantalla de carga inicial |
| `frmConID` | Conexi├│n por ID espec├¡fico |
| `frmDebugNpc` | Depuraci├│n de NPCs (herramienta de desarrollo) |
| `frmTrafic` | Monitor de tr├ífico de red |
| `frmUserList` | Lista de usuarios conectados |

---

## World Editor (`WE HISPANO AO/`)

| Componente | Descripci├│n |
|---|---|
| `WorldEditor.exe` | Aplicaci├│n de edici├│n de mapas (binario, sin c├│digo fuente en el repo) |
| `libreria de render.exe` | Librer├¡a auxiliar de renderizado para el editor |
| `zlib.dll` | Compresi├│n de datos (posiblemente para compresi├│n de mapas) |
| `Dat/` | Datos de referencia de objetos y NPCs para el editor |
| `Graficos/` | Gr├íficos de referencia para visualizaci├│n en el editor |
| `INIT/` | Configuraci├│n del editor |
| `Maps/` | Mapas de trabajo del editor |
| `MIDI/ / Wavs/` | Audio de referencia |

> **Nota:** El World Editor no tiene c├│digo fuente en este repositorio. Solo est├í disponible el ejecutable binario.
