# Componentes por Módulo — Hispano-AO

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Cliente (`client/CODIGO/`)

### Módulos de núcleo

#### `Application.bas`
- **Función:** Módulo de arranque de la aplicación (`Sub Main`)
- **Responsabilidades:** Detecta si la aplicación está activa en primer plano (`IsAppActive`). Entry point de la aplicación VB6.
- **Tamaño:** ~1.8 KB

#### `Declares.bas` (Mod_Declaraciones)
- **Función:** Declaraciones globales, tipos de datos y constantes del cliente
- **Contenido verificado:** Tipos `Canjeo`, instancias globales de clases (`Audio`, `Inventario`, `SurfaceDB`, etc.), constantes de intervalos de acción, constantes de sonido
- **Tamaño:** ~31 KB

#### `General.bas` (Mod_General)
- **Función:** Lógica general de juego del lado cliente (renderizado de personajes, movimiento, interacciones)
- **Tamaño:** ~57 KB (módulo de lógica principal del cliente)

#### `GameIni.bas`
- **Función:** Carga de configuración inicial del juego desde archivos `.ini` y `.dat`
- **Tamaño:** ~2 KB

#### `PrevInstance.bas`
- **Función:** Previene múltiples instancias simultáneas del cliente
- **Tamaño:** ~4.3 KB

#### `Resolution.bas`
- **Función:** Gestión de resolución de pantalla y modos de visualización
- **Tamaño:** ~8.2 KB

#### `ModAreas.bas`
- **Función:** Lógica de áreas del mapa (zonas seguras, zonas de combate, etc.)
- **Tamaño:** ~2.6 KB

---

### Módulos de red

#### `TCP.bas` (Mod_TCP)
- **Función:** Gestión de la conexión TCP con el servidor
- **Tecnología:** `CSWSK32.OCX` (Crescent WinSock)
- **Tamaño:** ~2.4 KB

#### `Protocol.bas`
- **Función:** Serialización de todos los comandos del cliente → servidor
- **Tamaño:** ~420 KB — el módulo más grande del cliente; contiene una función de envío por cada tipo de acción posible del jugador
- **Nota:** El tamaño refleja la amplitud del protocolo (movimiento, combate, comercio, guilds, hechizos, etc.)

#### `ProtocolCmdParse.bas`
- **Función:** Parser de todos los comandos recibidos del servidor
- **Tamaño:** ~108 KB — deserializa cada tipo de respuesta del servidor y actualiza el estado del juego
- **Nota:** Complemento directo de `Protocol.bas` del servidor

#### `clsByteQueue.cls`
- **Función:** Cola FIFO de bytes para serialización binaria del protocolo
- **Uso:** Instanciada como `incomingData` y `outgoingData` en `Declares.bas`
- **Tamaño:** ~45 KB (compartida con el servidor: mismo archivo)

---

### Motor gráfico

#### `TileEngine.bas` (Mod_TileEngine)
- **Función:** Motor de renderizado 2D basado en tiles
- **Tecnología:** GDI Windows con DIBSections (sin DirectX para gráficos)
- **Tamaño:** ~89 KB — núcleo del renderizado del mundo
- **Responsabilidades:** Dibujado de capas de tiles, personajes, NPCs, efectos visuales, animaciones

#### `cDIBSection.cls`
- **Función:** Encapsula una sección DIB (Device-Independent Bitmap) de Windows para renderizado eficiente
- **Tamaño:** ~19 KB

#### `clsSurfaceManager.cls`
- **Función:** Interfaz abstracta para manejo de superficies gráficas
- **Tamaño:** ~2.8 KB

#### `clsSurfaceManStatic.cls`
- **Función:** Gestión de superficies estáticas (tiles de terreno, objetos decorativos)
- **Tamaño:** ~13 KB

#### `clsSurfaceManDyn.cls`
- **Función:** Gestión de superficies dinámicas (personajes, NPCs en movimiento)
- **Tamaño:** ~16 KB

#### `Carteles.bas`
- **Función:** Sistema de carteles sobre las cabezas de los personajes (chat, nombres, daño)
- **Tamaño:** ~3.3 KB

#### `Mod_rDamage.bas`
- **Función:** Módulo de renderizado de daño visual (números flotantes de daño recibido/causado)
- **Tamaño:** ~5.5 KB

---

### Sistema de audio

#### `clsAudio.cls`
- **Función:** Gestión completa de audio del juego
- **Formatos soportados:** WAV (efectos), MIDI (música de fondo), MP3 (música alternativa)
- **Tecnología:** DirectX (quartz.dll — DirectShow), posiblemente dx8vb.dll
- **Tamaño:** ~41 KB

---

### Sistema de seguridad y cifrado

#### `clsCripto.cls`
- **Función:** Cifrado y descifrado de datos para comunicación segura
- **Tamaño:** ~52 KB — módulo de criptografía significativo
- **Nota:** El uso exacto en el protocolo de red requiere análisis adicional del código

---

### Clases de interfaz y UI

#### `clsGrapchicalInventory.cls`
- **Función:** Inventario gráfico interactivo con drag-and-drop visual
- **Instancias en Declares.bas:** `Inventario`, `InvBanco(1)`, `InvComUsu`, `InvOroComUsu(2)`, `InvOfferComUsu(1)`, `InvComNpc`, `InvLingosHerreria(4)`, `InvMaderasCarpinteria(4)`
- **Tamaño:** ~29 KB

#### `clsGraphicalButton.cls`
- **Función:** Botones gráficos personalizados con estados (normal, hover, click)
- **Tamaño:** ~7.2 KB

#### `clsDialogs.cls`
- **Función:** Sistema de diálogos modales del juego
- **Tamaño:** ~15 KB

#### `clsCustomKeys.cls`
- **Función:** Sistema de teclas personalizables por el usuario
- **Tamaño:** ~16 KB

#### `clsCustomMessages.cls`
- **Función:** Mensajes personalizados rápidos configurables por el usuario
- **Tamaño:** ~8.9 KB

#### `clsGuildDlg.cls`
- **Función:** Diálogos específicos del sistema de guilds/clanes
- **Tamaño:** ~6.5 KB

#### `clsFormMovementManager.cls`
- **Función:** Gestión de movimiento y posicionamiento de formularios flotantes
- **Tamaño:** ~3.5 KB

#### `clsForum.cls`
- **Función:** Sistema de foro/noticias de guild integrado
- **Tamaño:** ~7.7 KB

#### `clsIniManager.cls`
- **Función:** Lectura y escritura de archivos INI de configuración
- **Tamaño:** ~21 KB (compartida con servidor)

---

### Módulos auxiliares

#### `MainTimer.cls` (clsTimer)
- **Función:** Timer principal del cliente para actualizaciones de estado y animaciones
- **Tamaño:** ~9.3 KB

#### `modConsole.bas`
- **Función:** Consola de depuración interna
- **Tamaño:** ~9.1 KB

#### `modScreenCapture.bas`
- **Función:** Sistema de capturas de pantalla del juego
- **Tamaño:** ~26 KB

---

### Formularios de juego (interfaces)

| Formulario | Propósito | Tamaño |
|---|---|---|
| `frmMain` | Ventana principal del juego — mapa, chat, stats | ~88 KB |
| `frmConnect` | Pantalla de conexión/login | ~8 KB |
| `frmCrearPersonaje` | Creación de nuevo personaje | ~61 KB |
| `frmSkills3` | Panel de habilidades y skills | ~51 KB |
| `frmPanelGm` | Panel de administración GM | ~64 KB |
| `FrmEstadisticas` | Estadísticas del personaje | ~29 KB |
| `frmMapa` | Minimapa del mundo | ~4.9 KB |
| `frmComerciar` | Comercio con NPCs | ~16 KB |
| `frmComerciarUsu` | Comercio entre usuarios | ~28 KB |
| `frmHerrero` | Herrería (forja de armas/armaduras) | ~32 KB |
| `frmCarp` | Carpintería (construcción de objetos) | ~28 KB |
| `frmBancoObj` | Banco de objetos | ~15 KB |
| `frmParty` | Sistema de grupos (party) | ~14 KB |
| `frmGuildBrief` | Lista de clanes | ~17 KB |
| `frmGuildDetails` | Detalles de un clan | ~11 KB |
| `frmGuildLeader` | Panel de líder de clan | ~17 KB |
| `frmGuildMember` | Vista de miembro de clan | ~6.9 KB |
| `frmGuildFoundation` | Fundación de nuevo clan | ~4.5 KB |
| `frmGuildAdm` | Administración de clan | ~5.8 KB |
| `frmGuildNews` | Noticias del clan | ~4.8 KB |
| `frmGuildURL` | URL del clan | ~4.5 KB |
| `frmSolicitud` | Solicitud de ingreso a clan | ~5.1 KB |
| `frmUserRequest` | Peticiones de usuario | ~4.4 KB |
| `frmPeaceProp` | Propuesta de paz entre clanes | ~5.6 KB |
| `frmRetos` | Retos PvP 1vs1 / 2vs2 | ~7.2 KB |
| `frmCanjes` | Canje de puntos por ítems | ~7.3 KB |
| `frmOpciones` | Configuración de opciones | ~14 KB |
| `frmCustomKeys` | Configuración de teclas | ~22 KB |
| `frmMessageTxt` | Mensajes de texto flotantes | ~10 KB |
| `frmNewPassword` | Cambio de contraseña | ~4.7 KB |
| `frmPasswd` | Formulario de contraseña | ~7.8 KB |
| `frmMenu` | Menú principal | ~3.8 KB |
| `frmMSG` | Mensaje de sistema | ~6.8 KB |
| `frmMensaje` | Mensaje privado | ~4 KB |
| `frmForo` | Foro del clan | ~17 KB |
| `frmKeypad` | Teclado numérico | ~15 KB |
| `frmEntrenador` | NPC entrenador de mascotas | ~5.2 KB |
| `frmSpawnList` | Lista de spawns (GM) | ~3.4 KB |
| `frmEligeAlineacion` | Elección de alineación | ~5.8 KB |
| `frmCantidad` | Input de cantidad (genérico) | ~5.2 KB |
| `frmCantidadDrop` | Input de cantidad (tirar objeto) | ~5.2 KB |
| `frmCargando` | Pantalla de carga | ~2 KB |
| `frmCharInfo` | Información del personaje | ~14 KB |
| `frmCommet` | Comentarios/feedback | ~5.8 KB |
| `FrmControl` | Control interno (GM) | ~5.4 KB |
| `frmScreenshots` | Gestión de capturas | ~0.9 KB |

---

## Servidor (`server/Codigo/`)

### Módulos de núcleo

#### `Declares.bas` (Declaraciones)
- **Función:** Todas las declaraciones globales, tipos, enumeraciones y constantes del servidor
- **Contenido verificado:** Enumeraciones `eClass`, `eRaza`, `eGenero`, `eCiudad`, `PlayerType`, `ePrivileges`, `eTrigger`, `FXIDs`, `iMinerales`; constantes de ítems especiales (embarcaciones, armas mágicas)
- **Tamaño:** ~51 KB

#### `General.bas`
- **Función:** Funciones de utilidad general del servidor (asignación de cuerpos desnudos por raza/género, funciones auxiliares)
- **Tamaño:** ~64 KB

#### `GameLogic.bas` (Extra)
- **Función:** Lógica de juego principal (eventos del mundo, clima, efectos de área)
- **Tamaño:** ~83 KB

#### `FileIO.bas` (ES)
- **Función:** I/O completo de archivos — persistencia de personajes, carga de datos de juego, lectura de configuración
- **Estructura interna verificada:** Tipo `ConfigHAO` que mapea `Configuracion.ini`; arrays `ExpForLvl()` para tabla de experiencia por nivel
- **Tamaño:** ~122 KB

#### `Matematicas.bas`
- **Función:** Funciones matemáticas auxiliares (números aleatorios, estadísticas)
- **Tamaño:** ~1.8 KB

---

### Módulos de red

#### `TCP.bas`
- **Función:** Gestión de conexiones y operaciones de alto nivel sobre sockets
- **Tamaño:** ~72 KB (incluye lógica de cabezas/apariencias de personaje)

#### `wsksock.bas` (WSKSOCK)
- **Función:** Capa de socket de bajo nivel (abstracción de WinSock API)
- **Tamaño:** ~41 KB

#### `wskapiAO.bas`
- **Función:** API de socket específica para Argentum Online sobre `wsksock`
- **Tamaño:** ~22 KB

#### `Protocol.bas`
- **Función:** Protocolo completo servidor — decodifica todos los paquetes entrantes de clientes y ejecuta acciones de juego
- **Tamaño:** ~903 KB — el módulo más grande del proyecto

#### `modSendData.bas`
- **Función:** Envío centralizado de datos a clientes (unicast, broadcast por área/mapa)
- **Tamaño:** ~35 KB

#### `Queue.bas`
- **Función:** Cola genérica de mensajes para el sistema de red
- **Tamaño:** ~2.6 KB

#### `clsByteQueue.cls` / `clsByteBuffer.cls`
- **Función:** Serialización binaria del protocolo (mismo mecanismo que el cliente)
- **Tamaños:** ~45 KB / ~6 KB

---

### Gestión de usuarios

#### `Modulo_UsUaRiOs.bas` (UsUaRiOs)
- **Función:** Gestión completa de usuarios conectados (conexión, desconexión, estado, posición)
- **Tamaño:** ~150 KB — segundo módulo más grande del servidor

#### `Characters.bas`
- **Función:** Operaciones sobre personajes (carga, guardado, creación)
- **Tamaño:** ~2.3 KB

#### `modUserRecords.bas`
- **Función:** Records y estadísticas históricas de usuarios
- **Tamaño:** ~6.1 KB

#### `modPrivateMessages.bas`
- **Función:** Sistema de mensajes privados entre usuarios
- **Tamaño:** ~9.9 KB

---

### Sistema de combate

#### `SistemaCombate.bas`
- **Función:** Motor completo de combate PvP y PvE
- **Verificado:** Constantes `MAXDISTANCIAARCO = 18`, `MAXDISTANCIAMAGIA = 18`; modificadores de clase leídos desde `Balance.dat` (desde 2008)
- **Autores:** Diseño original por Pablo Márquez; corrección por Gerardo Saiz
- **Tamaño:** ~124 KB

#### `modHechizos.bas`
- **Función:** Sistema de hechizos (lanzamiento, efectos, cooldowns)
- **Hechizos especiales:** Apocalipsis (índice 25), Descarga eléctrica (índice 23)
- **Tamaño:** ~127 KB

#### `modNuevoTimer.bas`
- **Función:** Game loop con timers y gestión de intervalos de acción
- **Intervalos verificados:** 9 tipos configurables (ataque, flechas, hechizos, ítems, pociones, combos)
- **Tamaño:** ~20 KB

---

### IA y NPCs

#### `AI_NPC.bas` (AI)
- **Función:** Inteligencia artificial de NPCs (movimiento, detección de enemigos, ataque)
- **Tamaño:** ~62 KB

#### `MODULO_NPCs.bas` (NPCs)
- **Función:** Gestión de instancias de NPCs en el mundo (spawn, deathspawn, actualización)
- **Tamaño:** ~50 KB

#### `Modulo_InventANDobj.bas` (InvNpc)
- **Función:** Inventario de NPCs y gestión de objetos en el suelo
- **Tamaño:** ~14 KB

#### `PathFinding.bas`
- **Función:** Algoritmo de pathfinding para NPCs (probablemente A* o BFS)
- **Tamaño:** ~12 KB

#### `praetorians.bas` (PraetoriansCoopNPC)
- **Función:** Lógica cooperativa de NPCs Pretorianos
- **Tamaño:** ~40 KB

#### `clsClanPretoriano.cls`
- **Función:** Clase completa de clan Pretoriano con IA propia
- **Autores:** Diseño original por Mariano Barrou (El Oso); rediseño por ZaMa
- **Tamaño:** ~118 KB

---

### Sistemas de juego

#### `InvUsuario.bas`
- **Función:** Gestión completa del inventario del usuario (añadir, quitar, mover, usar objetos)
- **Tamaño:** ~109 KB

#### `Trabajo.bas`
- **Función:** Sistema de trabajos (pesca, tala, minería, carpintería, herrería)
- **Tamaño:** ~128 KB

#### `Comercio.bas` (modSistemaComercio)
- **Función:** Comercio de usuarios con NPCs comerciantes
- **Tamaño:** ~15 KB

#### `mdlCOmercioConUsuario.bas`
- **Función:** Comercio directo entre dos usuarios
- **Tamaño:** ~22 KB

#### `modBanco.bas`
- **Función:** Sistema de banco de objetos (depósito y retiro)
- **Tamaño:** ~15 KB

#### `modGuilds.bas`
- **Función:** Sistema completo de clanes/guilds
- **Tamaño:** ~85 KB

#### `clsClan.cls`
- **Función:** Clase de clan con miembros, rangos, guerras, alianzas
- **Tamaño:** ~33 KB

#### `mdParty.bas`
- **Función:** Sistema de grupos de juego (party)
- **Tamaño:** ~24 KB

#### `clsParty.cls`
- **Función:** Clase de party con gestión de miembros y distribución de experiencia
- **Tamaño:** ~23 KB

#### `ModFacciones.bas`
- **Función:** Sistema de facciones Real/Caos (estado del jugador, zonas de facción)
- **Tamaño:** ~41 KB

#### `Mod_Retos1vs1.bas`
- **Función:** Sistema de retos 1 vs 1 en arena
- **Tamaño:** ~26 KB

#### `Mod_Retos2vs2.bas`
- **Función:** Sistema de retos 2 vs 2 en arena
- **Tamaño:** ~39 KB

#### `Mod_Cofres.bas`
- **Función:** Sistema de cofres (baúles con acceso controlado)
- **Tamaño:** ~5.4 KB

#### `modForum.bas`
- **Función:** Foro de guild (noticias, posts de miembros)
- **Tamaño:** ~17 KB

#### `ModAreas.bas`
- **Función:** Zonas especiales del mapa (seguras, de combate, de facción)
- **Tamaño:** ~23 KB

#### `modInvisibles.bas`
- **Función:** Gestión de personajes en estado de invisibilidad
- **Tamaño:** ~1.6 KB

#### `History.bas`
- **Función:** Historial de acciones del juego (log de eventos)
- **Tamaño:** ~5.5 KB

#### `Statistics.bas`
- **Función:** Sistema de estadísticas del servidor (jugadores conectados, acciones, etc.)
- **Tamaño:** ~17 KB

---

### Administración y seguridad

#### `Admin.bas`
- **Función:** Comandos de administración para Game Masters (ban, kick, teleport, spawn, etc.)
- **Tamaño:** ~22 KB

#### `modCentinela.bas`
- **Función:** Sistema anti-macro — desafía a usuarios que trabajan sin interacción humana
- **Autores:** ImperiumAO (Barrin), Alkon AO (Juan Martín Sotuyo Dodero), ZaMa
- **Tamaño:** ~27 KB

#### `SecurityIp.bas`
- **Función:** Baneos por IP, control de multi-login, registro de IPs sospechosas
- **Tamaño:** ~14 KB

#### `clsAntiMassClon.cls`
- **Función:** Prevención de clonación masiva de personajes
- **Tamaño:** ~2.9 KB

#### `Acciones.bas`
- **Función:** Acciones del jugador (comandos de texto, emotes, interacciones)
- **Tamaño:** ~17 KB

#### `mod_DragAndDrop.bas`
- **Función:** Gestión del sistema de arrastrar y soltar objetos en el mapa
- **Tamaño:** ~14 KB

---

### Clases de utilidad

#### `clsIniManager.cls` / `clsIniReader.cls`
- **Función:** Lectura y escritura de archivos INI de configuración
- **Tamaños:** ~21 KB / ~16 KB

#### `ModCola.cls` (cCola)
- **Función:** Cola genérica de objetos
- **Tamaño:** ~6.7 KB

#### `cColaArray.cls`
- **Función:** Cola basada en array
- **Tamaño:** ~3.4 KB

#### `ConsultasPopulares.cls`
- **Función:** Cache de consultas frecuentes al servidor
- **Tamaño:** ~8 KB

#### `clsLimpiarMundo.cls`
- **Función:** Limpieza periódica del mundo (objetos caídos, mapas específicos)
- **Tamaño:** ~5 KB

#### `clsMapSoundManager.cls` (SoundMapInfo)
- **Función:** Gestión de sonido por mapa (MIDI/WAV según zona)
- **Tamaño:** ~5.4 KB

#### `clsEstadisticasIPC.cls`
- **Función:** Estadísticas de comunicación inter-proceso
- **Tamaño:** ~4.4 KB

#### `clsdicc.cls` (diccionario)
- **Función:** Diccionario genérico para búsquedas rápidas
- **Tamaño:** ~4.5 KB

#### `modHexaStrings.bas`
- **Función:** Conversión de datos a representación hexadecimal (para depuración y logs)
- **Tamaño:** ~2.8 KB

#### `Modulo_SysTray.bas` (SysTray)
- **Función:** Integración con la bandeja del sistema de Windows (minimizar a systray)
- **Tamaño:** ~3.7 KB

---

### Formularios del servidor

| Formulario | Propósito |
|---|---|
| `frmMain` | Panel de control principal del servidor (ventana de administración) |
| `frmServidor` | Consola de estado del servidor |
| `FrmInterv` | Monitor de intervención/moderación |
| `FrmStat` | Estadísticas en tiempo real |
| `frmAdmin` | Panel de administración avanzado |
| `frmCargando` | Pantalla de carga inicial |
| `frmConID` | Conexión por ID específico |
| `frmDebugNpc` | Depuración de NPCs (herramienta de desarrollo) |
| `frmTrafic` | Monitor de tráfico de red |
| `frmUserList` | Lista de usuarios conectados |

---

## World Editor (`WE HISPANO AO/`)

| Componente | Descripción |
|---|---|
| `WorldEditor.exe` | Aplicación de edición de mapas (binario, sin código fuente en el repo) |
| `libreria de render.exe` | Librería auxiliar de renderizado para el editor |
| `zlib.dll` | Compresión de datos (posiblemente para compresión de mapas) |
| `Dat/` | Datos de referencia de objetos y NPCs para el editor |
| `Graficos/` | Gráficos de referencia para visualización en el editor |
| `INIT/` | Configuración del editor |
| `Maps/` | Mapas de trabajo del editor |
| `MIDI/ / Wavs/` | Audio de referencia |

> **Nota:** El World Editor no tiene código fuente en este repositorio. Solo está disponible el ejecutable binario.
