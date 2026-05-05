# Arquitectura Técnica — Hispano-AO

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Visión de alto nivel

Hispano-AO es un MMORPG 2D clásico con arquitectura **cliente-servidor** donde ambos extremos están implementados en **Visual Basic 6**. La comunicación ocurre exclusivamente por **TCP binario** sobre el puerto 7666.

```
┌──────────────────────────────────────────────────────────────────┐
│  CLIENTE (Hispano Online.exe)                                    │
│                                                                  │
│  ┌────────────┐  ┌──────────────┐  ┌──────────────────────┐     │
│  │ TileEngine │  │  clsAudio    │  │ Formularios VB6      │     │
│  │ (render 2D)│  │ (WAV/MIDI/MP3│  │ (frmMain, frmSkills, │     │
│  │ DIBSection │  │ via DirectX) │  │  frmGuild*, frmComerc│     │
│  └────────────┘  └──────────────┘  └──────────────────────┘     │
│         ↕                                    ↕                   │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │              clsByteQueue (cola binaria)                 │     │
│  │  Protocol.bas (envío) ↔ ProtocolCmdParse.bas (recepción)│     │
│  └─────────────────────────────────────────────────────────┘     │
│                        ↕ TCP/IP (CSWSK32.OCX)                    │
└──────────────────────────────────┬───────────────────────────────┘
                                   │ Puerto 7666
┌──────────────────────────────────▼───────────────────────────────┐
│  SERVIDOR (server9arreglado.exe)                                  │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │  Capa de red: wsksock.bas + wskapiAO.bas (hasta 550)    │     │
│  │  clsByteQueue por usuario (entrada) + modSendData (salida│     │
│  └─────────────────────────────────────────────────────────┘     │
│                        ↕                                         │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │  Protocol.bas — decodifica paquetes entrantes,          │     │
│  │  despacha a módulos de lógica                           │     │
│  └──────────────────────────┬──────────────────────────────┘     │
│                             ↓                                    │
│  ┌──────────┐ ┌──────────┐ ┌───────────┐ ┌────────────────┐     │
│  │GameLogic │ │SistemaCom│ │modHechizos│ │Trabajo.bas     │     │
│  │General   │ │bate.bas  │ │           │ │(pesca/tala/mina│     │
│  └──────────┘ └──────────┘ └───────────┘ └────────────────┘     │
│  ┌──────────┐ ┌──────────┐ ┌───────────┐ ┌────────────────┐     │
│  │AI_NPC    │ │modGuilds │ │ModFaccione│ │Mod_Retos1vs1/2 │     │
│  │          │ │          │ │s          │ │                │     │
│  └──────────┘ └──────────┘ └───────────┘ └────────────────┘     │
│                             ↕                                    │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │  FileIO.bas — persistencia (lectura/escritura de disco)  │     │
│  │  /Charfile/ (personajes) /Maps/ (mapas) /Dat/ (datos)   │     │
│  └─────────────────────────────────────────────────────────┘     │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │  modNuevoTimer (game loop) — ciclo base 50 ms           │     │
│  │  Gestiona intervalos de: ataque, hechizos, NPC AI,      │     │
│  │  regeneración, hambre/sed, veneno, invisibilidad, etc.  │     │
│  └─────────────────────────────────────────────────────────┘     │
└──────────────────────────────────────────────────────────────────┘
```

---

## Capa de red

### Cliente
- Usa `CSWSK32.OCX` (Crescent WinSock Control 32-bit) para la conexión TCP
- Los datos enviados y recibidos pasan por `clsByteQueue`, que encola bytes en un buffer
- `Protocol.bas` serializa los comandos a enviar al servidor
- `ProtocolCmdParse.bas` parsea los comandos recibidos del servidor y actualiza el estado del juego

### Servidor
- Gestiona hasta **550 conexiones simultáneas** (configurable en `Server.ini → MaxUsers`)
- Cada usuario tiene su propio índice (`UserIndex`) en el array global `UserList()`
- `wsksock.bas` provee la capa de socket de bajo nivel
- `wskapiAO.bas` implementa la API de alto nivel sobre los sockets
- `modSendData.bas` centraliza el envío de datos a clientes
- La verificación MD5 del cliente está soportada (`[MD5Hush]` en `Server.ini`)

---

## Protocolo de comunicación

El protocolo es **binario propietario**, usando `clsByteQueue` para empaquetar/desempaquetar:

- **Formato:** Big-endian implícito, sin delimitadores de frame explícitos visibles en el código
- Cada paquete comienza con un byte de **opcode** que identifica el tipo de comando
- El parser en `ProtocolCmdParse.bas` (cliente) y `Protocol.bas` (servidor) despacha según el opcode
- No hay capa de seguridad TLS — el tráfico es en texto/binario plano
- La capa de cifrado (`clsCripto.cls`) existe en el cliente pero su alcance real en el protocolo requiere análisis adicional

---

## Motor gráfico (cliente)

- **Renderizado:** GDI + DIBSections (`cDIBSection.cls`)
- `clsSurfaceManStatic` — superficies estáticas (tiles de fondo, objetos fijos)
- `clsSurfaceManDyn` — superficies dinámicas (personajes, NPCs, proyectiles)
- `clsSurfaceManager` — interfaz abstracta sobre los dos managers anteriores
- `TileEngine.bas` (~89 KB) contiene la lógica completa de renderizado de tiles
- El juego usa graficos empaquetados referenciados por índices (`Graficos.ind`, `graficos.ini`)
- **No usa DirectX para gráficos** — usa GDI puro (DIBSection sobre HDC)
- DirectX solo se usa para **audio** (`dx7vb.dll`, `dx8vb.dll`)

---

## Sistema de audio (cliente)

- Implementado en `clsAudio.cls` (~41 KB)
- Soporta tres formatos:
  - **WAV** — efectos de sonido (directorio `WAV/`)
  - **MIDI** — música de fondo (directorio `MIDI/`)
  - **MP3** — música de alta calidad (directorio `MP3/`)
- Usa DirectX (`quartz.dll` — DirectShow) y posiblemente `dx8vb.dll`
- Los sonidos 3D se identifican con la constante `NO_3D_SOUND = 0` (coordenadas especiales)

---

## Sistema de mapas

Tres archivos por mapa, todos binarios o texto estructurado:

| Extensión | Tamaño típico | Contenido |
|---|---|---|
| `.map` | 30–50 KB | Capa gráfica de tiles (terreno, objetos decorativos) |
| `.dat` | ~260 bytes | Metadatos: nombre del mapa, música, zona segura, lluvia, etc. |
| `.inf` | 10–15 KB | Capa lógica: triggers, spawns de NPCs, posiciones de objetos dinámicos |

**Mapas especiales configurados (verificado en `Configuracion.ini`):**

| Mapa | Número | Uso |
|---|---|---|
| Mapa GM | 90 | Zona de Game Masters |
| Arena 1vs1 | 73 | Retos 1 vs 1 |
| Arena 2vs2 | 74 | Retos 2 vs 2 |
| Prisión | 89 | Destino de jugadores baneados/encarcelados (pos 75-47) |
| Libertad | 1 | Punto de libertad (pos 43-60) |
| Mapa Real | 3 | Zona facción Real |
| Mapa Caos | 4 | Zona facción Caos |
| Mapa Pretoriano | 88 | Zona de clanes Pretorianos |

---

## Sistema de persistencia

Todo el I/O de datos pasa por `FileIO.bas` (módulo `ES`):

### Personajes (`/Charfile/`)
- Un archivo binario por personaje
- Almacena: stats, inventario, posición, estado, hechizos aprendidos, habilidades, etc.
- La persistencia se activa cada `IntervaloGuardarUsuarios = 180` segundos (configurable)

### Datos de juego (`/Dat/`)
- Cargados al inicio del servidor, mantenidos en memoria
- `NPCs.dat` → array de estructuras NPC
- `obj.dat` → array de definiciones de objetos/ítems
- `Hechizos.dat` → array de hechizos con efectos y requerimientos

### Mapas (`/Maps/`)
- Los `.map` e `.inf` se cargan en memoria al inicio
- Los `.dat` contienen metadatos por mapa

---

## Game loop y temporización

El game loop del servidor está implementado en `modNuevoTimer.bas`:

- Ciclo base de **50 ms** (`IntervaloTimerExec`)
- Los intervalos por acción se cargan desde `Configuracion.ini` y `Server.ini`
- El servidor mantiene un contador de "tolerancia" (`Tolerancia_FailIntervalo = 7`) para acciones en intervalos incorrectos

**Intervalos críticos (valores por defecto de `Server.ini`):**

| Acción | Intervalo |
|---|---|
| Movimiento de NPC | 200 ms |
| IA de NPC | 380 ms |
| NPC puede atacar | 1600 ms |
| Usuario puede atacar | 1500 ms |
| Usuario puede lanzar hechizo | 1400 ms |
| Regeneración de HP (descansando) | 100 ms |
| Regeneración de HP (activo) | 1600 ms |
| Veneno / Parálisis | 500 ms |
| Invocación | 1001 ms |
| Chequeo de anti-macro (WS) | 180 min |
| Guardado de usuarios | 180 s |

---

## Sistema de seguridad y anti-trampa

- **`modCentinela.bas`** — sistema anti-macro: detecta usuarios que trabajan sin responder a desafíos
- **`SecurityIp.bas`** — baneos por IP, control de conexiones múltiples
- **`clsAntiMassClon.cls`** — previene clonación masiva de personajes
- **Verificación MD5** — el servidor valida el hash del ejecutable del cliente en la conexión
- **Condicional `Testeo = 0`** en el cliente indica que el modo de prueba estaba deshabilitado en producción

---

## Sistemas de juego destacados

### Sistema de Pretorianos
- NPCs especiales de facción implementados en `clsClanPretoriano.cls` (~118 KB) y `praetorians.bas`
- Los Pretorianos actúan como NPCs de escolta/guardia para clanes con acceso a `Pretorianos.dat`

### Sistema de facciones
- Dos facciones: **Real** (ciudadano) y **Caos** (criminal)
- `ModFacciones.bas` (~41 KB) gestiona las transiciones entre estados
- Mapa de facción Real (Mapa 3) y Caos (Mapa 4) con zonas de combate diferenciadas

### Sistema de retos
- `Mod_Retos1vs1.bas` y `Mod_Retos2vs2.bas` — combates en arenas controladas
- `frmRetos` en el cliente — interfaz de desafíos

### Sistema de clanes/guilds
- `modGuilds.bas` (~85 KB) — gestión completa de guilds
- `clsClan.cls` (~33 KB) — clase de clan con miembros, rangos, guerras, alianzas
- Persistencia en el directorio `server/Guilds/`
- El cliente tiene 7 formularios dedicados a guilds

### Sistema de trabajos
- `Trabajo.bas` (~128 KB) — pesca, tala, minería, carpintería, herrería
- Rates diferenciados entre clase Trabajador y otras clases
- Anti-macro integrado con `modCentinela.bas`

---

## Diagrama de datos del personaje

Los personajes tienen (inferido de `Declares.bas` y `FileIO.bas`):

- **Stats:** HP, MaxHP, Stamina, Hambre, Sed, Maná, Oro
- **Atributos:** Fuerza, Agilidad, Inteligencia, Carisma, Constitución (inferidos del sistema de creación)
- **Equipamiento:** Arma, Armadura, Casco, Escudo
- **Inventario:** múltiples slots
- **Habilidades:** sistema de skills con puntos asignables (`frmSkills3`)
- **Posición:** mapa, X, Y
- **Facción/Estado:** Real, Caos, Muerto, Invisible, Paralizado, Envenenado, Meditando
- **Raza + Clase + Género**
- **Guild** (si pertenece a uno)
- **Nivel + Experiencia**
