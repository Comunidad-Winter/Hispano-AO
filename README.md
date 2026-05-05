# Hispano AO
 Liberación del proyecto Hispano AO (v. 2017)

Antes de que digan "A PERO"...
1. Tengo permiso de AmishaR y fui junto a él, director del proyecto. Si preguntan por Hekamiah (el dueño previo), él mismo le dejo el proyecto a AmishaR en 2008, el resto pueden consultarlo con él.
2. Soy MAB (por si todavía alguno no capto el porque de la liberación).
3. Hispano AO es un mod de la versión 13.x de Argentum Online (https://www.comunidadargentum.com/).
4. No tenemos derechos de nada y no nos interesan tampoco.
5. El código tiene cosas que removí que no eran importantes, el resto esta tal cual y funciona tal cual.
6. Hispano AO nunca destacó por su código, sí por su Staff y la atención al usuario. Es algo que hasta el día de hoy muchos no logran ni van a lograr entender.
7. Agradecimientos a todos los que participaron del proyecto, tanto a todo miembro que paso por el Staff como a los usuarios que disfrutaron con nosotros.

---

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Qué es

**Hispano-AO** es un servidor privado de **Argentum Online** basado en el código fuente histórico de la versión **0.13.x** del juego. Esta rama (`2017`) corresponde al snapshot del código liberado originalmente por MAB en 2017.

Argentum Online (AO) es un MMORPG 2D de origen argentino creado originalmente por Pablo Ignacio Márquez, licenciado bajo la **Affero General Public License**.

Este repositorio contiene el código fuente de tres componentes operacionales:

- **Cliente** (`Hispano Online.exe`) — interfaz gráfica del jugador.
- **Servidor** (`server9arreglado.exe`) — lógica de juego y networking.
- **World Editor** — herramienta de edición de mapas.

---

## Estructura del repositorio

```
Hispano-AO/ (rama 2017)
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
│   └── *.OCX / *.DLL       # Dependencias en tiempo de ejecución
│
├── server/ (o Servidor/)   # Proyecto VB6 del servidor de juego
│   ├── Codigo/             # Código fuente completo del servidor
│   ├── Maps/               # Mapas del servidor (.map + .dat + .inf)
│   ├── Dat/                # Datos del mundo (NPCs, objetos, hechizos)
│   ├── Charfile/           # Archivos de personajes guardados (runtime)
│   ├── Server.ini          # Configuración principal del servidor
│   └── SERVER.VBP          # Archivo de proyecto Visual Basic 6
│
└── WE HISPANO AO/          # World Editor (herramienta de edición de mapas)
    ├── WorldEditor.exe     # Ejecutable del editor de mapas
    └── Maps/ Dat/ Graficos/ # Datos de referencia del editor
```

> **Nota sobre la estructura:** La rama `2017` puede presentar diferencias de organización de carpetas respecto a la rama `2018` (ej.: carpeta `Servidor/` en lugar de `server/`). Ver la rama `2018` para la versión más reciente y reorganizada.

---

## Componentes principales

### Cliente (`client/`)

- **Lenguaje:** Visual Basic 6
- **Entry point:** `Sub Main` en `Application.bas`
- **Formulario principal:** `frmMain`
- **Ejecutable:** `Hispano Online.exe`
- **Versión:** 0.13.x

### Servidor (`server/` o `Servidor/`)

- **Lenguaje:** Visual Basic 6
- **Entry point:** `Sub Main` en `frmMain.frm`
- **Ejecutable:** `server9arreglado.exe`
- **Puerto TCP:** 7666
- **Máximo usuarios:** configurable en `Server.ini`

### World Editor

- Ejecutable binario: `WorldEditor.exe`
- Sin código fuente disponible en el repositorio

---

## Cómo funciona

Arquitectura **cliente-servidor** TCP binario sobre puerto 7666:

```
CLIENTE (VB6)  ←──TCP 7666──→  SERVIDOR (VB6)
TileEngine (render)              modNuevoTimer (game loop 50ms)
Protocol.bas (envío)             Protocol.bas (recepción y despacho)
ProtocolCmdParse.bas (recep.)    SistemaCombate / AI_NPC / modHechizos
clsByteQueue (serialización)     FileIO.bas (persistencia en disco)
```

Para más detalle técnico, ver los archivos en `docs/`.

---

## Instalación / Compilación / Ejecución

### Requisitos

- Windows (XP / 7 / 10 con compatibilidad)
- Visual Basic 6.0 SP6
- Registrar las OCX incluidas en `client/` con `regsvr32`
- DirectX 7 o superior

### Compilar

```
Abrir client/Client.vbp en VB6 → File > Make Hispano Online.exe
Abrir server/SERVER.VBP en VB6 → File > Make server9arreglado.exe
```

### Configurar el servidor

1. Editar `server/Server.ini` (puerto, usuarios máximos, GMs)
2. Editar `server/Configuracion.ini` (rates, niveles, intervalos)
3. Crear directorios si no existen: `Charfile/`, `Logs/`, `Guilds/`

Ver [`docs/build-and-run.md`](docs/build-and-run.md) para instrucciones completas.

---

## Estado del proyecto

| Aspecto | Estado |
|---|---|
| Tipo | Preservación — snapshot histórico 2017 |
| Código fuente | Completo (cliente + servidor) con algunas partes removidas según el autor |
| Compilabilidad | Requiere VB6 — entorno obsoleto |
| Desarrollo activo | No — liberación histórica |
| Licencia base | Affero GPL (heredada del proyecto AO original) |

---

## Documentación técnica

| Documento | Contenido |
|---|---|
| [`docs/architecture.md`](docs/architecture.md) | Arquitectura técnica y diagramas |
| [`docs/components.md`](docs/components.md) | Detalle de todos los módulos |
| [`docs/build-and-run.md`](docs/build-and-run.md) | Compilación y ejecución |
| [`docs/notes.md`](docs/notes.md) | Notas, asunciones y advertencias |

---

## Créditos y licencia

- **Concepto y código base de AO:** Pablo Ignacio Márquez
- **Proyecto Hispano-AO:** AmishaR (director), Hekamiah (fundador), MAB (liberación 2017), y el equipo de Staff
- **Licencia:** [Affero General Public License v1+](http://www.affero.org/oagpl.html)

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*
