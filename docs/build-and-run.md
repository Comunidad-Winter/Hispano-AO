# Compilación y Ejecución — Hispano-AO

> *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Requisitos del entorno

### Obligatorios

| Requisito | Detalle |
|---|---|
| **Sistema operativo** | Windows (XP / 7 / 10 / 11 con modo de compatibilidad) |
| **IDE de compilación** | Microsoft Visual Basic 6.0 SP6 |
| **DirectX** | DirectX 7 o superior (incluido en Windows modernos) |
| **Runtime VB6** | `MSVBVM60.DLL` (incluida en `client/`) |

### OCX/DLL requeridas (se incluyen en `client/`)

Deben estar **registradas** en el sistema antes de compilar o ejecutar:

```powershell
# Registrar desde el directorio client/ (ejecutar como administrador)
regsvr32 COMCTL32.OCX
regsvr32 COMDLG32.OCX
regsvr32 CSWSK32.OCX
regsvr32 MSCOMCTL.OCX
regsvr32 MSINET.OCX
regsvr32 RICHTX32.OCX
regsvr32 VBALPROGBAR6.OCX
```

> **Nota:** En Windows 10/11 puede ser necesario copiar las OCX a `C:\Windows\SysWOW64\` (sistema 64-bit) antes de registrarlas.

---

## Compilar el cliente

### Pasos

1. Abrir `client/Client.vbp` en Visual Basic 6 IDE
2. Verificar que todas las referencias estén resueltas:
   - `Project > References`: OLE Automation, DirectX 7, Microsoft DAO 3.51, ActiveMovie
   - `Project > Components`: CSWSK32, MSCOMCTL, COMCTL32, MSINET, RICHTX32
3. Revisar la condición de compilación condicional: `Testeo = 0` (producción)
4. Compilar: `File > Make Hispano Online.exe`
5. El ejecutable se genera en `client/../` (directorio padre, según `Path32="..\..\"`)

### Parámetros del proyecto (Client.vbp)

```ini
Startup    = "Sub Main"        ; Entry point en Application.bas
ExeName32  = "Hispano Online.exe"
Title      = "HispanoAO"
MajorVer   = 0
MinorVer   = 13
RevisionVer = 8
CondComp   = "Testeo = 0"      ; Modo producción
```

---

## Compilar el servidor

### Pasos

1. Abrir `server/SERVER.VBP` en Visual Basic 6 IDE
2. Verificar referencias: `COMCTL32.OCX`
3. Revisar condiciones de compilación: `UsarQueSocket = 1 : ConUpTime = 1`
   - `UsarQueSocket = 1` → usa la implementación de socket `wsksock.bas`
   - `ConUpTime = 1` → activa el seguimiento de uptime del servidor
4. Compilar: `File > Make server9arreglado.exe`
5. El ejecutable se genera en `server/../` (directorio padre, según `Path32=".."`)

### Parámetros del proyecto (SERVER.VBP)

```ini
Startup    = "Sub Main"              ; Entry point en frmMain.frm
ExeName32  = "server9arreglado.exe"
Title      = "Argentum Online Server"
MajorVer   = 0
MinorVer   = 13
RevisionVer = 0
CondComp   = "UsarQueSocket = 1 : ConUpTime = 1"
```

---

## Configurar el servidor

### 1. `server/Server.ini` — Configuración principal

```ini
[INIT]
StartPort=7666              ; Puerto TCP de escucha
Hide=1                      ; Ocultar ventana al iniciar (1=sí)
AllowMultiLogins=1          ; Permitir múltiples cuentas por IP
IdleLimit=5                 ; Minutos de inactividad antes de kick
Version=0.13.7              ; Versión de cliente aceptada
IniciarDesdeBackUp=1        ; Iniciar desde backup si existe
CleanInterval=15            ; Intervalo de limpieza del mundo (minutos)
PuedeCrearPersonajes=1      ; Habilitar creación de personajes
ServerSoloGMs=0             ; 0=abierto a todos, 1=solo GMs
Testing=0                   ; 0=producción, 1=modo prueba
MaxUsers=550                ; Máximo de conexiones simultáneas
StartPos=1-58-45            ; Posición de inicio: Mapa-X-Y

[Admines]
Admin1=NombreAdmin1
Admin2=NombreAdmin2

[Dioses]
Dios1=NombreDios1

[MD5Hush]
Activado=0                  ; 0=desactivado, 1=verificar MD5 del cliente
```

### 2. `server/Configuracion.ini` — Parámetros del juego

```ini
[Servidor]
NivelMaximo=47              ; Nivel máximo alcanzable

[Rates]
Experiencia=8               ; Multiplicador de experiencia (8x)
Oro=5                       ; Multiplicador de oro (5x)

[Mapas]
MapaGm=90                   ; Mapa de Game Masters
Mapa1vs1=73                 ; Arena 1vs1
Mapa2vs2=74                 ; Arena 2vs2
Prision=89-75-47            ; Mapa-X-Y de prisión
Libertad=1-43-60            ; Mapa-X-Y de punto de libertad
MapReal=3                   ; Zona facción Real
MapCaos=4                   ; Zona facción Caos
MapaPretoriano=88-34-25-67-25
```

### 3. Estructura de directorios requerida

Antes de ejecutar el servidor, verificar que existan:

```
server/
├── Charfile/       ← DEBE EXISTIR (vacío está bien, almacena personajes)
├── Logs/           ← DEBE EXISTIR (vacío está bien, almacena logs)
├── Guilds/         ← DEBE EXISTIR (vacío está bien, almacena clanes)
├── WorldBackup/    ← Recomendado (backups automáticos del mundo)
├── Maps/           ← YA EXISTE en el repo (93 mapas)
└── Dat/            ← YA EXISTE en el repo (NPCs, objetos, hechizos)
```

```powershell
# Crear directorios faltantes desde server/
New-Item -ItemType Directory -Force -Path "Charfile", "Logs", "Guilds", "WorldBackup"
```

---

## Ejecutar el servidor

1. Navegar al directorio que contiene `server9arreglado.exe` (directorio padre de `server/`)
2. Ejecutar `server9arreglado.exe`
3. La ventana `frmMain` mostrará el panel de control
4. El servidor escucha en `0.0.0.0:7666` por defecto
5. Para minimizar a la bandeja del sistema: botón "Systray" en `frmMain`

---

## Ejecutar el cliente

1. Navegar al directorio que contiene `Hispano Online.exe` (directorio padre de `client/`)
2. Ejecutar `Hispano Online.exe`
3. En la pantalla `frmConnect`:
   - **Servidor:** `127.0.0.1` (local) o IP del servidor remoto
   - **Puerto:** `7666`
4. Crear personaje o conectarse con uno existente

---

## Actualización automática del cliente

El cliente incluye `Autoupdate.exe` (~36 KB). Este binario probablemente gestiona la descarga y aplicación de parches desde un servidor de actualización. Usa `MSINET.OCX` (HTTP) y `Unzip32.dll` (descompresión de parches `.zip`).

- La URL del servidor de actualizaciones no está en el repositorio (inferido: configurada en `INIT/Update.ini` o `INIT/versiones.ini`)
- `INIT/Ver.bin` contiene la versión actual instalada
- `INIT/versiones.ini` lista versiones y URLs de descarga

---

## Compilar desde línea de comandos (alternativa)

VB6 soporta compilación silenciosa con `vb6.exe`:

```bat
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make "client\Client.vbp"
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make "server\SERVER.VBP"
```

---

## World Editor

El World Editor está disponible como binario precompilado en `WE HISPANO AO/WorldEditor.exe`. No requiere compilación.

### Ejecutar el World Editor

1. Navegar a `WE HISPANO AO/`
2. Ejecutar `WorldEditor.exe`
3. El editor carga automáticamente los mapas desde `Maps/` y los datos de `Dat/`
4. Exportar mapas al formato `Mapa#.map + .dat + .inf`
5. Copiar los archivos exportados a `server/Maps/` y `client/Mapas/`

---

## Problemas conocidos al ejecutar en Windows moderno

| Problema | Causa probable | Solución |
|---|---|---|
| "Component not registered" al abrir VBP | OCX sin registrar | Registrar todas las OCX con `regsvr32` |
| Error de runtime al ejecutar el cliente | DLL en ubicación incorrecta | Colocar DLL y OCX en el mismo directorio que el EXE |
| El servidor no acepta conexiones | Firewall bloqueando puerto 7666 | Agregar regla de entrada al Firewall de Windows para TCP 7666 |
| Pantalla en negro en el cliente | Incompatibilidad DPI en Windows 10/11 | Deshabilitar escalado DPI en propiedades del EXE → Compatibilidad |
| Crash al cargar mapas | Directorios de datos ausentes | Verificar existencia de `Charfile/`, `Logs/`, `Guilds/` |
