# Arquitectura T├®cnica ÔÇö Hispano-AO

> *Esta documentaci├│n fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Visi├│n de alto nivel

Hispano-AO es un MMORPG 2D cl├ísico con arquitectura **cliente-servidor** donde ambos extremos est├ín implementados en **Visual Basic 6**. La comunicaci├│n ocurre exclusivamente por **TCP binario** sobre el puerto 7666.

```
ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ
Ôöé  CLIENTE (Hispano Online.exe)                                    Ôöé
Ôöé                                                                  Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  Ôöé TileEngine Ôöé  Ôöé  clsAudio    Ôöé  Ôöé Formularios VB6      Ôöé     Ôöé
Ôöé  Ôöé (render 2D)Ôöé  Ôöé (WAV/MIDI/MP3Ôöé  Ôöé (frmMain, frmSkills, Ôöé     Ôöé
Ôöé  Ôöé DIBSection Ôöé  Ôöé via DirectX) Ôöé  Ôöé  frmGuild*, frmComercÔöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé         Ôåò                                    Ôåò                   Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  Ôöé              clsByteQueue (cola binaria)                 Ôöé     Ôöé
Ôöé  Ôöé  Protocol.bas (env├¡o) Ôåö ProtocolCmdParse.bas (recepci├│n)Ôöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé                        Ôåò TCP/IP (CSWSK32.OCX)                    Ôöé
ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔö¼ÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ
                                   Ôöé Puerto 7666
ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔû╝ÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ
Ôöé  SERVIDOR (server9arreglado.exe)                                  Ôöé
Ôöé                                                                  Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  Ôöé  Capa de red: wsksock.bas + wskapiAO.bas (hasta 550)    Ôöé     Ôöé
Ôöé  Ôöé  clsByteQueue por usuario (entrada) + modSendData (salidaÔöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé                        Ôåò                                         Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  Ôöé  Protocol.bas ÔÇö decodifica paquetes entrantes,          Ôöé     Ôöé
Ôöé  Ôöé  despacha a m├│dulos de l├│gica                           Ôöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔö¼ÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé                             Ôåô                                    Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  ÔöéGameLogic Ôöé ÔöéSistemaComÔöé ÔöémodHechizosÔöé ÔöéTrabajo.bas     Ôöé     Ôöé
Ôöé  ÔöéGeneral   Ôöé Ôöébate.bas  Ôöé Ôöé           Ôöé Ôöé(pesca/tala/minaÔöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  ÔöéAI_NPC    Ôöé ÔöémodGuilds Ôöé ÔöéModFaccioneÔöé ÔöéMod_Retos1vs1/2 Ôöé     Ôöé
Ôöé  Ôöé          Ôöé Ôöé          Ôöé Ôöés          Ôöé Ôöé                Ôöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé                             Ôåò                                    Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  Ôöé  FileIO.bas ÔÇö persistencia (lectura/escritura de disco)  Ôöé     Ôöé
Ôöé  Ôöé  /Charfile/ (personajes) /Maps/ (mapas) /Dat/ (datos)   Ôöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
Ôöé                                                                  Ôöé
Ôöé  ÔöîÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÉ     Ôöé
Ôöé  Ôöé  modNuevoTimer (game loop) ÔÇö ciclo base 50 ms           Ôöé     Ôöé
Ôöé  Ôöé  Gestiona intervalos de: ataque, hechizos, NPC AI,      Ôöé     Ôöé
Ôöé  Ôöé  regeneraci├│n, hambre/sed, veneno, invisibilidad, etc.  Ôöé     Ôöé
Ôöé  ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ     Ôöé
ÔööÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÇÔöÿ
```

---

## Capa de red

### Cliente
- Usa `CSWSK32.OCX` (Crescent WinSock Control 32-bit) para la conexi├│n TCP
- Los datos enviados y recibidos pasan por `clsByteQueue`, que encola bytes en un buffer
- `Protocol.bas` serializa los comandos a enviar al servidor
- `ProtocolCmdParse.bas` parsea los comandos recibidos del servidor y actualiza el estado del juego

### Servidor
- Gestiona hasta **550 conexiones simult├íneas** (configurable en `Server.ini ÔåÆ MaxUsers`)
- Cada usuario tiene su propio ├¡ndice (`UserIndex`) en el array global `UserList()`
- `wsksock.bas` provee la capa de socket de bajo nivel
- `wskapiAO.bas` implementa la API de alto nivel sobre los sockets
- `modSendData.bas` centraliza el env├¡o de datos a clientes
- La verificaci├│n MD5 del cliente est├í soportada (`[MD5Hush]` en `Server.ini`)

---

## Protocolo de comunicaci├│n

El protocolo es **binario propietario**, usando `clsByteQueue` para empaquetar/desempaquetar:

- **Formato:** Big-endian impl├¡cito, sin delimitadores de frame expl├¡citos visibles en el c├│digo
- Cada paquete comienza con un byte de **opcode** que identifica el tipo de comando
- El parser en `ProtocolCmdParse.bas` (cliente) y `Protocol.bas` (servidor) despacha seg├║n el opcode
- No hay capa de seguridad TLS ÔÇö el tr├ífico es en texto/binario plano
- La capa de cifrado (`clsCripto.cls`) existe en el cliente pero su alcance real en el protocolo requiere an├ílisis adicional

---

## Motor gr├ífico (cliente)

- **Renderizado:** GDI + DIBSections (`cDIBSection.cls`)
- `clsSurfaceManStatic` ÔÇö superficies est├íticas (tiles de fondo, objetos fijos)
- `clsSurfaceManDyn` ÔÇö superficies din├ímicas (personajes, NPCs, proyectiles)
- `clsSurfaceManager` ÔÇö interfaz abstracta sobre los dos managers anteriores
- `TileEngine.bas` (~89 KB) contiene la l├│gica completa de renderizado de tiles
- El juego usa graficos empaquetados referenciados por ├¡ndices (`Graficos.ind`, `graficos.ini`)
- **No usa DirectX para gr├íficos** ÔÇö usa GDI puro (DIBSection sobre HDC)
- DirectX solo se usa para **audio** (`dx7vb.dll`, `dx8vb.dll`)

---

## Sistema de audio (cliente)

- Implementado en `clsAudio.cls` (~41 KB)
- Soporta tres formatos:
  - **WAV** ÔÇö efectos de sonido (directorio `WAV/`)
  - **MIDI** ÔÇö m├║sica de fondo (directorio `MIDI/`)
  - **MP3** ÔÇö m├║sica de alta calidad (directorio `MP3/`)
- Usa DirectX (`quartz.dll` ÔÇö DirectShow) y posiblemente `dx8vb.dll`
- Los sonidos 3D se identifican con la constante `NO_3D_SOUND = 0` (coordenadas especiales)

---

## Sistema de mapas

Tres archivos por mapa, todos binarios o texto estructurado:

| Extensi├│n | Tama├▒o t├¡pico | Contenido |
|---|---|---|
| `.map` | 30ÔÇô50 KB | Capa gr├ífica de tiles (terreno, objetos decorativos) |
| `.dat` | ~260 bytes | Metadatos: nombre del mapa, m├║sica, zona segura, lluvia, etc. |
| `.inf` | 10ÔÇô15 KB | Capa l├│gica: triggers, spawns de NPCs, posiciones de objetos din├ímicos |

**Mapas especiales configurados (verificado en `Configuracion.ini`):**

| Mapa | N├║mero | Uso |
|---|---|---|
| Mapa GM | 90 | Zona de Game Masters |
| Arena 1vs1 | 73 | Retos 1 vs 1 |
| Arena 2vs2 | 74 | Retos 2 vs 2 |
| Prisi├│n | 89 | Destino de jugadores baneados/encarcelados (pos 75-47) |
| Libertad | 1 | Punto de libertad (pos 43-60) |
| Mapa Real | 3 | Zona facci├│n Real |
| Mapa Caos | 4 | Zona facci├│n Caos |
| Mapa Pretoriano | 88 | Zona de clanes Pretorianos |

---

## Sistema de persistencia

Todo el I/O de datos pasa por `FileIO.bas` (m├│dulo `ES`):

### Personajes (`/Charfile/`)
- Un archivo binario por personaje
- Almacena: stats, inventario, posici├│n, estado, hechizos aprendidos, habilidades, etc.
- La persistencia se activa cada `IntervaloGuardarUsuarios = 180` segundos (configurable)

### Datos de juego (`/Dat/`)
- Cargados al inicio del servidor, mantenidos en memoria
- `NPCs.dat` ÔåÆ array de estructuras NPC
- `obj.dat` ÔåÆ array de definiciones de objetos/├¡tems
- `Hechizos.dat` ÔåÆ array de hechizos con efectos y requerimientos

### Mapas (`/Maps/`)
- Los `.map` e `.inf` se cargan en memoria al inicio
- Los `.dat` contienen metadatos por mapa

---

## Game loop y temporizaci├│n

El game loop del servidor est├í implementado en `modNuevoTimer.bas`:

- Ciclo base de **50 ms** (`IntervaloTimerExec`)
- Los intervalos por acci├│n se cargan desde `Configuracion.ini` y `Server.ini`
- El servidor mantiene un contador de "tolerancia" (`Tolerancia_FailIntervalo = 7`) para acciones en intervalos incorrectos

**Intervalos cr├¡ticos (valores por defecto de `Server.ini`):**

| Acci├│n | Intervalo |
|---|---|
| Movimiento de NPC | 200 ms |
| IA de NPC | 380 ms |
| NPC puede atacar | 1600 ms |
| Usuario puede atacar | 1500 ms |
| Usuario puede lanzar hechizo | 1400 ms |
| Regeneraci├│n de HP (descansando) | 100 ms |
| Regeneraci├│n de HP (activo) | 1600 ms |
| Veneno / Par├ílisis | 500 ms |
| Invocaci├│n | 1001 ms |
| Chequeo de anti-macro (WS) | 180 min |
| Guardado de usuarios | 180 s |

---

## Sistema de seguridad y anti-trampa

- **`modCentinela.bas`** ÔÇö sistema anti-macro: detecta usuarios que trabajan sin responder a desaf├¡os
- **`SecurityIp.bas`** ÔÇö baneos por IP, control de conexiones m├║ltiples
- **`clsAntiMassClon.cls`** ÔÇö previene clonaci├│n masiva de personajes
- **Verificaci├│n MD5** ÔÇö el servidor valida el hash del ejecutable del cliente en la conexi├│n
- **Condicional `Testeo = 0`** en el cliente indica que el modo de prueba estaba deshabilitado en producci├│n

---

## Sistemas de juego destacados

### Sistema de Pretorianos
- NPCs especiales de facci├│n implementados en `clsClanPretoriano.cls` (~118 KB) y `praetorians.bas`
- Los Pretorianos act├║an como NPCs de escolta/guardia para clanes con acceso a `Pretorianos.dat`

### Sistema de facciones
- Dos facciones: **Real** (ciudadano) y **Caos** (criminal)
- `ModFacciones.bas` (~41 KB) gestiona las transiciones entre estados
- Mapa de facci├│n Real (Mapa 3) y Caos (Mapa 4) con zonas de combate diferenciadas

### Sistema de retos
- `Mod_Retos1vs1.bas` y `Mod_Retos2vs2.bas` ÔÇö combates en arenas controladas
- `frmRetos` en el cliente ÔÇö interfaz de desaf├¡os

### Sistema de clanes/guilds
- `modGuilds.bas` (~85 KB) ÔÇö gesti├│n completa de guilds
- `clsClan.cls` (~33 KB) ÔÇö clase de clan con miembros, rangos, guerras, alianzas
- Persistencia en el directorio `server/Guilds/`
- El cliente tiene 7 formularios dedicados a guilds

### Sistema de trabajos
- `Trabajo.bas` (~128 KB) ÔÇö pesca, tala, miner├¡a, carpinter├¡a, herrer├¡a
- Rates diferenciados entre clase Trabajador y otras clases
- Anti-macro integrado con `modCentinela.bas`

---

## Diagrama de datos del personaje

Los personajes tienen (inferido de `Declares.bas` y `FileIO.bas`):

- **Stats:** HP, MaxHP, Stamina, Hambre, Sed, Man├í, Oro
- **Atributos:** Fuerza, Agilidad, Inteligencia, Carisma, Constituci├│n (inferidos del sistema de creaci├│n)
- **Equipamiento:** Arma, Armadura, Casco, Escudo
- **Inventario:** m├║ltiples slots
- **Habilidades:** sistema de skills con puntos asignables (`frmSkills3`)
- **Posici├│n:** mapa, X, Y
- **Facci├│n/Estado:** Real, Caos, Muerto, Invisible, Paralizado, Envenenado, Meditando
- **Raza + Clase + G├®nero**
- **Guild** (si pertenece a uno)
- **Nivel + Experiencia**
