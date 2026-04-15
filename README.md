# PreCita

Aplicacion de escritorio en Python para gestionar citas y contactos, sincronizar eventos de Google Calendar y enviar recordatorios por correo mediante Gmail.

Version actual: (vea VERSION).

## Caracteristicas actuales

- Sincronizacion de calendario con Google Calendar en ventana configurable de `7` a `49` dias (`15` por defecto).
- Seleccion explicita de calendario de Google para evitar leer eventos de calendarios no deseados.
- Vistas de agenda `diaria`, `semanal` y `mensual`, con navegacion por periodos y acceso rapido a "Hoy".
- Gestion de contactos: alta, edicion, eliminacion y vinculacion con citas.
- Envio de recordatorios por Gmail en modo manual (`Lanzar pendientes`) y automatico por temporizador.
- Plantilla de correo editable con asunto y cuerpo en HTML, formato basico (negrita/cursiva/subrayado/listas) y variables dinamicas.
- Adjuntos reutilizables en plantilla con controles de seguridad y limite de tamano total (16 MB).
- Ventana `Datos locales` para inspeccionar `~/.precita/` (tamano total, archivos clave y estado de adjuntos).
- Boton `Optimizar` en `Datos locales` que ejecuta `VACUUM` sobre `precita.db` para compactar la base de datos.
- Boton `Encriptar base de datos` en `Datos locales` para habilitar/deshabilitar el cifrado local de `precita.db` mediante contrasena.
- Indicador visual de estado de sesion de Google y opcion de cierre de sesion/desvinculacion local.
- Configuracion de tema (claro/oscuro), notificaciones de Windows e inicio automatico en Windows.
- Atajos globales: `Ctrl+H` para abrir `Ayuda`, `Ctrl+,` para `Configuracion`.

## Tecnologias

- Python `3.10+` (`3.12` recomendado)
- PyQt6
- SQLite
- Google Calendar API
- Gmail API

## Requisitos e instalacion

1. Clona o descarga este repositorio.
2. Crea y activa un entorno virtual.
3. Instala dependencias:

```bash
pip install PyQt6 PyQt6-WebEngine google-auth google-auth-oauthlib google-api-python-client pytz
```

## Ejecucion

Desde la raiz del proyecto:

```bash
python main.py
```

## Configuracion y almacenamiento local

PreCita guarda su informacion en la carpeta local del usuario `~/.precita/`:

- `precita.db`: base de datos SQLite.
- `db_encryption_config.json`: estado local de encriptacion de la base de datos.
- `token.json`: token OAuth local para Google.
- `client_secret.json`: opcional, para flujo OAuth alternativo.
- `template_attachments/`: copia local segura de adjuntos de plantilla.

Desde el menu principal puede abrir `Datos locales` para consultar esta carpeta y, cuando lo necesite, ejecutar `Optimizar` (comando SQLite `VACUUM`) para compactar `precita.db`.

Adicionalmente, en `Datos locales` puede abrir `Encriptar base de datos` para:

- Elegir entre `Deshabilitar encriptacion (predeterminado)` y `Habilitar encriptacion (recomendable)`.
- Definir una contrasena alfanumerica obligatoria para proteger `precita.db` cuando la encriptacion esta habilitada.
- Introducir la contrasena al arrancar PreCita cuando la base de datos esta cifrada.
- Mantener una clave compatible para descifrado en otras herramientas autorizadas, ya que la clave elegida se usa directamente en el cifrado local del fichero.

Variables de entorno opcionales:

- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`

Si no se definen, la app utiliza credenciales embebidas y solicitara autenticacion OAuth al usar funciones de Google.

## Flujo de uso recomendado

1. Abra PreCita.
2. Use `Sincronizar` para iniciar sesion con Google (si aplica) y traer citas del calendario seleccionado.
3. Revise o cree contactos en `Gestion de contactos` o `Anadir contacto`.
4. Ajuste asunto/cuerpo/adjuntos desde `Personalizar plantilla`.
5. Revise `Datos locales` para verificar el estado de `~/.precita/`, usar `Optimizar` y configurar `Encriptar base de datos` si desea proteger `precita.db` con clave.
6. Ejecute `Lanzar pendientes` para envio manual o deje activo el envio automatico.
7. Revise el log integrado para confirmar sincronizaciones, envios y posibles errores.

## Variables disponibles en la plantilla

- `{nombre_citado}`
- `{apellidos_citado}`
- `{correo_citado}`
- `{tlf_citado}`
- `{hora_cita}`
- `{fecha_cita}`
- `{dia_semana}`

## Reglas de seguridad para adjuntos

- Se bloquean extensiones ejecutables o de riesgo segun politica de Gmail.
- Se permite `*.zip` solo si su contenido no incluye archivos bloqueados.
- Se bloquean otros formatos comprimidos (`.rar`, `.7z`, `.tar`, `.gz`, `.bz2`, `.xz`).
- El tamano combinado de cuerpo + adjuntos esta limitado a `16 MB`.

## Configuracion funcional

Desde `Configuracion` (tambien con `Ctrl+,`):

- Tema visual (`Claro` / `Oscuro`).
- Intervalo de revision automatica de recordatorios.
- Dias de sincronizacion de calendario (rango 7-49).
- Calendario de Google seleccionado.
- Activacion/desactivacion de notificaciones de Windows.
- Inicio con Windows (solo en `win32`).
- Desvincular cuenta de Google en este equipo.
- Restablecer valores predeterminados.

Desde `Ayuda` (tambien con `Ctrl+H`) puede abrir enlaces directos a funciones internas como sincronizacion, contactos, plantilla, configuracion y almacenamiento.

## Estructura del proyecto

- `main.py`: interfaz, logica de negocio, persistencia, sincronizacion y envio.
- `VERSION`: version de aplicacion.
- `precita.ico`: icono principal.
- `LICENSE`: licencia GNU GPL v3.

## Politica de privacidad

PreCita es una herramienta de ejecucion local y prioriza el control del usuario sobre sus datos:

1. **Datos locales**: contactos, citas y configuracion se guardan en el equipo del usuario.
2. **Uso de APIs de Google**: se usa Google Calendar para lectura de eventos y Gmail para envio de recordatorios.
3. **Transferencia de datos**: no hay telemetria ni envio a terceros fuera de las APIs necesarias de Google.
4. **Control y eliminacion**: para borrar datos, cierre la app y elimine `~/.precita/`; opcionalmente revoque el acceso en su cuenta de Google.
5. **Licencia**: software distribuido bajo GNU GPL v3, sin garantias (vea LICENSE). 

Mas informacion: <https://eucarigo.com>
