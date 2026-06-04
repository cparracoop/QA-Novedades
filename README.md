# QA Novedades App

Aplicación local en Python para registrar novedades de pruebas de software por proyecto, adjuntar evidencias y generar reportes compartibles con proveedores.

## Funcionalidades principales

- Creación y gestión de proyectos
- Registro de pruebas y novedades dentro de cada proyecto
- Carga de evidencias (imágenes, PDF, logs y video corto)
- Seguimiento del ciclo de atención con estados como:
  - Reportada
  - En análisis
  - En ajuste proveedor
  - Ajuste aplicado
  - Subsanada
  - Resuelta
  - Se mantiene
  - Descartada
- Campo de **resumen para proveedor**
- Campo de **validación QA / re-test**
- Exportación por proyecto en:
  - CSV
  - Excel (.xlsx)
  - PDF ejecutivo

## Requisitos

- Python 3.11 o superior

## Instalación

```bash
python -m venv .venv
```

### Windows

```bash
.venv\Scripts\activate
```

### macOS / Linux

```bash
source .venv/bin/activate
```

Instala dependencias:

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
python app.py
```

Luego abre en el navegador:

```bash
http://127.0.0.1:5000
```

## Base de datos

La aplicación usa SQLite y crea automáticamente la base de datos en:

```text
instance/qa_novedades.db
```

Si vienes de una versión anterior, al iniciar la app se agregan automáticamente las nuevas columnas necesarias para los campos de seguimiento y cierre.

## Cómo usar

1. Crea un proyecto.
2. Entra al proyecto y registra una novedad.
3. Adjunta evidencias.
4. Usa el campo **Resumen para proveedor** para redactar una versión más breve y accionable.
5. Cuando el proveedor entregue ajuste, actualiza el estado y registra el resultado en **Validación QA / re-test**.
6. Exporta el reporte en Excel o PDF según el destinatario.

## Estructura

- `app.py`: aplicación principal
- `templates/`: vistas HTML
- `static/uploads/`: evidencias cargadas
- `instance/`: base de datos SQLite

## Observaciones

- Esta versión está pensada para uso local.
- Si más adelante necesitas multiusuario o acceso compartido, la base puede migrarse a PostgreSQL.
- También se puede agregar autenticación, dashboard y filtros avanzados.


## Compatibilidad

Esta version elimina la dependencia de Pillow para funcionar mejor en entornos con Python 3.14.
Las capturas pegadas desde el portapapeles se convierten a JPG desde el navegador para que el PDF pueda incluirlas sin librerias adicionales.
Si adjuntas imagenes manualmente, el PDF incrusta de forma confiable archivos JPG/JPEG; los demas formatos quedan referenciados por nombre en el informe.

## Ajuste v16 - PDF corporativo con membrete

- El PDF usa el membrete corporativo basado en `FR-CD-23 Hoja Membrete.doc` como fondo visual.
- Se incorporan logos, marco exterior, franja lateral Supersolidaria y pie de pagina institucional.
- Se eliminaron colores no corporativos y se usan verde, amarillo, negro/gris y borde institucional.
- Las evidencias se incluyen en JPG/JPEG/PNG, centradas y dentro de marco.
- El reporte PDF conserva el filtro de incidencias pendientes.

## Seguridad, usuarios y permisos

Esta versión incluye autenticación y control de acceso por roles.

La aplicación crea un administrador inicial automáticamente si la base de datos está vacía. Por seguridad, esos datos no se muestran en la pantalla de login. Puedes definirlos antes de inicializar la base de datos con:

```powershell
$env:ADMIN_EMAIL="admin@tuempresa.com"
$env:ADMIN_PASSWORD="UnaClaveTemporalSegura"
python app.py
```

El usuario de ingreso será siempre el correo electrónico registrado.

### Roles base

- **Administrador**: acceso total, usuarios, proyectos, incidencias y auditoría.
- **Gestor de proyectos**: crea y administra proyectos, asocia usuarios y gestiona incidencias.
- **Tester / Analista QA**: puede asociarse a proyectos y registrar incidencias cuando esté asociado.
- **Consulta**: solo puede ver proyectos asignados; no registra incidencias.

### Reglas principales

- Para ingresar a un proyecto, el usuario debe ser administrador/gestor o estar asociado al proyecto.
- Para registrar incidencias, el usuario debe tener permiso global de incidencias y permiso de registro dentro del proyecto.
- La administración de usuarios se realiza desde **Usuarios**.
- Las acciones relevantes quedan registradas en **Auditoría**.

## Cambios de seguridad e interacción agregados

- Se retiró del login el mensaje con el usuario y contraseña inicial.
- Los usuarios nuevos ya no se crean digitando contraseña manualmente.
- El correo electrónico registrado es también el usuario de ingreso.
- Al crear un usuario, la aplicación genera una contraseña temporal aleatoria y la envía al correo registrado.
- Desde administración se puede reenviar acceso; esto genera una nueva clave temporal y exige cambio de contraseña al ingresar.
- El primer ingreso del usuario obliga a cambiar la contraseña antes de continuar.
- La contraseña nunca se guarda en texto plano; se almacena como hash seguro.
- Se agregaron iconos visuales a navegación, formularios, métricas y acciones principales.
- En el registro inicial de novedades se retiró el campo de validación QA; este queda disponible al editar la novedad o cambiar su estado.
- Los campos principales del formulario de novedades son obligatorios; las evidencias/imágenes son opcionales.

## Configuración de correo SMTP

Para que el envío de claves temporales funcione, define estas variables de entorno antes de iniciar la aplicación:

```powershell
$env:SMTP_HOST="smtp.tuempresa.com"
$env:SMTP_PORT="587"
$env:SMTP_USER="usuario_smtp@tuempresa.com"
$env:SMTP_PASSWORD="clave_smtp"
$env:SMTP_FROM="usuario_smtp@tuempresa.com"
$env:SMTP_TLS="true"
$env:SMTP_SSL="false"
python app.py
```

Para servidores que usen SSL directo, normalmente puerto 465:

```powershell
$env:SMTP_PORT="465"
$env:SMTP_TLS="false"
$env:SMTP_SSL="true"
```

Si SMTP no está configurado o el servidor rechaza el envío, el usuario se crea, pero la aplicación mostrará el detalle técnico del error SMTP. Corrige la configuración y usa **Reenviar acceso** desde Usuarios para generar y enviar una nueva clave temporal.

En entornos corporativos puede ser necesario solicitar a TI: servidor SMTP autorizado, puerto, si usa TLS o SSL, usuario remitente permitido y autorización para enviar desde la red interna.


## Configuración SMTP

Crea un archivo `.env` en la misma carpeta de `app.py` con estas variables:

```env
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_USER=tu_correo@coopetrol.coop
SMTP_PASSWORD=tu_clave_o_clave_de_aplicacion
SMTP_FROM=tu_correo@coopetrol.coop
SMTP_USE_TLS=true
SMTP_USE_SSL=false
```

El usuario de ingreso es siempre el correo electrónico registrado.

## Administrador inicial corregido

Al iniciar la aplicación, la base de datos se corrige automáticamente para dejar como administrador activo:

- Usuario/correo: `cparra@coopetrol.coop`
- Clave inicial: `cparra123`

La cuenta histórica `admin@local` se desactiva y se renombra para evitar conflictos.

## Archivo de correo

El ZIP incluye `.env.example`. Copia ese archivo como `.env`, completa la clave real del correo y reinicia la aplicación.

## Ajuste de permisos aplicado

- El permiso **Crear/administrar proyectos** solo permite crear, editar o administrar proyectos.
- Para registrar o editar novedades el usuario debe tener **Registrar/editar novedades** y estar asociado al proyecto con autorización para registrar.
- Las acciones no permitidas se ocultan en la interfaz y, si se intenta acceder por URL directa, se muestra un mensaje amigable.
- Las novedades cerradas como **Subsanada** o **Descartada** no permiten edición, cambio de estado ni modificación de evidencias.
