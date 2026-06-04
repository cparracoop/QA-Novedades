QA Novedades

Aplicación web para la gestión de proyectos, usuarios y seguimiento de novedades, con control de acceso basado en roles y permisos para garantizar una administración segura de la información.

🚀 Características
Creación y administración de proyectos.
Gestión de usuarios y roles.
Asignación de miembros a proyectos.
Seguimiento y registro de novedades.
Control de permisos según perfil de usuario.
Interfaz intuitiva para la colaboración en equipo.
🔒 Seguridad y Permisos

La aplicación implementa un esquema de control de acceso para proteger la información y las acciones críticas:

Solo el creador del proyecto puede modificar el nombre y la descripción del proyecto.
La administración de usuarios está restringida según el nivel de permisos.
Los usuarios con perfiles inferiores no pueden eliminar ni modificar administradores del proyecto.
Validación de acciones tanto en la interfaz como en la lógica de negocio.
🛠️ Tecnologías

Actualice esta sección según las tecnologías utilizadas en el proyecto.

Frontend: React / Next.js
Backend: Node.js
Base de Datos: PostgreSQL / MySQL
Autenticación: JWT
ORM: Prisma
📦 Instalación
# Clonar el repositorio
git clone https://github.com/usuario/qa-novedades.git

# Ingresar al proyecto
cd qa-novedades

# Instalar dependencias
npm install

# Ejecutar en desarrollo
npm run dev
⚙️ Configuración

Crear un archivo .env con las variables necesarias para la conexión a la base de datos y servicios externos.

Ejemplo:

DATABASE_URL=
JWT_SECRET=
📁 Estructura General
src/
├── components/
├── pages/
├── services/
├── hooks/
├── utils/
├── models/
└── styles/
👥 Roles
Rol	Permisos
Administrador	Gestión completa del sistema y proyectos
Creador del Proyecto	Administración total de su proyecto
Miembro	Acceso según permisos asignados
📈 Mejoras Implementadas
Restricción de edición del nombre y descripción del proyecto únicamente al creador.
Control jerárquico en la administración de usuarios.
Protección contra eliminación o modificación de administradores por usuarios de menor privilegio.
🤝 Contribuciones

Las contribuciones son bienvenidas. Cree una rama para sus cambios y envíe un Pull Request para revisión.

📄 Licencia

Este proyecto está bajo la licencia MIT. Consulte el archivo LICENSE para más información.
