# MultitareaApp

Una aplicación de escritorio para la organización de archivos, gestión de bases de datos y personalización de temas.

## Descripción

MultitareaApp es una aplicación de escritorio desarrollada en Python utilizando Tkinter y ttkthemes. Permite a los usuarios organizar archivos en directorios según su tipo, gestionar conexiones a bases de datos MySQL (con soporte para túneles SSH) y personalizar la apariencia de la interfaz mediante la selección de temas.

## Características

-   **Organización de Archivos**: Permite seleccionar un directorio y organizar automáticamente los archivos en subdirectorios basados en categorías predefinidas (Imágenes, Documentos, Audio, Video, etc.).
-   **Gestión de Bases de Datos**: Facilita la conexión a bases de datos MySQL, con soporte para conexiones directas o a través de túneles SSH. Permite explorar las tablas de la base de datos y visualizar sus registros.
-   **Personalización de Temas**: Ofrece una variedad de temas visuales para personalizar la apariencia de la aplicación.
-   **Historial de Operaciones**: Permite deshacer la última operación de organización de archivos realizada.
-   **Generación de Reportes en Excel**: Crea un archivo Excel con un reporte detallado de los archivos organizados, incluyendo nombre, ruta, tipo, tamaño, fecha de creación y modificación.
-   **Logging**: Registra todas las operaciones y errores en un archivo de registro para facilitar el seguimiento y la resolución de problemas.

## Requisitos

-   Python 3.x
-   Librerías:
    -   tkinter
    -   ttkthemes
    -   openpyxl
    -   logging
    -   threading
    -   mysql.connector
    -   paramiko

Puedes instalar las dependencias con el siguiente comando:
pip install ttkthemes openpyxl mysql-connector-python paramiko

## Uso

1.  **Organizador de Archivos**:
    *   Selecciona un directorio utilizando el botón "Seleccionar Directorio".
    *   Haz clic en "Organizar Archivos" para mover los archivos a las carpetas correspondientes según su tipo.
    *   Utiliza el botón "Deshacer" para revertir la última operación de organización.
    *   Haz clic en "Generar Excel" para crear un reporte en formato Excel de los archivos organizados.
2.  **Base de Datos**:
    *   Introduce la información de conexión a la base de datos (host, puerto, usuario, contraseña, nombre de la base de datos).
    *   Si es necesario, marca la casilla "Conexión SSH" y proporciona la información de conexión SSH.
    *   Haz clic en "Conectar" para establecer la conexión.
    *   Selecciona una tabla del desplegable para visualizar sus registros.
3.  **Configuración**:
    *   Selecciona un tema del desplegable para cambiar la apariencia de la aplicación.

