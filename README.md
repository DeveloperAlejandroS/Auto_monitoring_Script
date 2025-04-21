🛠️ Herramienta de Monitoreo Automatizado - PlayLogger Monitor
    Esta es una herramienta con interfaz gráfica desarrollada en Python que permite procesar, auditar y generar reportes a partir de archivos de monitoreo tipo Play Logger. Su propósito es automatizar el flujo de trabajo en procesos de revisión y control de calidad, facilitando el análisis y validación de datos con base en fechas seleccionadas por el usuario.

✨ Características Principales
    - Interfaz gráfica moderna con CustomTkinter.
    - Carga y validación de archivos Excel de forma dinámica.
    - Selección de rangos de fechas mediante widgets de calendario.
    - Procesamiento automático:
        - Limpieza y transformación de datos.
        - Reemplazo de fórmulas y caracteres Unicode.
        - Generación de columnas auxiliares.
        - Filtrado por fechas y feed index.
    - Auditoría completa incluyendo:
        - Back to Back.
        - Spots.
        - Creativos.
        - Generación de reportes finales organizados por proveedor y marca.
        - Organización automática de archivos por estructura de carpetas basada en fecha.
        - Uso de threading para mantener la interfaz fluida.

🧩 Estructura del Proyecto

📁 Monitoria pruebas/
    │
    ├── main.py                      # Archivo principal que ejecuta la interfaz
    ├── get_BDD.py                   # Módulo para procesamiento de la base de datos
    ├── transform_excel.py           # Funciones para limpiar y transformar datos Excel
    ├── certificado.py               # Generación de columnas tipo certificado
    ├── revision.py                  # Lógica de revisión (b2b, spots, creativos)
    ├── report_generator.py          # Generación del archivo de reporte final
    ├── utils.py                     # Funciones utilitarias (logs, copias, etc.)
    └── README.md                    # Este archivo


🚀 Cómo usar
    1. Ejecuta main.py.
    2. Selecciona la opción "Monitoria" desde el menú lateral.
    3. Escoge una fecha de inicio y una de fin.
    4. Carga el archivo Excel base cuando sea solicitado.
    5. El sistema procesará automáticamente:
    6. Aplicará transformaciones.
    7. Filtrará la BDD.
    8. Generará columnas auxiliares.
    9. Realizará auditoría completa.
    10. Guardará todos los archivos estructurados.
    11. Sigue el proceso a través del estado_texto (log visual).

⚙️ Requisitos:
    - Python 3.12+
    - pandas
    - openpyxl
    - CustomTkinter
    - tkcalendar
    - shutil

Instala dependencias con:

pip install pandas openpyxl customtkinter tkcalendar
📁 Estructura de Carpetas Generadas

📁 /Año/
    └── 📁 Mes/
        └── 📁 PlayLogger[Revision MM DD to MM DD YYYY]
            ├── 📁 Recursos/
            ├── Archivo Final PlayLogger...
            ├── BDD Filtrada...
            └── Reporte Final...
✍️ Autor
    - Alejandro Sierra Vargas
    - https://github.com/DeveloperAlejandroS
