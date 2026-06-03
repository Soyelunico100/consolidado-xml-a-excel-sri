# Instalacion en Windows

## Requisitos

- Windows 10 u 11.
- Python 3.10 o superior.
- Internet para instalar dependencias la primera vez.

Descarga Python desde:

https://www.python.org/downloads/

Durante la instalacion de Python, marca la opcion **Add python.exe to PATH**.

## Instalar desde GitHub

1. Descarga el repositorio como ZIP desde GitHub.
2. Extrae el ZIP en una carpeta local, por ejemplo:

```text
C:\XML_A_EXCEL_SRI
```

3. Abre la carpeta extraida.
4. Ejecuta `instalar.bat` con doble clic.
5. Espera a que se cree `.venv` y se instalen las dependencias.

## Ejecutar

1. Ejecuta:

```text
abrir_app_web.bat
```

2. Se abrira una pantalla local en el navegador.

3. En la pantalla web, guarda las rutas de:

```text
Python consolidado XML a Excel
Carpeta XML
PDF Compras
PDF Notas de Credito Recibidas
PDF Retenciones Recibidas
PDF Ventas y Retenciones Emitidas
```

4. Presiona:

```text
EXTRAER INFORMACION A EXCEL
```

Tambien puedes ejecutar directamente:

```text
ejecutar.bat
```

5. La primera vez en cada computadora, escribe la clave de instalacion.

6. El Excel se generara en la carpeta:

```text
Python consolidado XML a Excel
```

## Rutas relativas

El programa detecta su carpeta automaticamente usando la ubicacion del archivo Python. No necesita rutas como `C:\Users\...`.

Esto permite mover o descargar el proyecto en otra computadora sin cambiar configuraciones.

## Solucion de problemas

- Si `instalar.bat` falla, revisa que Python este instalado y disponible en PATH.
- Si el Excel no se guarda, cierra cualquier archivo Excel abierto en la carpeta de salida y vuelve a ejecutar.
- Si no aparece ningun dato, confirma que los XML esten dentro de la carpeta XML configurada.

## Subir a GitHub

Sube el contenido de esta carpeta a GitHub. No subas XML reales ni Excel generados. El `.gitignore` ya protege las carpetas de datos locales.
