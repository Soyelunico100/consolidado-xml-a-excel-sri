# Consolidado XML a Excel

Programa local en Python para convertir XML autorizados del SRI Ecuador a un archivo Excel consolidado.

## Uso rapido

1. Ejecuta `instalar.bat` una sola vez.
2. Copia tus XML reales dentro de `entrada_xml`.
3. Ejecuta `ejecutar.bat`.
4. Escribe la clave de instalacion cuando el programa la pida por primera vez.
5. Revisa el Excel generado en `salida_excel`.

Al terminar, los XML procesados correctamente se mueven a `procesados`. Los XML no reconocidos o con error se mueven a `errores`.

## Carpetas

- `entrada_xml`: aqui se colocan los XML del SRI.
- `salida_excel`: aqui se genera el Excel.
- `procesados`: aqui se mueven los XML que ya fueron procesados.
- `errores`: aqui se mueven los XML con problema.
- `ejemplos`: contiene XML ficticios para pruebas. No contiene informacion real.

## Ejecucion

El flujo recomendado es `ejecutar.bat`. La interfaz Streamlit queda fuera de la instalacion normal para evitar errores de dependencias en Windows.

## Privacidad

No subas XML reales ni Excel generados a GitHub. El archivo `.gitignore` ya ignora:

- `.venv`
- `__pycache__`
- XML dentro de `entrada_xml`, `procesados` y `errores`
- Excel generados
- claves, certificados y archivos privados

El programa usa la misma clave de instalacion del programa anterior. La clave no esta escrita en texto dentro del codigo; solo se guarda un hash para validarla.

## Prueba con XML ficticio

Para probar sin datos reales:

1. Copia `ejemplos\factura_ficticia.xml` dentro de `entrada_xml`.
2. Ejecuta `ejecutar.bat`.
3. Confirma que se cree un Excel dentro de `salida_excel`.
