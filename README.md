# Consolidado XML a Excel

Programa local en Python para convertir XML autorizados del SRI Ecuador a un archivo Excel consolidado.

## Uso rapido

1. Ejecuta `instalar.bat` una sola vez.
2. Ejecuta `abrir_app_web.bat`.
3. En la pantalla web local, carga tus XML.
4. Presiona `EXTRAER INFORMACION A EXCEL`.
5. Escribe la clave de instalacion cuando el programa la pida por primera vez.
6. Revisa o descarga el Excel generado en `salida_excel`.

Al terminar, los XML procesados correctamente se mueven a `procesados`. Los XML no reconocidos o con error se mueven a `errores`.

## Carpetas

- `entrada_xml`: aqui se colocan los XML del SRI.
- `salida_excel`: aqui se genera el Excel.
- `procesados`: aqui se mueven los XML que ya fueron procesados.
- `errores`: aqui se mueven los XML con problema.
- `ejemplos`: contiene XML ficticios para pruebas. No contiene informacion real.

## App web local

El flujo recomendado es `abrir_app_web.bat`. Abre una pagina local en el navegador:

```text
http://127.0.0.1:8765
```

Esa pagina corre solo en tu computadora. Los XML no se envian a internet.

Si prefieres el modo directo sin pantalla web, usa `ejecutar.bat`.

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

1. Ejecuta `abrir_app_web.bat`.
2. Carga `ejemplos\factura_ficticia.xml` desde la pantalla web.
3. Presiona `EXTRAER INFORMACION A EXCEL`.
4. Confirma que se cree un Excel dentro de `salida_excel`.
