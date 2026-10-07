# Google Sheets + Apps Script

1. Crear una Google Sheet nueva.
2. Abrir Extensiones -> Apps Script.
3. Copiar apps-script/Code.gs a Code.gs.
4. Cambiar CAMBIAR-ESTA-CLAVE por una clave editorial propia.
5. Si Apps Script está vinculado a la Sheet, dejar spreadsheetId vacío.
6. Ejecutar setup una vez y autorizar. Se crea la hoja Comunicados.
7. Implementar -> Nueva implementación -> Aplicación web.
8. Ejecutar como: Yo.
9. Quién tiene acceso: Cualquiera.
10. Copiar la URL /exec y colocarla en config.js como apiUrl.

GET público:
- action=ping
- action=list
- action=get&id=ID

POST editorial protegido por key:
- save
- publish
- unpublish
- delete

La clave editorial no se guarda en el navegador. El panel la solicita para las operaciones de escritura. Para una etapa posterior conviene sustituirla por autenticación de Google y usuarios autorizados.