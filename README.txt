QA MANTO - APLICACIÓN WEB (CONEXIÓN DIRECTA POR ENLACE DE ONEDRIVE)
====================================================================

CÓMO FUNCIONA
-------------
No requiere registro en Azure ni iniciar sesión en OneDrive:
1. En tu OneDrive Personal, haz clic derecho sobre el Excel maestro -> Compartir -> Copiar vínculo.
   (Asegúrate de que el permiso diga "Cualquier persona que tenga el vínculo").
2. Entra al módulo "Proveedor" en QA MANTO y pega el enlace (queda guardado automáticamente en tu navegador).
3. Puedes hacer clic en "Verificar enlace" para comprobar que descarga bien y tiene las 41 columnas.
4. Selecciona tu Excel base (.xlsx) desde tu computadora y presiona "Sincronizar con enlace de OneDrive".
5. El sistema descarga el archivo desde OneDrive, crea un respaldo en la carpeta backups/, aplica los cambios
   (Agregar, Actualizar, Eliminar) y descarga automáticamente el Excel actualizado listo para reemplazar en OneDrive.

EJECUTAR EN LOCAL
-----------------
Desde la carpeta QA_MANTO_Web:

npm install
npm start

Abrir en el navegador:
http://localhost:3000
