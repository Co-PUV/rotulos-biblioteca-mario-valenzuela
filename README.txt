# Rótulos Biblioteca de Filosofía y Teología Mario Valenzuela, S.J.

## Estructura

- `index.html` — estructura visual de la aplicación.
- `styles.css` — colores y diseño.
- `app.js` — coordinación general de las pestañas.
- `excel-reader.js` — reservado para futuras mejoras específicas de lectura.
- `word-generator.js` — reservado para futuras mejoras específicas de Word.
- `manifest.json` — configuración de instalación PWA.
- `service-worker.js` — caché y actualizaciones.
- `icons/` — iconos de la aplicación.

## Para probar

No abra la PWA directamente como archivo local si quiere probar la instalación.
Use un servidor web o publíquela en un hosting HTTPS.

## Para actualizar

1. Modifique el archivo correspondiente.
2. Cambie el número de versión en `service-worker.js`, por ejemplo:
   `rotulos-biblioteca-v1` → `rotulos-biblioteca-v2`
3. Publique nuevamente los archivos.

Esto obliga a la PWA a detectar la nueva versión.

## Importante

La lógica actual de lectura de Excel y generación de Word se conserva en `app.js`
para mantener el funcionamiento probado. En una siguiente fase se puede separar
completamente en `excel-reader.js` y `word-generator.js` sin cambiar la interfaz.
