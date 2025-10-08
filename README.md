<div align="center">
  <h3>📨 Massive Mail Sender 📨</h3>
  <p>Aplicación de escritorio desarrollada en Python diseñada específicamente para facilitar a equipos de marketing el envío masivo de correos electrónicos. Cuenta con un editor enriquecido que permite la edición visual del contenido de los emails, facilitando la incorporación de texto formateado, enlaces y otros elementos multimedia.</p>
</div>
<img src="/public/Demo.png" alt="Demo"/>

## Principales funcionalidades
- Carga de destinatarios desde: Excel (.xlsx), CSV, PDF (extrae emails por texto) o Google Sheets.
- Editor enriquecido para Asunto y Mensaje: negrita, cursiva, subrayado, color, tamaño, centrar, listas y enlaces.
- Plantillas visuales: Vanilla (sin envoltura), Promocional, Corporativa, Minimalista, Navidad y Halloween.
- Imágenes inline: inserción por placeholders `{{image1}}`, `{{image2}}`, etc. y selector de archivos.
- Panel SMTP: email, contraseña, servidor, puerto (587/465), TLS y SSL (autoajuste según puerto).



## Uso de la aplicación
1. Carga de destinatarios
   - Excel/CSV/PDF: el archivo debe tener la primera columna llamada `email`. En PDF se extraen correos por texto.
   - Google Sheets: coloca `credentials.json` en la carpeta del proyecto y selecciona la hoja por ID.
2. Escribe Asunto y Mensaje
   - Usa la barra de herramientas para aplicar estilos. El Asunto admite varias líneas; para el envío se colapsan a espacios.
3. Imágenes
   - Inserta un placeholder en el texto (`Insertar imagen` → `{{imageN}}`) y luego agrega archivos con “Agregar imágenes…”. En el envío se enlazan como CID.
4. Selecciona plantilla
   - `Vanilla` no agrega HTML extra; el resto aplica headers y estilos predefinidos.
5. SMTP
   - Completa email/contraseña/servidor. El puerto admite `587` (TLS) o `465` (SSL). Al cambiar el puerto se ajusta automáticamente TLS/SSL.
6. Enviar
   - Confirma el envío. Se muestra progreso y resumen (exitosos/fallidos).


## Licencia
Aplicacion de uso interno. Ajusta a las necesidades de tu organización.
