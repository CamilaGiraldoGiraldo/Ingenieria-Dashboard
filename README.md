# CORE IP · Centro de Control Dinámico

Dashboard React/Vite que integra dos fuentes:
- Formato_Bitacora_2026 - copia.xlsx: proyectos, clientes, servicios, bolsa de horas y mesa de ayuda.
- Seguimiento Diario - Sabana (1).xlsx: KPIs de ingeniería, actividades, carga, utilización, capacidad y avance.

## Ejecutar
npm install
npm run dev

## Módulos
- Resumen ejecutivo
- KPIs Ingeniería
- Proyectos
- Clientes
- Servicios
- Bolsa de horas
- Mesa de ayuda
- Bitácora

## Capacidad
7 horas por persona por día laboral. Sábados y domingos no generan capacidad. El KPI de carga usa horas estimadas y el de utilización usa horas productivas.

Los campos de cumplimiento y reproceso no se inventan cuando la fuente no tiene datos.

## Google Drive
El dashboard puede actualizar automáticamente el Seguimiento Diario desde un archivo compartido de Google Drive. Necesitas dos cosas:

1. **El archivo compartido como "Cualquier persona con el enlace" (Lector)** en Google Drive. Copia ese enlace.
2. **Una clave API de Google Cloud** con la API de Google Drive habilitada (Google Cloud Console → APIs y servicios → Credenciales → Crear credenciales → Clave de API). Se recomienda restringirla a la API de Drive.

Pega el enlace y la clave en la barra "Seguimiento desde Google Drive" y pulsa "Guardar fuente". Ambos quedan guardados en el navegador y, al volver a entrar, se intenta consultar la versión actual del mismo archivo. También existe el botón "Actualizar Google Drive".

Debajo hay una segunda barra, "Bitácora comercial desde Google Drive", con su propio campo de enlace (usa la misma clave API de arriba) y botones "Guardar fuente" / "Actualizar".

Para despliegue centralizado (misma fuente para todo el equipo), configura estas variables de entorno de Vite antes de compilar:
- `VITE_GDRIVE_SEGUIMIENTO_URL`: enlace del Seguimiento Diario.
- `VITE_GDRIVE_BITACORA_URL`: enlace de la Bitácora comercial (opcional).
- `VITE_GDRIVE_API_KEY`: la clave API de Google.

El enlace puede pegarse tal cual lo entrega Google Drive (`https://drive.google.com/file/d/ID/view?usp=sharing`) o directamente el ID del archivo.

Si Google Drive rechaza la solicitud (403/404), revisa que el archivo tenga el permiso correcto y que la clave tenga la API de Drive habilitada. El dashboard conserva la última copia válida en el navegador y muestra el error de sincronización en lugar de borrar los datos.
