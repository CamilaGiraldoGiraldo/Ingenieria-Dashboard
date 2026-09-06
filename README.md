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
