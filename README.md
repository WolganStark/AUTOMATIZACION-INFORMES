# 🚀 Automatización de Base de Datos Clínico-Administrativa con VBA (Excel)

## Un Sistema Inteligente de Gestión de Eventos y Generación de Informes

Este proyecto resuelve el desafío de gestionar datos de pacientes con múltiples eventos en una estructura de **"una fila por paciente"** (formato ancho) y transforma esa información en **informes gerenciales y de seguimiento** listos para el análisis. La solución elimina horas de trabajo manual en la preparación de datos y la creación de Tablas Dinámicas. 

---

## ✨ Caracteristicas Principales
* **Simplificación de Datos (Normalización):** El núcleo del proyecto es la macro `BuildEventDetail`, que automáticamente convierte el formato ancho de la base de datos (multiples columnas de eventos por paciente) a un formato largo (`Eventos_Detallados`). Esto crea una fuente de datos estructurada y optimizada para el análisis.
* **Generación de informes con un Clic:** El personal puede generar informes mensuales, trimestrales y anuales con solo pulsar un botón en la hoja `MENU`.
    * Genera el informe del **Mes Anterior** (Ej: `Informe_2025-08`).
    * Genera el informe del **Trimestre Anterior** (Ej: `Informe_T3_2025`).
    * Genera el informe del **Año Anterior** (Ej: `Informe_2024`).
* **Reportes Estándar:** Los informes generados son Tablas Dinámicas preconfiguradas que muestran el **Conteo de Eventos** desglosado por **Tipo de Evento** (Filas) y **Fase del Evento** (Columnas), listas para el seguimiento y la toma de decisiones.
* **Mantenimiento Sencillo:** El código es modular y fácil de configurar a través de las constantes de VBA (`TABLE_NAME`, `EVENTS_SHEET`, `MAX_EVENTS`).