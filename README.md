# CRM Admisiones, Familias y Entrevistas (Modelo REA)

Aplicación web estática (Local-First) para la gestión de admisiones escolares, embudo Kanban, calendario de visitas/entrevistas, actas y control de plazas.

## Despliegue rápido en GitHub + Vercel

1. Sube el contenido de esta carpeta (`index.html` y `vercel.json`) a un repositorio nuevo en **GitHub** (por ejemplo, `crm-admisiones-colegios`).
2. Entra en **Vercel** → **Add New Project** → selecciona ese repositorio.
3. Framework Preset: **Other** (sin comando de build, directorio raíz `./`) → pulsa **Deploy**.

En unos segundos tendrás una URL pública (`https://tu-proyecto.vercel.app`) lista para compartir con otros colegios. Cada colegio guarda sus propios datos de forma privada en su navegador (`localStorage`) y puede configurarse desde la pestaña **Administración**.
