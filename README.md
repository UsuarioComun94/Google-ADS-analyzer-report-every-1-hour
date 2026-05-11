# Google ADS Analyzer Report Every 1 Hour — HMA Demo

## 1. Descripción

Este repositorio contiene una demo técnica del sistema **HMA — Hourly Metrics Automator**.

El objetivo del sistema es automatizar el monitoreo horario de campañas de marketing digital, calcular KPIs principales, generar alertas básicas y entregar reportes descargables.

Actualmente el sistema funciona con **datos simulados**. La arquitectura está preparada para reemplazar el simulador por conectores reales de Google Ads API y Meta Marketing API cuando existan credenciales y permisos de la empresa.

---

## 2. Estado actual

### Implementado

- Ejecución local en Python.
- Ejecución automática con GitHub Actions.
- Generación de reporte Markdown.
- Generación de CSV con métricas por campaña.
- Generación de JSON con resumen ejecutivo.
- Generación de JSON con estado de conexión API.
- Artifact descargable por cada ejecución.
- Nombre de artifact con cliente, fecha, hora y sufijo.
- Zona horaria configurable.
- Umbrales configurables.
- Validación de credenciales faltantes.

### Pendiente

- Conexión real a Google Ads API.
- Conexión real a Meta Marketing API.
- Persistencia histórica externa.
- Alertas externas por email, Slack, Telegram o webhook.
- Dashboard.
- Validación contra datos reales de plataforma.

---

## 3. Fuente de datos

La fuente actual es:

```text
simulated_data