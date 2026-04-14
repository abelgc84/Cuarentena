# Script de Cuarentena para SharePoint Online

Script en **PowerShell** para auditar bibliotecas de SharePoint Online, identificar archivos que superan un umbral de tamaño, moverlos a una biblioteca de **Cuarentena**, eliminar archivos antiguos de dicha cuarentena y generar un **informe CSV** con el resultado.

Además, el script permite notificar por correo tanto errores de conexión como los movimientos realizados, apoyándose en un script externo de envío de correo mediante Microsoft Graph.

---

## Funcionalidad

Este script realiza las siguientes tareas:

- Recorre una lista de sitios de **SharePoint Online**.
- Se conecta a cada sitio mediante **PnP PowerShell**.
- Revisa la biblioteca **Cuarentena** y elimina archivos antiguos según una fecha de expiración.
- Analiza la biblioteca **Documentos compartidos**.
- Detecta archivos que:
  - superan los **9 MB**;
  - han sido creados dentro del rango de fechas configurado.
- Excluye archivos o carpetas concretas definidos manualmente.
- Mueve los archivos detectados a la biblioteca **Cuarentena**.
- Genera un **CSV** con el detalle de los archivos procesados.
- Registra la ejecución en un archivo de log local.
- Puede enviar:
  - alertas por error de conexión;
  - notificaciones al autor del documento movido;
  - informe final por correo.

---

## Requisitos

- **PowerShell 7** recomendado.
- Módulo **PnP.PowerShell** instalado.
- Acceso a los sitios de SharePoint Online indicados.
- Aplicación registrada en Azure / Entra ID con permisos adecuados para SharePoint.
- Script externo de envío de correo:
  - `C:\sendMailGraph.ps1`

---

## Estructura general del script

El script se divide en varias partes:

### 1. Configuración global

Se definen:

- modo debug;
- credenciales;
- rutas locales;
- fechas de búsqueda y borrado;
- ID de aplicación;
- lista de sitios;
- exclusiones;
- ruta del log.

### 2. Funciones auxiliares

Incluye funciones para:

- escribir en log;
- mover archivos;
- comprobar exclusiones por archivo;
- comprobar exclusiones por carpeta.

### 3. Procesamiento por sitio

Para cada sitio:

- conecta con SharePoint;
- revisa y limpia la biblioteca **Cuarentena**;
- analiza los elementos de **Documentos compartidos**;
- mueve a cuarentena los archivos que cumplan la condición;
- añade los resultados al informe.

### 4. Generación del informe final

Si hubo archivos procesados:

- crea un CSV con fecha;
- opcionalmente lo envía por correo.

---

## Variables principales

### Modo de ejecución

```powershell
$debugMode = $true
