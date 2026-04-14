<#
.SYNOPSIS
    Script para auditar el almacenamiento en Sites de Sharepoint.

.DESCRIPTION
    Este script escanea sitios de SharePoint, mueve los archivos que superan el límite de tamaño
    y genera un informe CSV, además de limpiar archivos antiguos de la cuarentena.
#>

Import-Module PnP.Powershell

# --- Configuración Global y Debug ---
$debugMode = $true  # Cambiar a $false para ejecución real. Si es $true, no se mueven/borran archivos ni se envían correos.

# --- Configuración de Credenciales ---
$correoAdmin = "admin365@contoso.onmicrosoft.com"
$credencialesAdmin = ConvertTo-SecureString "**********" -AsPlainText -Force
$credencialesAdminConvertido = New-Object System.Management.Automation.PSCredential ($correoAdmin, $credencialesAdmin)

# Directorio local para los archivos CSV
$directorioCSV = "C:\Informe Cuarentena"
if (-not (Test-Path -Path $directorioCSV)) {
    New-Item -ItemType Directory -Path $directorioCSV | Out-Null
}

# Configuración de fechas
$fechaBusqueda = (Get-Date).AddDays(-21)    # Fecha para auditar archivos 
$fechaBorrado = (Get-Date).AddDays(-30)     # Fecha para eliminar archivos de la cuarentena
$idAplicacion = "*****-****-****-****"      # ID de la app para la conexión a sharepoint

# Umbral de tamaño (9 MB)
$tamanoMaximoBytes = 9MB

# Lista de sitios de SharePoint
$listaSites = @(
    "https://contoso.sharepoint.com/sites/PNBSENCURSO",
    "https://contoso.sharepoint.com/sites/CLIENTES",
    "https://contoso.sharepoint.com/sites/CARPETASEQUIPO"
)

# Exclusiones
$archivosExcluidos = @(
    "/sites/CLIENTES/Documentos compartidos/España/Presentación_Diversity Essentials.pptx",
    "/sites/PNBSENCURSO/Documentos compartidos/España/Propuesta técnica_Escuela de Liderazgo España_Iberdrola.pptx"
)

$carpetasExcluidas = @(
    "/sites/CLIENTES/Documentos compartidos/España/1. DESARROLLO DEL PROYECTO"
)

$LogPath = "C:\cuarentenaSharepoint_Log.txt"
$logDirectory = Split-Path -Path $LogPath -Parent
if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

# --- Funciones Auxiliares ---

function Write-LocalLog {
    param (
        [string]$Message,
        [ValidateSet("INFO", "ERROR", "EXITO")]
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "[$Timestamp] [$Level] $Message"
    
    try {
        $Entry | Out-File -FilePath $LogPath -Append -Encoding utf8 -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "No se pudo escribir en el archivo de log: $LogPath"
    }
}

function moverArchivo {
    param (
        [string]$fileRef,
        [string]$cuarentena,
        [string]$autor
    )
    Try {
        if ($debugMode) {
            Write-Host "[DEBUG] Simulación: Se movería el archivo $fileRef a la carpeta $cuarentena" -ForegroundColor Yellow
            Write-LocalLog -Message "[DEBUG-SIM] Movimiento simulado: $fileRef" -Level "INFO"
            return $true
        }

        Write-Host "Moviendo Archivo: " -f Blue -NoNewline; Write-Host $fileRef -f Green -NoNewline; Write-Host " a: " -f Blue -NoNewline; Write-Host $cuarentena -f Green            
        Move-PnPFile -SourceUrl $fileRef -TargetUrl $cuarentena -Force
        Write-LocalLog -Message "Movido: $fileRef | Destino: $cuarentena | Autor: $autor" -Level "EXITO"
        return $true                 
    }
    Catch {
        $errText = $_.Exception.Message
        $isFalseMoveError = ($errText -match 'SPMigrationQosException') -and ($errText -match '0x80070002' -or $errText -match 'cannot find the file specified')

        if ($isFalseMoveError) {
            Write-Host -ForegroundColor Yellow "Aviso: El archivo ya no existe en el origen (se asume como movido)."
            Write-LocalLog -Message "Aviso: El archivo $fileRef ya no existe en origen (asumido como movido)." -Level "INFO"
            return $true
        }            

        Write-Host -f Red "Error al mover el archivo '$fileRef': $_"
        Write-LocalLog -Message "Error crítico al mover: $fileRef | Error: $_" -Level "ERROR"
        return $false
    }
}

function esArchivoExcluido {
    param ([string]$fileRef)
    return $archivosExcluidos -contains $fileRef
}

function esCarpetaExcluida {
    param ([string]$fileRef)
    foreach ($carpeta in $carpetasExcluidas) {
        if ($fileRef -like "$carpeta/*" -or "$fileRef" -eq "$carpeta") { return $true }
    }
    return $false
}

# --- Inicio del Script ---

if ($debugMode) {
    Write-Host "***************************************************" -ForegroundColor Yellow
    Write-Host "MODO DEBUG ACTIVADO: No se realizarán cambios reales" -ForegroundColor Yellow
    Write-Host "***************************************************" -ForegroundColor Yellow
    Write-LocalLog -Message "--- INICIO DE EJECUCIÓN EN MODO DEBUG ---" -Level "INFO"
}

$resultados = New-Object System.Collections.Generic.List[PSObject]

foreach ($siteUrl in $listaSites) {
    try {
        try {
            Connect-PnPOnline -Url $siteUrl -ClientId $idAplicacion -Credentials $credencialesAdminConvertido
            Write-Host -f Blue "Conectado al sitio: $siteUrl"
            Write-LocalLog -Message "Conectado al sitio: $siteUrl" -Level "INFO"
        }
        catch {
            $errorConexion = $_.Exception.Message
            Write-Host "Error de conexión al sitio $siteUrl : $errorConexion" -ForegroundColor Red
            Write-LocalLog -Message "Error de conexión al sitio $siteUrl : $errorConexion" -Level "ERROR"

            if ($debugMode) {
                Write-Host "[DEBUG] Simulación: Se enviaría alerta de error de conexión por correo." -ForegroundColor Yellow
            } else {
                $htmlBodyErr = "<h3>Error Crítico de Conexión</h3><br/>No se pudo conectar al sitio SharePoint: <b>$siteUrl</b>.<br/><br/><b>Detalle técnico:</b> $errorConexion"
                pwsh.exe -ExecutionPolicy Bypass -File "C:\sendMailGraph.ps1" `
                    -ToUser "usuario@correo.com" `
                    -FromUser "admin365@contoso.onmicrosoft.com" `
                    -Subject "ALERTA: Fallo de conexión SharePoint - $siteUrl" `
                    -HtmlBody $htmlBodyErr
            }
            continue # Salta al siguiente sitio de la lista
        }

        # 1. Limpieza de la Cuarentena
        Write-Host -f Cyan "Verificando archivos antiguos en la Cuarentena..."
        $filesCuarentena = Get-PnPListItem -List "Cuarentena" -Fields "FileRef", "Modified", "File_x0020_Type" -ErrorAction SilentlyContinue
        
        if ($null -ne $filesCuarentena) {
            foreach ($file in $filesCuarentena) {
                $fileUrl = $file["FileRef"]
                $fileModified = [DateTime]$file["Modified"]
                if ($fileModified -lt $fechaBorrado -and $null -ne $file["File_x0020_Type"]) {
                    if ($debugMode) {
                        Write-Host "[DEBUG] Simulación: Se eliminaría por expiración: $fileUrl" -ForegroundColor Yellow
                    } else {
                        Write-Host -f Red "Eliminando archivo expirado: $fileUrl"
                        Remove-PnPFile -ServerRelativeUrl $fileUrl -Force -Recycle
                        Write-LocalLog -Message "Eliminado por expiración: $fileUrl" -Level "EXITO"
                    }
                }
            }
        }

        # 2. Procesamiento de Documentos Compartidos
        $procesados = New-Object System.Collections.Generic.HashSet[string]
        $datos = Get-PnPListItem -List "Documentos compartidos" -PageSize 500

        foreach ($item in $datos) {
            $fileRef = $item.FieldValues['FileRef']
            $nombreFichero = $item.FieldValues['FileLeafRef']
            
            if (-not $procesados.Add($fileRef)) { continue }

            $fileSizeVal = if ($item.FieldValues.File_x0020_Size) { $item.FieldValues.File_x0020_Size } else { 0 }
            $fileSizeKB = [Math]::Round(($fileSizeVal / 1KB), 2)
            $fileCreationDate = [DateTime]$item.FieldValues['Created']

            if (esArchivoExcluido -fileRef $fileRef -or esCarpetaExcluida -fileRef $fileRef) {
                Write-Host -f Yellow "Ignorado (Excluido): $fileRef"
                continue
            }

            if ($fileSizeKB -gt $tamanoMaximoBytes -and $fileCreationDate -ge $fechaBusqueda) {
                
                $autorObj = $item["Author"]
                $autorEmail = $autorObj.Email
                $autorLookup = $autorObj.LookupValue

                if ([string]::IsNullOrEmpty($autorEmail) -and $autorObj.LookupId -gt 0) {
                    try {
                        $usuarioCompleto = Get-PnPUser -Identity $autorObj.LookupId
                        $autorEmail = $usuarioCompleto.Email
                    } catch { }
                }

                $cuarentena = "Cuarentena"  

                if (moverArchivo -fileRef $fileRef -cuarentena $cuarentena -autor $autorLookup) {
                    $resObj = [PSCustomObject]@{
                        "Nombre del fichero" = $nombreFichero
                        "Ubicación"          = $fileRef
                        "Creado por"         = $autorEmail
                        "Fecha de creación"  = $item.FieldValues['Created_x0020_Date']
                        "Tamaño (KB)"        = $fileSizeKB
                        "Movido a"           = $cuarentena
                    }
                    $resultados.Add($resObj)

                    if ($debugMode) {
                        Write-Host "[DEBUG] Simulación: Se enviaría notificación a $autorEmail" -ForegroundColor Yellow
                    } else {
                        $htmlBody = @"
Hola $($autorLookup),<br/><br/>
El documento <b>$($nombreFichero)</b> supera el límite de 9 MB y ha sido movido a la carpeta de <b>Cuarentena</b>.<br/>
Por favor, optimiza las imágenes y devuélvelo a su ubicación original una vez reducido su tamaño.<br/><br/>
Atentamente,<br/>Equipo de Sistemas.
"@
                        pwsh.exe -ExecutionPolicy Bypass -File "C:\sendMailGraph.ps1" `
                            -ToUser $autorEmail -FromUser "admin365@contoso.onmicrosoft.com" `
                            -Subject "Documento en Cuarentena: $nombreFichero" -HtmlBody $htmlBody `
                            -CcUser "usuario@correo.com"
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error al procesar el sitio $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
        Write-LocalLog -Message "Error al procesar el sitio $siteUrl : $($_.Exception.Message)" -Level "ERROR"
    }
}

# --- Generación del Informe Final ---

if ($resultados.Count -gt 0) {
    $fechaActual = (Get-Date).ToString("yyyyMMdd")
    $rutaCSV = Join-Path -Path $directorioCSV -ChildPath "Informe_Cuarentena_${fechaActual}.csv"

    Write-Host -f Blue "Exportando informe CSV..."
    $resultados | Export-Csv -Path $rutaCSV -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    if ($debugMode) {
        Write-Host "[DEBUG] Simulación: Se enviaría el informe final por correo a administración." -ForegroundColor Yellow
        Write-LocalLog -Message "Ejecución DEBUG finalizada. Archivos identificados: $($resultados.Count)" -Level "INFO"
    } else {
        pwsh.exe -ExecutionPolicy Bypass -File "C:\sendMailGraph.ps1" `
            -ToUser "usuario@correo.com,usuario2@correo.com" `
            -FromUser "admin365@contoso.onmicrosoft.com" `
            -Subject "Informe de Cuarentena SharePoint" `
            -HtmlBody "<h3>Se adjunta el informe detallado de los archivos movidos hoy.</h3>" `
            -AttachmentPath $rutaCSV

        Write-LocalLog -Message "Ejecución finalizada con éxito. Archivos movidos: $($resultados.Count)" -Level "EXITO"
    }
} else {
    Write-LocalLog -Message "Ejecución finalizada. No se encontraron archivos pendientes." -Level "INFO"
}

Disconnect-PnPOnline
