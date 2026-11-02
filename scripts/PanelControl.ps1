#requires -Version 5.1
<#
.SYNOPSIS
    Panel de control local para desarrollar, comprobar y abrir ONAT · Disciplina de pago.
.DESCRIPTION
    No modifica datos fiscales ni realiza publicaciones. Ejecuta tareas npm del repositorio
    y puede iniciar los servidores locales HTTPS/HTTP en ventanas separadas.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$script:IsWindows = $env:OS -eq 'Windows_NT'
$script:NodeCommand = Get-Command node -ErrorAction SilentlyContinue
$script:NpmCommand = Get-Command npm.cmd -ErrorAction SilentlyContinue
if (-not $script:NpmCommand) {
    $script:NpmCommand = Get-Command npm -ErrorAction SilentlyContinue
}
$script:NpmExecutable = $null
if ($script:NpmCommand) {
    $script:NpmExecutable = $script:NpmCommand.Source
    if ([string]::IsNullOrWhiteSpace($script:NpmExecutable)) {
        $script:NpmExecutable = $script:NpmCommand.Definition
    }
}

$script:ImprovementIdeas = @(
    '1. Validar la importación: duplicados, filas incompletas, fechas y períodos anómalos antes del análisis.',
    '2. Mapear encabezados: guardar equivalencias entre nombres de columnas y elegir qué campos incluir.',
    '3. Conciliar contribuyentes: revisar coincidencias, duplicados y pagos sin correspondencia con el Patrón.',
    '4. Analizar el Patrón: filtrar por periodicidad, valor del tributo, antigüedad y atributos adicionales.',
    '5. Comparar períodos: mostrar cambios mensuales/anuales de importe, puntualidad y pagos pendientes.',
    '6. Trazabilidad: explicar cada estado con vencimiento, días de atraso, fila y hoja del Excel de origen.',
    '7. Revisar casos: añadir notas y marcas de seguimiento, con respaldo local descargable.',
    '8. Mejorar accesibilidad y rendimiento: teclado, lectores de pantalla y paginación de tablas grandes.',
    '9. Colaboración protegida: perfiles y permisos con backend seguro; nunca publicar datos fiscales en Pages.',
    '10. Alertas locales: recordar vencimientos próximos sin transmitir datos.',
    '11. Plantillas de informes: guardar filtros, columnas y formatos por tributo o periodicidad.',
    '12. Modo presentación: mostrar indicadores agregados y ocultar nombres/códigos.'
)

function Write-PanelHeader {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host '       ONAT · PANEL DE CONTROL DE DESARROLLO' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host ('Repositorio: ' + $script:RepoRoot) -ForegroundColor DarkGray
    $nodeState = if ($script:NodeCommand) { 'disponible' } else { 'no encontrado' }
    $npmState = if ($script:NpmExecutable) { 'disponible' } else { 'no encontrado' }
    Write-Host ('Node.js: ' + $nodeState + '    npm: ' + $npmState) -ForegroundColor DarkGray
    Write-Host ''
}

function Wait-ForUser {
    [void](Read-Host 'Presiona Enter para volver al Panel de Control')
}

function Test-LocalPort {
    param([Parameter(Mandatory = $true)][int]$Port)

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $pending = $client.BeginConnect('127.0.0.1', $Port, $null, $null)
        if (-not $pending.AsyncWaitHandle.WaitOne(350)) {
            return $false
        }
        $client.EndConnect($pending)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Close()
    }
}

function Open-LocalPage {
    param([Parameter(Mandatory = $true)][string]$Url)

    if ($script:IsWindows) {
        Start-Process -FilePath $Url
    }
    else {
        Write-Host ('Abre esta dirección en el navegador: ' + $Url) -ForegroundColor Cyan
    }
}

function Start-LocalServer {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('https', 'preview')][string]$Mode
    )

    if (-not $script:NpmExecutable) {
        Write-Host 'No se encontró npm. Instala Node.js 20.19+ o 22.12+ y vuelve a intentarlo.' -ForegroundColor Red
        Wait-ForUser
        return
    }

    if ($Mode -eq 'https') {
        $port = 5173
        $url = 'https://localhost:5173/'
        $commandLine = 'npm run dev'
        $npmArguments = @('run', 'dev')
        $description = 'servidor local HTTPS'
    }
    else {
        $port = 4173
        $url = 'http://localhost:4173/'
        $commandLine = 'npm run dev:preview -- --port 4173'
        $npmArguments = @('run', 'dev:preview', '--', '--port', '4173')
        $description = 'vista previa HTTP'
    }

    if (Test-LocalPort -Port $port) {
        Write-Host ('Ya hay un servicio escuchando en el puerto ' + $port + '.') -ForegroundColor Yellow
    }
    else {
        Write-Host ('Iniciando ' + $description + ' en una ventana nueva...') -ForegroundColor Cyan
        try {
            if ($script:IsWindows) {
                Start-Process -FilePath $env:ComSpec -ArgumentList @('/c', $commandLine) -WorkingDirectory $script:RepoRoot | Out-Null
            }
            else {
                Start-Process -FilePath $script:NpmExecutable -ArgumentList $npmArguments -WorkingDirectory $script:RepoRoot | Out-Null
            }
        }
        catch {
            Write-Host ('No se pudo iniciar el servidor: ' + $_.Exception.Message) -ForegroundColor Red
            Wait-ForUser
            return
        }

        $ready = $false
        for ($attempt = 0; $attempt -lt 15; $attempt += 1) {
            if (Test-LocalPort -Port $port) {
                $ready = $true
                break
            }
            Start-Sleep -Seconds 1
        }
        if (-not $ready) {
            Write-Host 'El servidor aún no responde. Revisa la ventana de npm para ver el error.' -ForegroundColor Yellow
            Wait-ForUser
            return
        }
    }

    Write-Host ('Abriendo ' + $url) -ForegroundColor Green
    if ($Mode -eq 'https') {
        Write-Host 'El certificado local es autofirmado; el navegador puede pedirte que aceptes la excepción.' -ForegroundColor Yellow
    }
    Open-LocalPage -Url $url
    Write-Host 'Para detener el servidor, cierra la ventana de terminal que se abrió.' -ForegroundColor DarkGray
    Wait-ForUser
}

function Invoke-NpmTask {
    param(
        [Parameter(Mandatory = $true)][string[]]$NpmArguments,
        [Parameter(Mandatory = $true)][string]$TaskName
    )

    if (-not $script:NpmExecutable) {
        Write-Host 'No se encontró npm. Instala Node.js y vuelve a intentarlo.' -ForegroundColor Red
        Wait-ForUser
        return
    }

    Write-Host ('Ejecutando ' + $TaskName + '...') -ForegroundColor Cyan
    Push-Location -LiteralPath $script:RepoRoot
    try {
        & $script:NpmExecutable @NpmArguments
        $exitCode = $LASTEXITCODE
    }
    catch {
        $exitCode = 1
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    finally {
        Pop-Location
    }

    if ($exitCode -eq 0) {
        Write-Host ($TaskName + ' terminó correctamente.') -ForegroundColor Green
    }
    else {
        Write-Host ($TaskName + ' terminó con código ' + $exitCode + '.') -ForegroundColor Red
    }
    Wait-ForUser
}

function Show-ImprovementOptions {
    Write-Host ''
    Write-Host 'IDEAS PARA PERFECCIONAR LA WEB, UNA ETAPA A LA VEZ' -ForegroundColor Cyan
    Write-Host ''
    foreach ($idea in $script:ImprovementIdeas) {
        Write-Host ('  ' + $idea) -ForegroundColor Gray
    }
    Write-Host ''
    Write-Host 'Siguiente paso recomendado: validar y conciliar la importación antes del análisis.' -ForegroundColor Green
    Wait-ForUser
}

function Show-ControlPanel {
    while ($true) {
        Write-PanelHeader
        Write-Host '  1. Iniciar/abrir la Web segura (HTTPS · puerto 5173)'
        Write-Host '  2. Iniciar/abrir la vista previa (HTTP · puerto 4173)'
        Write-Host '  3. Instalar dependencias (npm ci)'
        Write-Host '  4. Ejecutar pruebas (npm test)'
        Write-Host '  5. Comprobación completa (npm run check)'
        Write-Host '  6. Generar compilación de producción (npm run build)'
        Write-Host '  7. Revisar dependencias (npm audit)'
        Write-Host '  8. Mostrar opciones de mejora continua'
        Write-Host '  0. Salir'
        Write-Host ''
        $choice = (Read-Host 'Selecciona una opción').Trim()

        switch ($choice) {
            '1' { Start-LocalServer -Mode 'https' }
            '2' { Start-LocalServer -Mode 'preview' }
            '3' { Invoke-NpmTask -NpmArguments @('ci') -TaskName 'Instalación de dependencias' }
            '4' { Invoke-NpmTask -NpmArguments @('test') -TaskName 'Pruebas' }
            '5' { Invoke-NpmTask -NpmArguments @('run', 'check') -TaskName 'Comprobación completa' }
            '6' { Invoke-NpmTask -NpmArguments @('run', 'build') -TaskName 'Compilación de producción' }
            '7' { Invoke-NpmTask -NpmArguments @('audit') -TaskName 'Auditoría de dependencias' }
            '8' { Show-ImprovementOptions }
            '0' { return }
            default {
                Write-Host 'Opción no válida.' -ForegroundColor Yellow
                Start-Sleep -Milliseconds 900
            }
        }
    }
}

Show-ControlPanel
