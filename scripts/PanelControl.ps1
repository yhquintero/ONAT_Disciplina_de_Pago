#requires -Version 5.1
<#
.SYNOPSIS
    Interfaz gráfica local para desarrollar y comprobar ONAT · Disciplina de pago.
.DESCRIPTION
    Abre un Panel de Control Windows Forms. Inicia la Web HTTPS local y la vista previa,
    ejecuta tareas npm en segundo plano y presenta su actividad en la propia interfaz.
    No modifica los Excel importados ni publica cambios.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$script:IsWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
if (-not $script:IsWindows) {
    throw 'El Panel de Control visual requiere Windows y Windows Forms. Ejecútalo en Windows PowerShell 5.1 o PowerShell 7 para Windows.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

$script:Palette = @{
    Navy       = [System.Drawing.Color]::FromArgb(18, 43, 60)
    NavyLight  = [System.Drawing.Color]::FromArgb(27, 59, 77)
    Teal       = [System.Drawing.Color]::FromArgb(46, 157, 148)
    Blue       = [System.Drawing.Color]::FromArgb(42, 119, 181)
    Ink        = [System.Drawing.Color]::FromArgb(31, 49, 66)
    Muted      = [System.Drawing.Color]::FromArgb(119, 136, 150)
    Line       = [System.Drawing.Color]::FromArgb(226, 233, 238)
    Canvas     = [System.Drawing.Color]::FromArgb(244, 247, 250)
    White      = [System.Drawing.Color]::White
    Green      = [System.Drawing.Color]::FromArgb(48, 139, 112)
    GreenLight = [System.Drawing.Color]::FromArgb(232, 246, 239)
    Amber      = [System.Drawing.Color]::FromArgb(166, 119, 50)
    AmberLight = [System.Drawing.Color]::FromArgb(250, 244, 232)
    Red        = [System.Drawing.Color]::FromArgb(171, 81, 76)
    RedLight   = [System.Drawing.Color]::FromArgb(251, 239, 237)
}

$script:ImprovementIdeas = @(
    'Validar la importación: detectar duplicados, filas incompletas, fechas y períodos anómalos antes del análisis.',
    'Mapear encabezados: guardar equivalencias entre nombres de columnas y elegir qué campos incluir.',
    'Conciliar contribuyentes: revisar coincidencias, duplicados y pagos sin correspondencia con el Patrón.',
    'Analizar el Patrón: filtrar por periodicidad, valor del tributo, antigüedad y atributos adicionales.',
    'Comparar períodos: mostrar cambios mensuales/anuales de importe, puntualidad y pagos pendientes.',
    'Trazabilidad: explicar cada estado con vencimiento, días de atraso, fila y hoja del Excel de origen.',
    'Revisar casos: añadir notas y marcas de seguimiento, con respaldo local descargable.',
    'Mejorar accesibilidad y rendimiento: teclado, lectores de pantalla y paginación de tablas grandes.',
    'Colaboración protegida: perfiles y permisos con backend seguro; nunca publicar datos fiscales en Pages.',
    'Alertas locales: recordar vencimientos próximos sin transmitir datos.',
    'Plantillas de informes: guardar filtros, columnas y formatos por tributo o periodicidad.',
    'Modo presentación: mostrar indicadores agregados y ocultar nombres/códigos.'
)

$script:Form = $null
$script:OutputBox = $null
$script:StatusBadge = $null
$script:StatusText = $null
$script:NodeBadge = $null
$script:NpmBadge = $null
$script:RepoPathLabel = $null
$script:StopServersButton = $null
$script:PrivacyCard = $null
$script:TaskButtons = @()
$script:TaskWorker = $null
$script:ServerProcesses = @{}
$script:ServerPollTimer = $null

function New-UiLabel {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [float]$FontSize = 9,
        [bool]$Bold = $false,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::FromArgb(31, 49, 66),
        [bool]$AutoSize = $true
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $AutoSize
    $label.ForeColor = $ForeColor
    $style = [System.Drawing.FontStyle]::Regular
    if ($Bold) { $style = [System.Drawing.FontStyle]::Bold }
    $label.Font = [System.Drawing.Font]::new('Segoe UI', $FontSize, $style)
    $label.BackColor = [System.Drawing.Color]::Transparent
    return $label
}

function Set-PanelStatus {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [ValidateSet('ready', 'running', 'warning', 'error')][string]$Tone = 'ready'
    )

    if ($script:StatusBadge) {
        $script:StatusBadge.Text = $Text.ToUpperInvariant()
        switch ($Tone) {
            'ready' {
                $script:StatusBadge.BackColor = $script:Palette.GreenLight
                $script:StatusBadge.ForeColor = $script:Palette.Green
            }
            'running' {
                $script:StatusBadge.BackColor = [System.Drawing.Color]::FromArgb(232, 241, 249)
                $script:StatusBadge.ForeColor = $script:Palette.Blue
            }
            'warning' {
                $script:StatusBadge.BackColor = $script:Palette.AmberLight
                $script:StatusBadge.ForeColor = $script:Palette.Amber
            }
            'error' {
                $script:StatusBadge.BackColor = $script:Palette.RedLight
                $script:StatusBadge.ForeColor = $script:Palette.Red
            }
        }
    }
    if ($script:StatusText) {
        $script:StatusText.Text = $Text
        switch ($Tone) {
            'ready' { $script:StatusText.ForeColor = $script:Palette.Green }
            'running' { $script:StatusText.ForeColor = $script:Palette.Blue }
            'warning' { $script:StatusText.ForeColor = $script:Palette.Amber }
            'error' { $script:StatusText.ForeColor = $script:Palette.Red }
        }
    }
}

function Add-PanelLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('normal', 'success', 'warning', 'error')][string]$Tone = 'normal'
    )

    if (-not $script:OutputBox) { return }
    $color = $script:Palette.Muted
    switch ($Tone) {
        'success' { $color = $script:Palette.Green }
        'warning' { $color = $script:Palette.Amber }
        'error' { $color = $script:Palette.Red }
    }
    $entry = '[{0}]  {1}{2}' -f (Get-Date -Format 'HH:mm:ss'), $Message, [Environment]::NewLine
    $start = $script:OutputBox.TextLength
    $script:OutputBox.AppendText($entry)
    $script:OutputBox.Select($start, $entry.Length)
    $script:OutputBox.SelectionColor = $color
    $script:OutputBox.Select($script:OutputBox.TextLength, 0)
    $script:OutputBox.ScrollToCaret()

    if ($script:OutputBox.TextLength -gt 24000) {
        $trimmed = $script:OutputBox.Text.Substring($script:OutputBox.TextLength - 18000)
        $script:OutputBox.Clear()
        $script:OutputBox.AppendText($trimmed)
        $script:OutputBox.Select($script:OutputBox.TextLength, 0)
    }
}

function ConvertTo-CmdToken {
    param([Parameter(Mandatory = $true)][string]$Value)
    if ($Value -match '[\s&|<>^()]' -or $Value.Contains('"')) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

function New-NpmShellArguments {
    param(
        [Parameter(Mandatory = $true)][string[]]$NpmArguments,
        [switch]$MergeError
    )

    $tokens = @((ConvertTo-CmdToken -Value $script:NpmExecutable))
    foreach ($argument in $NpmArguments) {
        $tokens += ConvertTo-CmdToken -Value $argument
    }
    $commandLine = $tokens -join ' '
    if ($MergeError) { $commandLine += ' 2>&1' }
    return '/d /s /c "' + $commandLine + '"'
}

function Update-TaskButtons {
    $canRun = [bool]$script:NodeCommand -and [bool]$script:NpmExecutable
    $busy = $script:TaskWorker -and $script:TaskWorker.IsBusy
    foreach ($button in $script:TaskButtons) {
        $button.Enabled = $canRun -and (-not $busy)
    }
}

function Start-NpmTask {
    param(
        [Parameter(Mandatory = $true)][string[]]$NpmArguments,
        [Parameter(Mandatory = $true)][string]$TaskName
    )

    if (-not $script:NodeCommand -or -not $script:NpmExecutable) {
        Add-PanelLog 'No se encontró Node.js o npm. Instala Node.js 20.19+ o 22.12+ y vuelve a intentarlo.' 'error'
        Set-PanelStatus -Text 'Falta Node.js o npm' -Tone 'error'
        return
    }
    if ($script:TaskWorker -and $script:TaskWorker.IsBusy) {
        Add-PanelLog 'Ya hay una tarea npm en ejecución. Espera a que termine.' 'warning'
        return
    }

    $arguments = New-NpmShellArguments -NpmArguments $NpmArguments -MergeError
    $configuration = [pscustomobject]@{
        Shell = $env:ComSpec
        Arguments = $arguments
        WorkingDirectory = $script:RepoRoot
        TaskName = $TaskName
    }

    $worker = New-Object System.ComponentModel.BackgroundWorker
    $worker.WorkerReportsProgress = $true
    $worker.add_DoWork({
        param($sender, $eventArgs)
        $settings = $eventArgs.Argument
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $settings.Shell
        $startInfo.Arguments = $settings.Arguments
        $startInfo.WorkingDirectory = $settings.WorkingDirectory
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw 'No se pudo iniciar npm.' }
        $sender.ReportProgress(0, ('Iniciado: ' + $settings.TaskName))
        while ($null -ne ($line = $process.StandardOutput.ReadLine())) {
            if ($line.Length -gt 0) { $sender.ReportProgress(0, $line) }
        }
        $process.WaitForExit()
        $eventArgs.Result = [pscustomobject]@{ ExitCode = $process.ExitCode }
        $process.Dispose()
    })
    $worker.add_ProgressChanged({
        param($sender, $eventArgs)
        $message = [string]$eventArgs.UserState
        $tone = 'normal'
        if ($message -match '(?i)\b(error|failed|fatal)\b') { $tone = 'error' }
        elseif ($message -match '(?i)\b(warn|warning)\b') { $tone = 'warning' }
        Add-PanelLog -Message $message -Tone $tone
    })
    $worker.add_RunWorkerCompleted({
        param($sender, $eventArgs)
        if ($eventArgs.Error) {
            Add-PanelLog ('La tarea terminó con error: ' + $eventArgs.Error.Exception.Message) 'error'
            Set-PanelStatus -Text 'Tarea finalizada con error' -Tone 'error'
        }
        elseif ($eventArgs.Result.ExitCode -eq 0) {
            Add-PanelLog 'La tarea terminó correctamente.' 'success'
            Set-PanelStatus -Text 'Tarea completada' -Tone 'ready'
        }
        else {
            Add-PanelLog ('La tarea terminó con código ' + $eventArgs.Result.ExitCode + '.') 'error'
            Set-PanelStatus -Text 'Revisa el resultado' -Tone 'error'
        }
        $script:TaskWorker = $null
        Update-TaskButtons
        $sender.Dispose()
    })

    $script:TaskWorker = $worker
    Update-TaskButtons
    Set-PanelStatus -Text ('Ejecutando ' + $TaskName) -Tone 'running'
    Add-PanelLog ('Ejecutando ' + $TaskName + ' en segundo plano…')
    try {
        $worker.RunWorkerAsync($configuration)
    }
    catch {
        $script:TaskWorker = $null
        Update-TaskButtons
        Add-PanelLog ('No se pudo iniciar la tarea: ' + $_.Exception.Message) 'error'
        Set-PanelStatus -Text 'No se pudo iniciar' -Tone 'error'
    }
}

function Test-LocalPort {
    param([Parameter(Mandatory = $true)][int]$Port)
    $client = New-Object System.Net.Sockets.TcpClient
    $pending = $null
    try {
        $pending = $client.BeginConnect('127.0.0.1', $Port, $null, $null)
        if (-not $pending.AsyncWaitHandle.WaitOne(180)) { return $false }
        $client.EndConnect($pending)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($pending) { $pending.AsyncWaitHandle.Close() }
        $client.Close()
    }
}

function Open-LocalPage {
    param([Parameter(Mandatory = $true)][string]$Url)
    try {
        Start-Process -FilePath $Url | Out-Null
        Add-PanelLog ('Abriendo en el navegador: ' + $Url) 'success'
    }
    catch {
        Add-PanelLog ('No se pudo abrir el navegador. Dirección local: ' + $Url) 'warning'
    }
}

function Get-ServerDetails {
    param([Parameter(Mandatory = $true)][ValidateSet('https', 'preview')][string]$Mode)
    if ($Mode -eq 'https') {
        return [pscustomobject]@{
            Port = 5173
            Url = 'https://localhost:5173/'
            NpmArguments = @('run', 'dev')
            Name = 'Web segura local'
        }
    }
    return [pscustomobject]@{
        Port = 4173
        Url = 'http://localhost:4173/'
        NpmArguments = @('run', 'dev:preview', '--', '--port', '4173')
        Name = 'Vista previa HTTP'
    }
}

function Update-ServerControls {
    if ($script:StopServersButton) {
        $script:StopServersButton.Enabled = $script:ServerProcesses.Count -gt 0
    }
}

function Start-LocalServer {
    param([Parameter(Mandatory = $true)][ValidateSet('https', 'preview')][string]$Mode)

    if (-not $script:NodeCommand -or -not $script:NpmExecutable) {
        Add-PanelLog 'No se encontró Node.js o npm. Instala Node.js 20.19+ o 22.12+.' 'error'
        Set-PanelStatus -Text 'Falta Node.js o npm' -Tone 'error'
        return
    }

    $details = Get-ServerDetails -Mode $Mode
    if ($script:ServerProcesses.ContainsKey($Mode)) {
        $existing = $script:ServerProcesses[$Mode]
        try { $existing.Process.Refresh() } catch { }
        if (-not $existing.Process.HasExited) {
            if ($existing.Ready -or (Test-LocalPort -Port $details.Port)) {
                $existing.Ready = $true
                Open-LocalPage -Url $details.Url
                Set-PanelStatus -Text 'Servidor local activo' -Tone 'ready'
            }
            else {
                Add-PanelLog ($details.Name + ' todavía está iniciando.') 'warning'
            }
            return
        }
        $script:ServerProcesses.Remove($Mode)
    }

    if (Test-LocalPort -Port $details.Port) {
        Add-PanelLog ('Ya hay un servicio en el puerto ' + $details.Port + '. Se abrirá la dirección local.') 'warning'
        Open-LocalPage -Url $details.Url
        Set-PanelStatus -Text 'Servicio detectado' -Tone 'ready'
        return
    }

    try {
        $arguments = New-NpmShellArguments -NpmArguments $details.NpmArguments
        $serverProcess = Start-Process -FilePath $env:ComSpec -ArgumentList $arguments -WorkingDirectory $script:RepoRoot -WindowStyle Hidden -PassThru
        $script:ServerProcesses[$Mode] = [pscustomobject]@{
            Process = $serverProcess
            Mode = $Mode
            Port = $details.Port
            Url = $details.Url
            Name = $details.Name
            Ready = $false
            PollCount = 0
            Warned = $false
        }
        $script:ServerPollTimer.Start()
        Update-ServerControls
        Set-PanelStatus -Text ('Iniciando ' + $details.Name) -Tone 'running'
        Add-PanelLog ('Iniciando ' + $details.Name + ' en segundo plano, sin abrir una ventana de comandos…')
        if ($Mode -eq 'https') {
            Add-PanelLog 'HTTPS local usa un certificado autofirmado; el navegador puede pedir aceptar una excepción.' 'warning'
        }
    }
    catch {
        Add-PanelLog ('No se pudo iniciar el servidor: ' + $_.Exception.Message) 'error'
        Set-PanelStatus -Text 'No se pudo iniciar' -Tone 'error'
    }
}

function Stop-LocalServers {
    if ($script:ServerProcesses.Count -eq 0) {
        Add-PanelLog 'No hay servidores iniciados desde este Panel de Control.' 'warning'
        return
    }

    $taskkill = Join-Path $env:WINDIR 'System32\taskkill.exe'
    foreach ($mode in @($script:ServerProcesses.Keys)) {
        $entry = $script:ServerProcesses[$mode]
        try {
            $entry.Process.Refresh()
            if (-not $entry.Process.HasExited) {
                if (Test-Path -LiteralPath $taskkill) {
                    Start-Process -FilePath $taskkill -ArgumentList @('/PID', [string]$entry.Process.Id, '/T', '/F') -WindowStyle Hidden -Wait | Out-Null
                }
                else {
                    $entry.Process.Kill()
                }
            }
            Add-PanelLog ($entry.Name + ' detenido.') 'success'
        }
        catch {
            Add-PanelLog ('No se pudo detener ' + $entry.Name + ': ' + $_.Exception.Message) 'error'
        }
    }
    $script:ServerProcesses.Clear()
    if ($script:ServerPollTimer) { $script:ServerPollTimer.Stop() }
    Update-ServerControls
    Set-PanelStatus -Text 'Servidores detenidos' -Tone 'ready'
}

function New-ActionButton {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Action,
        [bool]$Primary = $false,
        [int]$Width = 140
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Action.Text
    $button.Width = $Width
    $button.Height = 30
    $button.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    $button.Font = [System.Drawing.Font]::new('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 1
    $button.Tag = [pscustomobject]$Action
    $button.AccessibleName = $Action.Text
    if ($Primary) {
        $button.BackColor = $script:Palette.Blue
        $button.ForeColor = $script:Palette.White
        $button.FlatAppearance.BorderColor = $script:Palette.Blue
        $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(31, 101, 158)
    }
    else {
        $button.BackColor = [System.Drawing.Color]::FromArgb(247, 250, 252)
        $button.ForeColor = [System.Drawing.Color]::FromArgb(74, 96, 114)
        $button.FlatAppearance.BorderColor = $script:Palette.Line
        $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(237, 245, 248)
    }
    $button.UseVisualStyleBackColor = $false
    $button.add_Click({
        param($sender, $eventArgs)
        $action = $sender.Tag
        switch ([string]$action.Type) {
            'server' { Start-LocalServer -Mode ([string]$action.Mode) }
            'task' { Start-NpmTask -NpmArguments ([string[]]$action.Arguments) -TaskName ([string]$action.TaskName) }
            'roadmap' { Show-ImprovementDialog }
            'stop' { Stop-LocalServers }
            'folder' { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:RepoRoot) | Out-Null }
        }
    })
    if ([string]$Action.Type -eq 'task') { $script:TaskButtons += $button }
    return $button
}

function New-ActionCard {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Description,
        [Parameter(Mandatory = $true)][hashtable]$PrimaryAction,
        [Parameter(Mandatory = $true)][hashtable]$SecondaryAction,
        [int]$PrimaryWidth = 165,
        [int]$SecondaryWidth = 120
    )

    $card = New-Object System.Windows.Forms.Panel
    $card.Dock = [System.Windows.Forms.DockStyle]::Fill
    $card.Margin = [System.Windows.Forms.Padding]::new(0, 0, 10, 8)
    $card.Padding = [System.Windows.Forms.Padding]::new(12)
    $card.BackColor = $script:Palette.White
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $titleLabel = New-UiLabel -Text $Title -FontSize 9.5 -Bold $true -ForeColor $script:Palette.Ink
    $titleLabel.Location = [System.Drawing.Point]::new(12, 8)
    $descriptionLabel = New-UiLabel -Text $Description -FontSize 8 -ForeColor $script:Palette.Muted -AutoSize $false
    $descriptionLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $descriptionLabel.Height = 22
    $descriptionLabel.Margin = [System.Windows.Forms.Padding]::new(12, 30, 12, 0)
    $descriptionLabel.AutoEllipsis = $true

    $buttonRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonRow.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $buttonRow.Height = 36
    $buttonRow.WrapContents = $false
    $buttonRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $buttonRow.Padding = [System.Windows.Forms.Padding]::new(0, 3, 0, 0)
    $buttonRow.BackColor = [System.Drawing.Color]::Transparent
    $buttonRow.Controls.Add((New-ActionButton -Action $PrimaryAction -Primary $true -Width $PrimaryWidth))
    $buttonRow.Controls.Add((New-ActionButton -Action $SecondaryAction -Width $SecondaryWidth))

    $card.Controls.Add($buttonRow)
    $card.Controls.Add($descriptionLabel)
    $card.Controls.Add($titleLabel)
    return $card
}

function Show-ImprovementDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Hoja de ruta de mejoras'
    $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dialog.MinimizeBox = $false
    $dialog.MaximizeBox = $false
    $dialog.ClientSize = [System.Drawing.Size]::new(720, 590)
    $dialog.MinimumSize = [System.Drawing.Size]::new(620, 460)
    $dialog.BackColor = $script:Palette.Canvas
    $dialog.Font = [System.Drawing.Font]::new('Segoe UI', 9)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Top
    $header.Height = 82
    $header.Padding = [System.Windows.Forms.Padding]::new(22, 16, 20, 8)
    $header.BackColor = $script:Palette.White
    $heading = New-UiLabel -Text 'Mejoras propuestas' -FontSize 16 -Bold $true -ForeColor $script:Palette.Ink
    $heading.Location = [System.Drawing.Point]::new(22, 13)
    $subheading = New-UiLabel -Text 'Ideas para evolucionar la Web gradualmente, manteniendo el procesamiento local.' -FontSize 8.5 -ForeColor $script:Palette.Muted
    $subheading.Location = [System.Drawing.Point]::new(23, 46)
    $header.Controls.Add($heading)
    $header.Controls.Add($subheading)

    $list = New-Object System.Windows.Forms.FlowLayoutPanel
    $list.Dock = [System.Windows.Forms.DockStyle]::Fill
    $list.AutoScroll = $true
    $list.WrapContents = $false
    $list.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $list.Padding = [System.Windows.Forms.Padding]::new(15, 14, 15, 12)
    $list.BackColor = $script:Palette.Canvas

    $number = 0
    foreach ($idea in $script:ImprovementIdeas) {
        $number += 1
        $item = New-Object System.Windows.Forms.Panel
        $item.Width = 660
        $item.Height = 59
        $item.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 8)
        $item.BackColor = $script:Palette.White
        $item.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

        $numberBadge = New-Object System.Windows.Forms.Label
        $numberBadge.Text = '{0:D2}' -f $number
        $numberBadge.Location = [System.Drawing.Point]::new(12, 13)
        $numberBadge.Size = [System.Drawing.Size]::new(34, 30)
        $numberBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $numberBadge.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 244)
        $numberBadge.ForeColor = $script:Palette.Teal
        $numberBadge.Font = [System.Drawing.Font]::new('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)

        $ideaLabel = New-UiLabel -Text $idea -FontSize 8.5 -ForeColor $script:Palette.Ink -AutoSize $false
        $ideaLabel.Location = [System.Drawing.Point]::new(57, 9)
        $ideaLabel.Size = [System.Drawing.Size]::new(585, 40)
        $ideaLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $item.Controls.Add($numberBadge)
        $item.Controls.Add($ideaLabel)
        $list.Controls.Add($item)
    }

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.Height = 58
    $footer.BackColor = $script:Palette.White
    $footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $footerLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $footerLayout.ColumnCount = 2
    $footerLayout.RowCount = 1
    $footerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $footerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 120)) | Out-Null
    $footerSpacer = New-Object System.Windows.Forms.Panel
    $footerSpacer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Cerrar'
    $closeButton.Dock = [System.Windows.Forms.DockStyle]::Fill
    $closeButton.Margin = [System.Windows.Forms.Padding]::new(8, 11, 14, 11)
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.FlatAppearance.BorderColor = $script:Palette.Line
    $closeButton.BackColor = $script:Palette.White
    $footerLayout.Controls.Add($footerSpacer, 0, 0)
    $footerLayout.Controls.Add($closeButton, 1, 0)
    $footer.Controls.Add($footerLayout)
    $dialog.Controls.Add($list)
    $dialog.Controls.Add($footer)
    $dialog.Controls.Add($header)
    $dialog.AcceptButton = $closeButton
    [void]$dialog.ShowDialog($script:Form)
    $dialog.Dispose()
}

function Update-EnvironmentBadges {
    if ($script:NodeBadge) {
        if ($script:NodeCommand) {
            $script:NodeBadge.Text = 'Node.js  ·  disponible'
            $script:NodeBadge.BackColor = $script:Palette.GreenLight
            $script:NodeBadge.ForeColor = $script:Palette.Green
        }
        else {
            $script:NodeBadge.Text = 'Node.js  ·  no encontrado'
            $script:NodeBadge.BackColor = $script:Palette.AmberLight
            $script:NodeBadge.ForeColor = $script:Palette.Amber
        }
    }
    if ($script:NpmBadge) {
        if ($script:NpmExecutable) {
            $script:NpmBadge.Text = 'npm  ·  disponible'
            $script:NpmBadge.BackColor = $script:Palette.GreenLight
            $script:NpmBadge.ForeColor = $script:Palette.Green
        }
        else {
            $script:NpmBadge.Text = 'npm  ·  no encontrado'
            $script:NpmBadge.BackColor = $script:Palette.AmberLight
            $script:NpmBadge.ForeColor = $script:Palette.Amber
        }
    }
    if ($script:RepoPathLabel) { $script:RepoPathLabel.Text = 'REPOSITORIO  ·  ' + $script:RepoRoot }
    Update-TaskButtons
}

function Initialize-ControlPanel {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'ONAT · Disciplina de pago | Panel de control'
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.ClientSize = [System.Drawing.Size]::new(1160, 800)
    $form.MinimumSize = [System.Drawing.Size]::new(980, 690)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.BackColor = $script:Palette.Canvas
    $form.Font = [System.Drawing.Font]::new('Segoe UI', 9)
    $form.MaximizeBox = $true
    $form.MinimizeBox = $true
    $script:Form = $form

    $outerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $outerLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $outerLayout.ColumnCount = 2
    $outerLayout.RowCount = 1
    $outerLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $outerLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $outerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 236)) | Out-Null
    $outerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $outerLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null

    $sidebar = New-Object System.Windows.Forms.Panel
    $sidebar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $sidebar.BackColor = $script:Palette.Navy

    $logoTile = New-Object System.Windows.Forms.Panel
    $logoTile.Location = [System.Drawing.Point]::new(19, 25)
    $logoTile.Size = [System.Drawing.Size]::new(45, 45)
    $logoTile.BackColor = [System.Drawing.Color]::FromArgb(37, 82, 96)
    $logoLabel = New-UiLabel -Text 'O' -FontSize 23 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(151, 215, 210))
    $logoLabel.Location = [System.Drawing.Point]::new(11, 4)
    $logoLabel.BackColor = [System.Drawing.Color]::Transparent
    $logoTile.Controls.Add($logoLabel)
    $brandLabel = New-UiLabel -Text 'ONAT' -FontSize 19 -Bold $true -ForeColor $script:Palette.White
    $brandLabel.Location = [System.Drawing.Point]::new(76, 24)
    $brandCaption = New-UiLabel -Text 'CONTROL DE PAGO' -FontSize 7 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(133, 166, 181))
    $brandCaption.Location = [System.Drawing.Point]::new(78, 49)

    $sidebarRule = New-Object System.Windows.Forms.Panel
    $sidebarRule.Location = [System.Drawing.Point]::new(18, 94)
    $sidebarRule.Size = [System.Drawing.Size]::new(200, 1)
    $sidebarRule.BackColor = [System.Drawing.Color]::FromArgb(48, 74, 91)
    $sideSection = New-UiLabel -Text 'ESPACIO DE TRABAJO' -FontSize 7 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(119, 149, 165))
    $sideSection.Location = [System.Drawing.Point]::new(22, 119)

    $activeNav = New-Object System.Windows.Forms.Panel
    $activeNav.Location = [System.Drawing.Point]::new(12, 142)
    $activeNav.Size = [System.Drawing.Size]::new(212, 44)
    $activeNav.BackColor = $script:Palette.NavyLight
    $navMarker = New-Object System.Windows.Forms.Panel
    $navMarker.Location = [System.Drawing.Point]::new(0, 11)
    $navMarker.Size = [System.Drawing.Size]::new(3, 22)
    $navMarker.BackColor = [System.Drawing.Color]::FromArgb(112, 203, 191)
    $navLabel = New-UiLabel -Text 'Panel de control' -FontSize 9 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(239, 247, 249))
    $navLabel.Location = [System.Drawing.Point]::new(18, 13)
    $activeNav.Controls.Add($navMarker)
    $activeNav.Controls.Add($navLabel)

    $accessLabel = New-UiLabel -Text 'ACCIONES' -FontSize 7 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(119, 149, 165))
    $accessLabel.Location = [System.Drawing.Point]::new(22, 212)
    $openFolderAction = @{ Type = 'folder'; Text = 'Abrir carpeta del proyecto' }
    $folderButton = New-ActionButton -Action $openFolderAction -Width 194
    $folderButton.Location = [System.Drawing.Point]::new(16, 235)
    $folderButton.Height = 35
    $folderButton.BackColor = $script:Palette.Navy
    $folderButton.ForeColor = [System.Drawing.Color]::FromArgb(190, 208, 216)
    $folderButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(62, 92, 108)
    $folderButton.FlatAppearance.MouseOverBackColor = $script:Palette.NavyLight
    $folderButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $folderButton.Padding = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
    $folderButton.Margin = [System.Windows.Forms.Padding]::Empty

    $script:StopServersButton = New-ActionButton -Action @{ Type = 'stop'; Text = 'Detener servidores' } -Width 200
    $script:StopServersButton.Location = [System.Drawing.Point]::new(18, 535)
    $script:StopServersButton.Size = [System.Drawing.Size]::new(200, 36)
    $script:StopServersButton.BackColor = [System.Drawing.Color]::FromArgb(39, 62, 77)
    $script:StopServersButton.ForeColor = [System.Drawing.Color]::FromArgb(218, 231, 236)
    $script:StopServersButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(66, 92, 106)
    $script:StopServersButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(52, 75, 89)
    $script:StopServersButton.Enabled = $false

    $privacyCard = New-Object System.Windows.Forms.Panel
    $privacyCard.Location = [System.Drawing.Point]::new(17, 591)
    $privacyCard.Size = [System.Drawing.Size]::new(202, 71)
    $privacyCard.BackColor = [System.Drawing.Color]::FromArgb(25, 57, 74)
    $script:PrivacyCard = $privacyCard
    $privacyTitle = New-UiLabel -Text 'EJECUCIÓN LOCAL' -FontSize 7 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(144, 204, 196))
    $privacyTitle.Location = [System.Drawing.Point]::new(12, 11)
    $privacyText = New-UiLabel -Text 'La Web procesa tus Excel en el navegador.' -FontSize 7.5 -ForeColor ([System.Drawing.Color]::FromArgb(179, 200, 208)) -AutoSize $false
    $privacyText.Location = [System.Drawing.Point]::new(12, 31)
    $privacyText.Size = [System.Drawing.Size]::new(178, 29)
    $privacyCard.Controls.Add($privacyTitle)
    $privacyCard.Controls.Add($privacyText)

    $sidebar.Controls.Add($logoTile)
    $sidebar.Controls.Add($brandLabel)
    $sidebar.Controls.Add($brandCaption)
    $sidebar.Controls.Add($sidebarRule)
    $sidebar.Controls.Add($sideSection)
    $sidebar.Controls.Add($activeNav)
    $sidebar.Controls.Add($accessLabel)
    $sidebar.Controls.Add($folderButton)
    $sidebar.Controls.Add($script:StopServersButton)
    $sidebar.Controls.Add($privacyCard)
    $sidebar.add_SizeChanged({
        param($sender, $eventArgs)
        $width = $sender.ClientSize.Width
        $height = $sender.ClientSize.Height
        if ($script:StopServersButton) {
            $script:StopServersButton.Location = [System.Drawing.Point]::new(18, [Math]::Max(250, $height - 140))
            $script:StopServersButton.Width = [Math]::Max(150, $width - 36)
        }
        if ($script:PrivacyCard) {
            $script:PrivacyCard.Location = [System.Drawing.Point]::new(17, [Math]::Max(320, $height - 88))
            $script:PrivacyCard.Width = [Math]::Max(150, $width - 34)
        }
    })

    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainPanel.BackColor = $script:Palette.Canvas
    $mainPanel.Padding = [System.Windows.Forms.Padding]::new(25, 20, 25, 18)

    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 6
    $mainLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $mainLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $mainLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 62)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 134)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 91)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 222)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 28)) | Out-Null

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Fill
    $header.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 8)
    $headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $headerLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $headerLayout.ColumnCount = 2
    $headerLayout.RowCount = 1
    $headerLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $headerLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $headerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $headerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 155)) | Out-Null
    $headerContent = New-Object System.Windows.Forms.Panel
    $headerContent.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pageTitle = New-UiLabel -Text 'Centro de control' -FontSize 20 -Bold $true -ForeColor $script:Palette.Ink
    $pageTitle.Location = [System.Drawing.Point]::new(0, 0)
    $pageSubtitle = New-UiLabel -Text 'Herramientas locales para desarrollar, comprobar y abrir la aplicación.' -FontSize 8.5 -ForeColor $script:Palette.Muted -AutoSize $false
    $pageSubtitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $pageSubtitle.Height = 22
    $pageSubtitle.Margin = [System.Windows.Forms.Padding]::new(1, 33, 0, 0)
    $pageSubtitle.AutoEllipsis = $true
    $headerContent.Controls.Add($pageSubtitle)
    $headerContent.Controls.Add($pageTitle)
    $statusSlot = New-Object System.Windows.Forms.Panel
    $statusSlot.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:StatusBadge = New-Object System.Windows.Forms.Label
    $script:StatusBadge.Text = 'LISTO'
    $script:StatusBadge.Size = [System.Drawing.Size]::new(145, 25)
    $script:StatusBadge.Location = [System.Drawing.Point]::new(0, 9)
    $script:StatusBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:StatusBadge.Font = [System.Drawing.Font]::new('Segoe UI', 7, [System.Drawing.FontStyle]::Bold)
    $script:StatusBadge.BackColor = $script:Palette.GreenLight
    $script:StatusBadge.ForeColor = $script:Palette.Green
    $script:StatusBadge.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $statusSlot.Controls.Add($script:StatusBadge)
    $headerLayout.Controls.Add($headerContent, 0, 0)
    $headerLayout.Controls.Add($statusSlot, 1, 0)
    $header.Controls.Add($headerLayout)

    $hero = New-Object System.Windows.Forms.Panel
    $hero.Dock = [System.Windows.Forms.DockStyle]::Fill
    $hero.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    $hero.Padding = [System.Windows.Forms.Padding]::new(20, 15, 20, 12)
    $hero.BackColor = $script:Palette.Navy
    $heroLabel = New-UiLabel -Text 'ONAT  ·  ENTORNO DE DESARROLLO' -FontSize 7 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(151, 194, 203))
    $heroLabel.Location = [System.Drawing.Point]::new(20, 13)
    $heroTitle = New-UiLabel -Text 'Todo listo para poner el proyecto en marcha.' -FontSize 15 -Bold $true -ForeColor $script:Palette.White
    $heroTitle.Location = [System.Drawing.Point]::new(20, 34)
    $heroText = New-UiLabel -Text 'Inicia el sitio seguro, valida los cambios y consulta el resultado sin salir de este panel.' -FontSize 8.5 -ForeColor ([System.Drawing.Color]::FromArgb(188, 211, 218)) -AutoSize $false
    $heroText.Dock = [System.Windows.Forms.DockStyle]::Top
    $heroText.Height = 25
    $heroText.Margin = [System.Windows.Forms.Padding]::new(21, 66, 20, 0)
    $heroText.AutoEllipsis = $true
    $heroDecoration = New-Object System.Windows.Forms.Panel
    $heroDecoration.Size = [System.Drawing.Size]::new(7, 7)
    $heroDecoration.Location = [System.Drawing.Point]::new(20, 105)
    $heroDecoration.BackColor = [System.Drawing.Color]::FromArgb(105, 211, 175)
    $heroCaption = New-UiLabel -Text 'LOCAL  ·  HTTPS  ·  SIN MODIFICAR LOS DATOS' -FontSize 6.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(151, 194, 203))
    $heroCaption.Location = [System.Drawing.Point]::new(35, 102)
    $hero.Controls.Add($heroLabel)
    $hero.Controls.Add($heroTitle)
    $hero.Controls.Add($heroText)
    $hero.Controls.Add($heroDecoration)
    $hero.Controls.Add($heroCaption)

    $environment = New-Object System.Windows.Forms.Panel
    $environment.Dock = [System.Windows.Forms.DockStyle]::Fill
    $environment.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    $environment.BackColor = $script:Palette.White
    $environment.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $environmentTitle = New-UiLabel -Text 'ESTADO DEL ENTORNO' -FontSize 7 -Bold $true -ForeColor $script:Palette.Muted
    $environmentTitle.Location = [System.Drawing.Point]::new(13, 9)
    $script:NodeBadge = New-Object System.Windows.Forms.Label
    $script:NodeBadge.Location = [System.Drawing.Point]::new(172, 7)
    $script:NodeBadge.Size = [System.Drawing.Size]::new(153, 24)
    $script:NodeBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:NodeBadge.Font = [System.Drawing.Font]::new('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)
    $script:NpmBadge = New-Object System.Windows.Forms.Label
    $script:NpmBadge.Location = [System.Drawing.Point]::new(333, 7)
    $script:NpmBadge.Size = [System.Drawing.Size]::new(142, 24)
    $script:NpmBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:NpmBadge.Font = [System.Drawing.Font]::new('Segoe UI', 7.5, [System.Drawing.FontStyle]::Bold)
    $script:RepoPathLabel = New-UiLabel -Text '' -FontSize 7 -Bold $false -ForeColor ([System.Drawing.Color]::FromArgb(129, 145, 157)) -AutoSize $false
    $script:RepoPathLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $script:RepoPathLabel.Height = 22
    $script:RepoPathLabel.Margin = [System.Windows.Forms.Padding]::new(13, 42, 13, 0)
    $script:RepoPathLabel.AutoEllipsis = $true
    $environment.Controls.Add($environmentTitle)
    $environment.Controls.Add($script:NodeBadge)
    $environment.Controls.Add($script:NpmBadge)
    $environment.Controls.Add($script:RepoPathLabel)

    $actions = New-Object System.Windows.Forms.TableLayoutPanel
    $actions.Dock = [System.Windows.Forms.DockStyle]::Fill
    $actions.ColumnCount = 2
    $actions.RowCount = 2
    $actions.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 8)
    $actions.Padding = [System.Windows.Forms.Padding]::Empty
    $actions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 50)) | Out-Null
    $actions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 50)) | Out-Null
    $actions.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 50)) | Out-Null
    $actions.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 50)) | Out-Null

    $cardWeb = New-ActionCard -Title 'Abrir la aplicación' -Description 'Servidor seguro local y vista previa para pruebas rápidas.' `
        -PrimaryAction @{ Type = 'server'; Mode = 'https'; Text = 'Web HTTPS'; TaskName = 'Web segura' } `
        -SecondaryAction @{ Type = 'server'; Mode = 'preview'; Text = 'Preview HTTP'; TaskName = 'Vista previa' } `
        -PrimaryWidth 112 -SecondaryWidth 112
    $cardChecks = New-ActionCard -Title 'Calidad del proyecto' -Description 'Ejecuta la batería de pruebas y la comprobación integral.' `
        -PrimaryAction @{ Type = 'task'; Arguments = @('run', 'check'); Text = 'Comprobación completa'; TaskName = 'Comprobación completa' } `
        -SecondaryAction @{ Type = 'task'; Arguments = @('test'); Text = 'Solo pruebas'; TaskName = 'Pruebas' } `
        -PrimaryWidth 164 -SecondaryWidth 104
    $cardBuild = New-ActionCard -Title 'Compilación' -Description 'Genera la versión de producción o instala dependencias limpias.' `
        -PrimaryAction @{ Type = 'task'; Arguments = @('run', 'build'); Text = 'Compilar producción'; TaskName = 'Compilación de producción' } `
        -SecondaryAction @{ Type = 'task'; Arguments = @('ci'); Text = 'Instalar npm'; TaskName = 'Instalación de dependencias' } `
        -PrimaryWidth 148 -SecondaryWidth 110
    $cardMaintain = New-ActionCard -Title 'Mantenimiento' -Description 'Consulta las dependencias y la hoja de ruta de mejoras.' `
        -PrimaryAction @{ Type = 'task'; Arguments = @('audit'); Text = 'Auditoría npm'; TaskName = 'Auditoría de dependencias' } `
        -SecondaryAction @{ Type = 'roadmap'; Text = 'Ideas futuras' } `
        -PrimaryWidth 114 -SecondaryWidth 112

    $actions.Controls.Add($cardWeb, 0, 0)
    $actions.Controls.Add($cardChecks, 1, 0)
    $actions.Controls.Add($cardBuild, 0, 1)
    $actions.Controls.Add($cardMaintain, 1, 1)

    $logCard = New-Object System.Windows.Forms.Panel
    $logCard.Dock = [System.Windows.Forms.DockStyle]::Fill
    $logCard.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 8)
    $logCard.Padding = [System.Windows.Forms.Padding]::new(1)
    $logCard.BackColor = $script:Palette.Line
    $logLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $logLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $logLayout.RowCount = 2
    $logLayout.ColumnCount = 1
    $logLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $logLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $logLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 34)) | Out-Null
    $logLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $logHeading = New-Object System.Windows.Forms.Panel
    $logHeading.Dock = [System.Windows.Forms.DockStyle]::Fill
    $logHeading.BackColor = $script:Palette.White
    $logTitle = New-UiLabel -Text 'Actividad reciente' -FontSize 9 -Bold $true -ForeColor $script:Palette.Ink
    $logTitle.Location = [System.Drawing.Point]::new(12, 9)
    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = 'Limpiar'
    $clearButton.Size = [System.Drawing.Size]::new(62, 25)
    $clearButton.Dock = [System.Windows.Forms.DockStyle]::Right
    $clearButton.Margin = [System.Windows.Forms.Padding]::new(0, 3, 8, 3)
    $clearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $clearButton.FlatAppearance.BorderSize = 0
    $clearButton.BackColor = $script:Palette.White
    $clearButton.ForeColor = $script:Palette.Muted
    $clearButton.Font = [System.Drawing.Font]::new('Segoe UI', 7.5)
    $clearButton.add_Click({ $script:OutputBox.Clear() })
    $logHeading.Controls.Add($logTitle)
    $logHeading.Controls.Add($clearButton)
    $script:OutputBox = New-Object System.Windows.Forms.RichTextBox
    $script:OutputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:OutputBox.ReadOnly = $true
    $script:OutputBox.DetectUrls = $false
    $script:OutputBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:OutputBox.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 252)
    $script:OutputBox.ForeColor = [System.Drawing.Color]::FromArgb(78, 96, 111)
    $script:OutputBox.Font = [System.Drawing.Font]::new('Consolas', 8.5)
    $script:OutputBox.HideSelection = $false
    $script:OutputBox.MaxLength = 30000
    $script:OutputBox.Margin = [System.Windows.Forms.Padding]::new(10)
    $logLayout.Controls.Add($logHeading, 0, 0)
    $logLayout.Controls.Add($script:OutputBox, 0, 1)
    $logCard.Controls.Add($logLayout)

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $footer.Margin = [System.Windows.Forms.Padding]::Empty
    $footerLabel = New-UiLabel -Text 'ONAT  ·  Herramientas de desarrollo local' -FontSize 7 -ForeColor ([System.Drawing.Color]::FromArgb(138, 152, 164))
    $footerLabel.Location = [System.Drawing.Point]::new(0, 6)
    $script:StatusText = New-UiLabel -Text 'Listo' -FontSize 7 -Bold $true -ForeColor $script:Palette.Green -AutoSize $false
    $script:StatusText.Dock = [System.Windows.Forms.DockStyle]::Right
    $script:StatusText.Width = 180
    $script:StatusText.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $footer.Controls.Add($footerLabel)
    $footer.Controls.Add($script:StatusText)

    $mainLayout.Controls.Add($header, 0, 0)
    $mainLayout.Controls.Add($hero, 0, 1)
    $mainLayout.Controls.Add($environment, 0, 2)
    $mainLayout.Controls.Add($actions, 0, 3)
    $mainLayout.Controls.Add($logCard, 0, 4)
    $mainLayout.Controls.Add($footer, 0, 5)
    $mainPanel.Controls.Add($mainLayout)
    $outerLayout.Controls.Add($sidebar, 0, 0)
    $outerLayout.Controls.Add($mainPanel, 1, 0)
    $form.Controls.Add($outerLayout)

    $script:ServerPollTimer = New-Object System.Windows.Forms.Timer
    $script:ServerPollTimer.Interval = 650
    $script:ServerPollTimer.add_Tick({
        foreach ($mode in @($script:ServerProcesses.Keys)) {
            $entry = $script:ServerProcesses[$mode]
            try { $entry.Process.Refresh() } catch { }
            if ($entry.Process.HasExited) {
                Add-PanelLog ($entry.Name + ' se cerró antes de quedar disponible.') 'error'
                $script:ServerProcesses.Remove($mode)
                continue
            }
            $entry.PollCount += 1
            if (Test-LocalPort -Port $entry.Port) {
                if (-not $entry.Ready) {
                    $entry.Ready = $true
                    Add-PanelLog ($entry.Name + ' listo en ' + $entry.Url) 'success'
                    Set-PanelStatus -Text 'Servidor local activo' -Tone 'ready'
                    Open-LocalPage -Url $entry.Url
                }
            }
            elseif (($entry.PollCount -ge 25) -and (-not $entry.Warned)) {
                $entry.Warned = $true
                Add-PanelLog ('El servidor aún no responde en el puerto ' + $entry.Port + '. Revisa que npm termine de iniciar.') 'warning'
                Set-PanelStatus -Text 'Esperando al servidor' -Tone 'warning'
            }
        }
        if ($script:ServerProcesses.Count -eq 0) { $script:ServerPollTimer.Stop() }
        Update-ServerControls
    })

    if (-not $script:NodeCommand -or -not $script:NpmExecutable) {
        $heroTitle.Text = 'Instala Node.js para habilitar las herramientas.'
        $heroText.Text = 'El panel visual está listo; instala Node.js 20.19+ o 22.12+ y npm para ejecutar las tareas del proyecto.'
    }

    Update-EnvironmentBadges
    Add-PanelLog 'Panel visual listo. Los comandos y sus resultados se muestran aquí.' 'success'
    Add-PanelLog ('Repositorio local: ' + $script:RepoRoot)
    if (-not $script:NodeCommand -or -not $script:NpmExecutable) {
        Add-PanelLog 'Instala Node.js 20.19+ o 22.12+ para habilitar las acciones del proyecto.' 'warning'
        Set-PanelStatus -Text 'Configura Node.js y npm' -Tone 'warning'
    }
    else {
        Add-PanelLog 'Node.js y npm disponibles.' 'success'
        Set-PanelStatus -Text 'Entorno listo' -Tone 'ready'
    }

    $form.add_FormClosing({
        param($sender, $eventArgs)
        if ($script:TaskWorker -and $script:TaskWorker.IsBusy) {
            [System.Windows.Forms.MessageBox]::Show(
                'Hay una tarea npm en ejecución. Espera a que termine antes de cerrar el panel.',
                'Tarea en curso',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            $eventArgs.Cancel = $true
            return
        }
        if ($script:ServerProcesses.Count -gt 0) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                'Hay servidores locales iniciados desde este panel. ¿Quieres detenerlos antes de salir?' + [Environment]::NewLine + [Environment]::NewLine + 'Selecciona No para dejarlos ejecutándose.',
                'Servidores locales activos',
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) {
                $eventArgs.Cancel = $true
                return
            }
            if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { Stop-LocalServers }
        }
        if ($script:ServerPollTimer) {
            $script:ServerPollTimer.Stop()
            $script:ServerPollTimer.Dispose()
        }
    })

    [System.Windows.Forms.Application]::Run($form)
    $form.Dispose()
}

Initialize-ControlPanel
