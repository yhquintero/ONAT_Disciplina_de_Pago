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
$script:PathToolTip = $null
$script:StopServersButton = $null
$script:PrivacyCard = $null
$script:PrivacyTextLabel = $null
$script:TaskButtons = @()
$script:TaskWorker = $null
$script:ServerProcesses = @{}
$script:ServerPollTimer = $null

function New-UiLabel {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
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
    $label.Margin = [System.Windows.Forms.Padding]::Empty
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.AutoEllipsis = $false
    $label.UseMnemonic = $false
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
    $button.AutoSize = $true
    $button.MinimumSize = [System.Drawing.Size]::new($Width, 38)
    $button.Padding = [System.Windows.Forms.Padding]::new(10, 5, 10, 5)
    $button.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 5)
    $button.Font = [System.Drawing.Font]::new('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
    $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $button.UseMnemonic = $false
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
    $card.AutoSize = $true
    $card.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $card.MinimumSize = [System.Drawing.Size]::new(280, 0)
    $card.Margin = [System.Windows.Forms.Padding]::new(0, 0, 10, 10)
    $card.Padding = [System.Windows.Forms.Padding]::new(12)
    $card.BackColor = $script:Palette.White
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $cardLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $cardLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cardLayout.AutoSize = $true
    $cardLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $cardLayout.ColumnCount = 1
    $cardLayout.RowCount = 3
    $cardLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $cardLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $cardLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $cardLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $cardLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $cardLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    $titleLabel = New-UiLabel -Text $Title -FontSize 10.5 -Bold $true -ForeColor $script:Palette.Ink
    $titleLabel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)

    $descriptionLabel = New-UiLabel -Text $Description -FontSize 9.5 -ForeColor $script:Palette.Muted
    $descriptionLabel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 8)
    $descriptionLabel.MaximumSize = [System.Drawing.Size]::new(520, 0)

    $buttonRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $buttonRow.AutoSize = $true
    $buttonRow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $buttonRow.MinimumSize = [System.Drawing.Size]::new(0, 40)
    $buttonRow.WrapContents = $true
    $buttonRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $buttonRow.Padding = [System.Windows.Forms.Padding]::Empty
    $buttonRow.Margin = [System.Windows.Forms.Padding]::Empty
    $buttonRow.BackColor = [System.Drawing.Color]::Transparent
    $buttonRow.Controls.Add((New-ActionButton -Action $PrimaryAction -Primary $true -Width $PrimaryWidth))
    $buttonRow.Controls.Add((New-ActionButton -Action $SecondaryAction -Width $SecondaryWidth))

    $cardLayout.Controls.Add($titleLabel, 0, 0)
    $cardLayout.Controls.Add($descriptionLabel, 0, 1)
    $cardLayout.Controls.Add($buttonRow, 0, 2)
    $card.Controls.Add($cardLayout)
    $card.Tag = [pscustomobject]@{ Description = $descriptionLabel }
    $card.add_SizeChanged({
        param($sender, $eventArgs)
        if ($sender.Tag -and $sender.Tag.Description -and $sender.ClientSize.Width -gt 0) {
            $availableWidth = [Math]::Max(180, $sender.ClientSize.Width - $sender.Padding.Horizontal - 6)
            $newMaximum = [System.Drawing.Size]::new($availableWidth, 0)
            if ($sender.Tag.Description.MaximumSize -ne $newMaximum) {
                $sender.Tag.Description.MaximumSize = $newMaximum
            }
        }
    })
    return $card
}

function Show-ImprovementDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Hoja de ruta de mejoras'
    $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dialog.MinimizeBox = $false
    $dialog.MaximizeBox = $false
    $dialog.ClientSize = [System.Drawing.Size]::new(820, 680)
    $dialog.MinimumSize = [System.Drawing.Size]::new(680, 520)
    $dialog.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $dialog.BackColor = $script:Palette.Canvas
    $dialog.Font = [System.Drawing.Font]::new('Segoe UI', 9.5)

    $dialogLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $dialogLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dialogLayout.ColumnCount = 1
    $dialogLayout.RowCount = 3
    $dialogLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $dialogLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $dialogLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $dialogLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $dialogLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $dialogLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    $header = New-Object System.Windows.Forms.TableLayoutPanel
    $header.Dock = [System.Windows.Forms.DockStyle]::Fill
    $header.AutoSize = $true
    $header.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $header.ColumnCount = 1
    $header.RowCount = 2
    $header.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 1)
    $header.Padding = [System.Windows.Forms.Padding]::new(22, 16, 20, 14)
    $header.BackColor = $script:Palette.White
    $header.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $header.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $header.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $heading = New-UiLabel -Text 'Mejoras propuestas' -FontSize 17 -Bold $true -ForeColor $script:Palette.Ink
    $heading.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 4)
    $subheading = New-UiLabel -Text 'Ideas para evolucionar la Web gradualmente, manteniendo el procesamiento local.' -FontSize 9.5 -ForeColor $script:Palette.Muted
    $header.Controls.Add($heading, 0, 0)
    $header.Controls.Add($subheading, 0, 1)

    $list = New-Object System.Windows.Forms.FlowLayoutPanel
    $list.Dock = [System.Windows.Forms.DockStyle]::Fill
    $list.AutoScroll = $true
    $list.WrapContents = $false
    $list.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $list.Padding = [System.Windows.Forms.Padding]::new(16, 14, 16, 12)
    $list.Margin = [System.Windows.Forms.Padding]::Empty
    $list.BackColor = $script:Palette.Canvas
    $list.add_SizeChanged({
        param($sender, $eventArgs)
        $availableWidth = [Math]::Max(480, $sender.ClientSize.Width - $sender.Padding.Horizontal - 26)
        foreach ($itemPanel in $sender.Controls) {
            $itemPanel.Width = $availableWidth
            $ideaText = $itemPanel.Tag
            if ($ideaText -and $itemPanel.ClientSize.Width -gt 0) {
                $textWidth = [Math]::Max(300, $itemPanel.ClientSize.Width - 82)
                $newMaximum = [System.Drawing.Size]::new($textWidth, 0)
                if ($ideaText.MaximumSize -ne $newMaximum) {
                    $ideaText.MaximumSize = $newMaximum
                }
            }
        }
    })

    $number = 0
    foreach ($idea in $script:ImprovementIdeas) {
        $number += 1
        $item = New-Object System.Windows.Forms.TableLayoutPanel
        $item.AutoSize = $true
        $item.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowOnly
        $item.ColumnCount = 2
        $item.RowCount = 1
        $item.Width = 660
        $item.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 9)
        $item.Padding = [System.Windows.Forms.Padding]::new(10, 8, 10, 8)
        $item.BackColor = $script:Palette.White
        $item.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $item.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 44)) | Out-Null
        $item.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
        $item.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

        $numberBadge = New-Object System.Windows.Forms.Label
        $numberBadge.Text = '{0:D2}' -f $number
        $numberBadge.Size = [System.Drawing.Size]::new(36, 36)
        $numberBadge.MinimumSize = [System.Drawing.Size]::new(36, 36)
        $numberBadge.Margin = [System.Windows.Forms.Padding]::Empty
        $numberBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $numberBadge.Anchor = [System.Windows.Forms.AnchorStyles]::None
        $numberBadge.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 244)
        $numberBadge.ForeColor = $script:Palette.Teal
        $numberBadge.Font = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

        $ideaLabel = New-UiLabel -Text $idea -FontSize 9.5 -ForeColor $script:Palette.Ink
        $ideaLabel.Name = 'IdeaText'
        $ideaLabel.MaximumSize = [System.Drawing.Size]::new(580, 0)
        $ideaLabel.Margin = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)
        $item.Controls.Add($numberBadge, 0, 0)
        $item.Controls.Add($ideaLabel, 1, 0)
        $item.Tag = $ideaLabel
        $list.Controls.Add($item)
    }

    $footer = New-Object System.Windows.Forms.FlowLayoutPanel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $footer.AutoSize = $true
    $footer.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $footer.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $footer.WrapContents = $false
    $footer.Padding = [System.Windows.Forms.Padding]::new(12, 10, 14, 10)
    $footer.Margin = [System.Windows.Forms.Padding]::Empty
    $footer.BackColor = $script:Palette.White
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = 'Cerrar'
    $closeButton.AutoSize = $true
    $closeButton.MinimumSize = [System.Drawing.Size]::new(96, 36)
    $closeButton.Padding = [System.Windows.Forms.Padding]::new(12, 6, 12, 6)
    $closeButton.Margin = [System.Windows.Forms.Padding]::Empty
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.FlatAppearance.BorderColor = $script:Palette.Line
    $closeButton.BackColor = $script:Palette.White
    $closeButton.Font = [System.Drawing.Font]::new('Segoe UI', 9.5)
    $footer.Controls.Add($closeButton)

    $dialogLayout.Controls.Add($header, 0, 0)
    $dialogLayout.Controls.Add($list, 0, 1)
    $dialogLayout.Controls.Add($footer, 0, 2)
    $dialog.Controls.Add($dialogLayout)
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
    $workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $initialWidth = [int][Math]::Min(1440, [Math]::Max(900, $workArea.Width - 48))
    $initialHeight = [int][Math]::Min(920, [Math]::Max(620, $workArea.Height - 60))
    $form.ClientSize = [System.Drawing.Size]::new($initialWidth, $initialHeight)
    $form.MinimumSize = [System.Drawing.Size]::new(900, 560)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.BackColor = $script:Palette.Canvas
    $form.Font = [System.Drawing.Font]::new('Segoe UI', 9.5)
    $form.MaximizeBox = $true
    $form.MinimizeBox = $true
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
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

    # Sidebar: filas AutoSize + un separador elástico para que el pie no invada las acciones.
    $sidebar = New-Object System.Windows.Forms.TableLayoutPanel
    $sidebar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $sidebar.ColumnCount = 1
    $sidebar.RowCount = 9
    $sidebar.Padding = [System.Windows.Forms.Padding]::new(16, 18, 16, 14)
    $sidebar.Margin = [System.Windows.Forms.Padding]::Empty
    $sidebar.BackColor = $script:Palette.Navy
    $sidebar.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 1)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $sidebar.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    $brandRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $brandRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $brandRow.AutoSize = $true
    $brandRow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $brandRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $brandRow.WrapContents = $false
    $brandRow.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 16)
    $brandRow.Padding = [System.Windows.Forms.Padding]::Empty
    $brandRow.BackColor = [System.Drawing.Color]::Transparent

    $logoTile = New-Object System.Windows.Forms.Panel
    $logoTile.Size = [System.Drawing.Size]::new(46, 46)
    $logoTile.MinimumSize = [System.Drawing.Size]::new(46, 46)
    $logoTile.Margin = [System.Windows.Forms.Padding]::Empty
    $logoTile.BackColor = [System.Drawing.Color]::FromArgb(37, 82, 96)
    $logoLabel = New-UiLabel -Text 'O' -FontSize 23 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(151, 215, 210)) -AutoSize $false
    $logoLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $logoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $logoTile.Controls.Add($logoLabel)

    $brandStack = New-Object System.Windows.Forms.FlowLayoutPanel
    $brandStack.AutoSize = $true
    $brandStack.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $brandStack.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $brandStack.WrapContents = $false
    $brandStack.Margin = [System.Windows.Forms.Padding]::new(12, 0, 0, 0)
    $brandStack.Padding = [System.Windows.Forms.Padding]::Empty
    $brandStack.BackColor = [System.Drawing.Color]::Transparent
    $brandLabel = New-UiLabel -Text 'ONAT' -FontSize 19 -Bold $true -ForeColor $script:Palette.White
    $brandCaption = New-UiLabel -Text 'CONTROL DE PAGO' -FontSize 8.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(133, 166, 181))
    $brandCaption.Margin = [System.Windows.Forms.Padding]::new(0, 2, 0, 0)
    $brandStack.Controls.Add($brandLabel)
    $brandStack.Controls.Add($brandCaption)
    $brandRow.Controls.Add($logoTile)
    $brandRow.Controls.Add($brandStack)

    $sidebarRule = New-Object System.Windows.Forms.Panel
    $sidebarRule.Dock = [System.Windows.Forms.DockStyle]::Fill
    $sidebarRule.Margin = [System.Windows.Forms.Padding]::Empty
    $sidebarRule.BackColor = [System.Drawing.Color]::FromArgb(48, 74, 91)

    $sideSection = New-UiLabel -Text 'ESPACIO DE TRABAJO' -FontSize 8.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(144, 166, 178))
    $sideSection.Margin = [System.Windows.Forms.Padding]::new(4, 16, 0, 10)

    $activeNav = New-Object System.Windows.Forms.TableLayoutPanel
    $activeNav.Dock = [System.Windows.Forms.DockStyle]::Fill
    $activeNav.AutoSize = $true
    $activeNav.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $activeNav.MinimumSize = [System.Drawing.Size]::new(0, 44)
    $activeNav.ColumnCount = 2
    $activeNav.RowCount = 1
    $activeNav.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 18)
    $activeNav.Padding = [System.Windows.Forms.Padding]::new(0, 8, 8, 8)
    $activeNav.BackColor = $script:Palette.NavyLight
    $activeNav.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 3)) | Out-Null
    $activeNav.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $activeNav.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $navMarker = New-Object System.Windows.Forms.Panel
    $navMarker.Dock = [System.Windows.Forms.DockStyle]::Fill
    $navMarker.Margin = [System.Windows.Forms.Padding]::Empty
    $navMarker.BackColor = [System.Drawing.Color]::FromArgb(112, 203, 191)
    $navLabel = New-UiLabel -Text 'Panel de control' -FontSize 9.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(239, 247, 249))
    $navLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $navLabel.Margin = [System.Windows.Forms.Padding]::new(12, 0, 0, 0)
    $activeNav.Controls.Add($navMarker, 0, 0)
    $activeNav.Controls.Add($navLabel, 1, 0)

    $accessLabel = New-UiLabel -Text 'ACCIONES' -FontSize 8.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(144, 166, 178))
    $accessLabel.Margin = [System.Windows.Forms.Padding]::new(4, 0, 0, 8)
    $openFolderAction = @{ Type = 'folder'; Text = 'Abrir carpeta del proyecto' }
    $folderButton = New-ActionButton -Action $openFolderAction -Width 194
    $folderButton.AutoSize = $false
    $folderButton.Dock = [System.Windows.Forms.DockStyle]::Fill
    $folderButton.MinimumSize = [System.Drawing.Size]::new(0, 42)
    $folderButton.Height = 42
    $folderButton.Margin = [System.Windows.Forms.Padding]::Empty
    $folderButton.BackColor = $script:Palette.Navy
    $folderButton.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 231)
    $folderButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(62, 92, 108)
    $folderButton.FlatAppearance.MouseOverBackColor = $script:Palette.NavyLight
    $folderButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $folderButton.Padding = [System.Windows.Forms.Padding]::new(10, 6, 8, 6)

    $script:StopServersButton = New-ActionButton -Action @{ Type = 'stop'; Text = 'Detener servidores' } -Width 200
    $script:StopServersButton.AutoSize = $false
    $script:StopServersButton.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:StopServersButton.MinimumSize = [System.Drawing.Size]::new(0, 42)
    $script:StopServersButton.Height = 42
    $script:StopServersButton.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 12)
    $script:StopServersButton.BackColor = [System.Drawing.Color]::FromArgb(39, 62, 77)
    $script:StopServersButton.ForeColor = [System.Drawing.Color]::FromArgb(218, 231, 236)
    $script:StopServersButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(66, 92, 106)
    $script:StopServersButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(52, 75, 89)
    $script:StopServersButton.Enabled = $false

    $privacyCard = New-Object System.Windows.Forms.TableLayoutPanel
    $privacyCard.Dock = [System.Windows.Forms.DockStyle]::Fill
    $privacyCard.AutoSize = $true
    $privacyCard.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $privacyCard.ColumnCount = 1
    $privacyCard.RowCount = 2
    $privacyCard.Padding = [System.Windows.Forms.Padding]::new(12)
    $privacyCard.Margin = [System.Windows.Forms.Padding]::Empty
    $privacyCard.BackColor = [System.Drawing.Color]::FromArgb(25, 57, 74)
    $privacyCard.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $privacyCard.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $privacyCard.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $script:PrivacyCard = $privacyCard
    $privacyTitle = New-UiLabel -Text 'EJECUCIÓN LOCAL' -FontSize 8.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(144, 204, 196))
    $privacyTitle.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 7)
    $script:PrivacyTextLabel = New-UiLabel -Text 'La Web procesa tus Excel en el navegador.' -FontSize 9 -ForeColor ([System.Drawing.Color]::FromArgb(196, 211, 218))
    $script:PrivacyTextLabel.MaximumSize = [System.Drawing.Size]::new(180, 0)
    $privacyCard.Controls.Add($privacyTitle, 0, 0)
    $privacyCard.Controls.Add($script:PrivacyTextLabel, 0, 1)

    $sidebar.Controls.Add($brandRow, 0, 0)
    $sidebar.Controls.Add($sidebarRule, 0, 1)
    $sidebar.Controls.Add($sideSection, 0, 2)
    $sidebar.Controls.Add($activeNav, 0, 3)
    $sidebar.Controls.Add($accessLabel, 0, 4)
    $sidebar.Controls.Add($folderButton, 0, 5)
    $sidebar.Controls.Add($script:StopServersButton, 0, 7)
    $sidebar.Controls.Add($privacyCard, 0, 8)

    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainPanel.BackColor = $script:Palette.Canvas
    $mainPanel.Padding = [System.Windows.Forms.Padding]::new(24, 18, 24, 14)

    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 6
    $mainLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $mainLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $mainLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    # Encabezado: columnas medidas por contenido; no se mezclan ubicaciones absolutas con Dock.
    $header = New-Object System.Windows.Forms.TableLayoutPanel
    $header.Dock = [System.Windows.Forms.DockStyle]::Fill
    $header.AutoSize = $true
    $header.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $header.ColumnCount = 2
    $header.RowCount = 1
    $header.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    $header.Padding = [System.Windows.Forms.Padding]::Empty
    $header.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $header.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $header.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    $headerContent = New-Object System.Windows.Forms.TableLayoutPanel
    $headerContent.Dock = [System.Windows.Forms.DockStyle]::Fill
    $headerContent.AutoSize = $true
    $headerContent.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $headerContent.ColumnCount = 1
    $headerContent.RowCount = 2
    $headerContent.Margin = [System.Windows.Forms.Padding]::Empty
    $headerContent.Padding = [System.Windows.Forms.Padding]::Empty
    $headerContent.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $headerContent.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $headerContent.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $pageTitle = New-UiLabel -Text 'Centro de control' -FontSize 20 -Bold $true -ForeColor $script:Palette.Ink
    $pageTitle.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 2)
    $pageSubtitle = New-UiLabel -Text 'Herramientas locales para desarrollar, comprobar y abrir la aplicación.' -FontSize 10 -ForeColor $script:Palette.Muted
    $headerContent.Controls.Add($pageTitle, 0, 0)
    $headerContent.Controls.Add($pageSubtitle, 0, 1)

    $script:StatusBadge = New-UiLabel -Text 'LISTO' -FontSize 8.5 -Bold $true -ForeColor $script:Palette.Green
    $script:StatusBadge.AutoSize = $true
    $script:StatusBadge.Padding = [System.Windows.Forms.Padding]::new(12, 8, 12, 8)
    $script:StatusBadge.Margin = [System.Windows.Forms.Padding]::new(12, 0, 0, 0)
    $script:StatusBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:StatusBadge.BackColor = $script:Palette.GreenLight
    $script:StatusBadge.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:StatusBadge.MaximumSize = [System.Drawing.Size]::new(260, 0)
    $header.Controls.Add($headerContent, 0, 0)
    $header.Controls.Add($script:StatusBadge, 1, 0)

    # Hero de altura intrínseca. Cada renglón ocupa una fila propia y puede crecer con el texto.
    $hero = New-Object System.Windows.Forms.TableLayoutPanel
    $hero.Dock = [System.Windows.Forms.DockStyle]::Fill
    $hero.AutoSize = $true
    $hero.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $hero.ColumnCount = 1
    $hero.RowCount = 4
    $hero.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    $hero.Padding = [System.Windows.Forms.Padding]::new(18, 15, 18, 14)
    $hero.BackColor = $script:Palette.Navy
    $hero.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    for ($row = 0; $row -lt 4; $row++) {
        $hero.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    }
    $heroLabel = New-UiLabel -Text 'ONAT  ·  ENTORNO DE DESARROLLO' -FontSize 8.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(151, 194, 203))
    $heroLabel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 4)
    $heroTitle = New-UiLabel -Text 'Todo listo para poner el proyecto en marcha.' -FontSize 17 -Bold $true -ForeColor $script:Palette.White
    $heroTitle.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
    $heroText = New-UiLabel -Text 'Inicia la Web local, valida los cambios y consulta los resultados sin salir del panel.' -FontSize 10 -ForeColor ([System.Drawing.Color]::FromArgb(205, 220, 226))
    $heroText.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
    $heroCaption = New-UiLabel -Text 'LOCAL  ·  HTTPS  ·  SIN MODIFICAR LOS DATOS' -FontSize 8.5 -Bold $true -ForeColor ([System.Drawing.Color]::FromArgb(151, 194, 203))
    $hero.Controls.Add($heroLabel, 0, 0)
    $hero.Controls.Add($heroTitle, 0, 1)
    $hero.Controls.Add($heroText, 0, 2)
    $hero.Controls.Add($heroCaption, 0, 3)
    $hero.Tag = [pscustomobject]@{ Body = $heroText }
    $hero.add_SizeChanged({
        param($sender, $eventArgs)
        if ($sender.Tag -and $sender.Tag.Body -and $sender.ClientSize.Width -gt 0) {
            $availableWidth = [Math]::Max(240, $sender.ClientSize.Width - $sender.Padding.Horizontal - 4)
            $newMaximum = [System.Drawing.Size]::new($availableWidth, 0)
            if ($sender.Tag.Body.MaximumSize -ne $newMaximum) {
                $sender.Tag.Body.MaximumSize = $newMaximum
            }
        }
    })

    # Estado del entorno: insignias en una fila con ajuste automático y ruta en una fila aparte.
    $environment = New-Object System.Windows.Forms.TableLayoutPanel
    $environment.Dock = [System.Windows.Forms.DockStyle]::Fill
    $environment.AutoSize = $true
    $environment.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $environment.ColumnCount = 1
    $environment.RowCount = 2
    $environment.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    $environment.Padding = [System.Windows.Forms.Padding]::new(12)
    $environment.BackColor = $script:Palette.White
    $environment.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $environment.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $environment.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $environment.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    $environmentHeader = New-Object System.Windows.Forms.FlowLayoutPanel
    $environmentHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
    $environmentHeader.AutoSize = $true
    $environmentHeader.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $environmentHeader.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $environmentHeader.WrapContents = $true
    $environmentHeader.Margin = [System.Windows.Forms.Padding]::Empty
    $environmentHeader.Padding = [System.Windows.Forms.Padding]::Empty
    $environmentHeader.BackColor = [System.Drawing.Color]::Transparent
    $environmentTitle = New-UiLabel -Text 'ESTADO DEL ENTORNO' -FontSize 8.5 -Bold $true -ForeColor $script:Palette.Muted
    $environmentTitle.Margin = [System.Windows.Forms.Padding]::new(0, 7, 12, 7)

    $script:NodeBadge = New-Object System.Windows.Forms.Label
    $script:NodeBadge.AutoSize = $true
    $script:NodeBadge.Padding = [System.Windows.Forms.Padding]::new(9, 5, 9, 5)
    $script:NodeBadge.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    $script:NodeBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:NodeBadge.Font = [System.Drawing.Font]::new('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $script:NpmBadge = New-Object System.Windows.Forms.Label
    $script:NpmBadge.AutoSize = $true
    $script:NpmBadge.Padding = [System.Windows.Forms.Padding]::new(9, 5, 9, 5)
    $script:NpmBadge.Margin = [System.Windows.Forms.Padding]::Empty
    $script:NpmBadge.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $script:NpmBadge.Font = [System.Drawing.Font]::new('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $environmentHeader.Controls.Add($environmentTitle)
    $environmentHeader.Controls.Add($script:NodeBadge)
    $environmentHeader.Controls.Add($script:NpmBadge)

    $script:RepoPathLabel = New-UiLabel -Text '' -FontSize 9 -Bold $false -ForeColor ([System.Drawing.Color]::FromArgb(108, 126, 140))
    $script:RepoPathLabel.Margin = [System.Windows.Forms.Padding]::new(0, 7, 0, 0)
    $script:RepoPathLabel.MaximumSize = [System.Drawing.Size]::new(1000, 0)
    $script:PathToolTip = New-Object System.Windows.Forms.ToolTip
    $script:PathToolTip.AutoPopDelay = 12000
    $script:PathToolTip.SetToolTip($script:RepoPathLabel, $script:RepoRoot)
    $environment.Controls.Add($environmentHeader, 0, 0)
    $environment.Controls.Add($script:RepoPathLabel, 0, 1)
    $environment.Tag = [pscustomobject]@{ PathLabel = $script:RepoPathLabel }
    $environment.add_SizeChanged({
        param($sender, $eventArgs)
        if ($sender.Tag -and $sender.Tag.PathLabel -and $sender.ClientSize.Width -gt 0) {
            $availableWidth = [Math]::Max(200, $sender.ClientSize.Width - $sender.Padding.Horizontal - 4)
            $newMaximum = [System.Drawing.Size]::new($availableWidth, 0)
            if ($sender.Tag.PathLabel.MaximumSize -ne $newMaximum) {
                $sender.Tag.PathLabel.MaximumSize = $newMaximum
            }
        }
    })

    $actions = New-Object System.Windows.Forms.TableLayoutPanel
    $actions.Dock = [System.Windows.Forms.DockStyle]::Fill
    $actions.AutoSize = $true
    $actions.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $actions.ColumnCount = 2
    $actions.RowCount = 2
    $actions.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    $actions.Padding = [System.Windows.Forms.Padding]::Empty
    $actions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 50)) | Out-Null
    $actions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 50)) | Out-Null
    $actions.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $actions.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null

    $cardWeb = New-ActionCard -Title 'Abrir la aplicación' -Description 'Servidor seguro local y vista previa para pruebas rápidas.' `
        -PrimaryAction @{ Type = 'server'; Mode = 'https'; Text = 'Web HTTPS'; TaskName = 'Web segura' } `
        -SecondaryAction @{ Type = 'server'; Mode = 'preview'; Text = 'Vista previa HTTP'; TaskName = 'Vista previa' } `
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
    $logLayout.ColumnCount = 1
    $logLayout.RowCount = 2
    $logLayout.Margin = [System.Windows.Forms.Padding]::Empty
    $logLayout.Padding = [System.Windows.Forms.Padding]::Empty
    $logLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $logLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $logLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null

    $logHeading = New-Object System.Windows.Forms.TableLayoutPanel
    $logHeading.Dock = [System.Windows.Forms.DockStyle]::Fill
    $logHeading.AutoSize = $true
    $logHeading.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $logHeading.ColumnCount = 2
    $logHeading.RowCount = 1
    $logHeading.Margin = [System.Windows.Forms.Padding]::Empty
    $logHeading.Padding = [System.Windows.Forms.Padding]::new(12, 4, 10, 4)
    $logHeading.BackColor = $script:Palette.White
    $logHeading.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $logHeading.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $logHeading.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $logTitle = New-UiLabel -Text 'Actividad reciente' -FontSize 10.5 -Bold $true -ForeColor $script:Palette.Ink
    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = 'Limpiar'
    $clearButton.AutoSize = $true
    $clearButton.MinimumSize = [System.Drawing.Size]::new(72, 32)
    $clearButton.Padding = [System.Windows.Forms.Padding]::new(10, 5, 10, 5)
    $clearButton.Margin = [System.Windows.Forms.Padding]::Empty
    $clearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $clearButton.FlatAppearance.BorderSize = 0
    $clearButton.BackColor = $script:Palette.White
    $clearButton.ForeColor = $script:Palette.Muted
    $clearButton.Font = [System.Drawing.Font]::new('Segoe UI', 9)
    $clearButton.add_Click({ $script:OutputBox.Clear() })
    $logHeading.Controls.Add($logTitle, 0, 0)
    $logHeading.Controls.Add($clearButton, 1, 0)

    $script:OutputBox = New-Object System.Windows.Forms.RichTextBox
    $script:OutputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:OutputBox.ReadOnly = $true
    $script:OutputBox.DetectUrls = $false
    $script:OutputBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:OutputBox.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 252)
    $script:OutputBox.ForeColor = [System.Drawing.Color]::FromArgb(78, 96, 111)
    $script:OutputBox.Font = [System.Drawing.Font]::new('Consolas', 9)
    $script:OutputBox.WordWrap = $true
    $script:OutputBox.HideSelection = $false
    $script:OutputBox.MaxLength = 30000
    $script:OutputBox.Margin = [System.Windows.Forms.Padding]::new(8)
    $logLayout.Controls.Add($logHeading, 0, 0)
    $logLayout.Controls.Add($script:OutputBox, 0, 1)
    $logCard.Controls.Add($logLayout)

    $footer = New-Object System.Windows.Forms.TableLayoutPanel
    $footer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $footer.AutoSize = $true
    $footer.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $footer.ColumnCount = 2
    $footer.RowCount = 1
    $footer.Margin = [System.Windows.Forms.Padding]::Empty
    $footer.Padding = [System.Windows.Forms.Padding]::Empty
    $footer.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100)) | Out-Null
    $footer.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $footer.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize)) | Out-Null
    $footerLabel = New-UiLabel -Text 'ONAT  ·  Herramientas de desarrollo local' -FontSize 8.5 -ForeColor ([System.Drawing.Color]::FromArgb(138, 152, 164))
    $script:StatusText = New-UiLabel -Text 'Listo' -FontSize 8.5 -Bold $true -ForeColor $script:Palette.Green
    $script:StatusText.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $script:StatusText.Margin = [System.Windows.Forms.Padding]::new(12, 0, 0, 0)
    $footer.Controls.Add($footerLabel, 0, 0)
    $footer.Controls.Add($script:StatusText, 1, 0)

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
    if ($script:PathToolTip) { $script:PathToolTip.Dispose() }
}

Initialize-ControlPanel
