# -----------------------------
# Helper function: Attempt to resolve WSD IP (may return $null)
# -----------------------------
function Get-WsdIp {
    param (
        [string]$PortName
    )

    try {
        $port = Get-CimInstance Win32_TCPIPPrinterPort |
            Where-Object { $_.Name -eq $PortName }

        return $port.HostAddress
    }
    catch {
        return $null
    }
}

# -----------------------------
# Get printers and ports
# -----------------------------
$printers = Get-Printer | Where-Object {
    $_.DriverName -notmatch '^Microsoft' -and
    $_.Name       -notmatch 'PDF|OneNote|XPS' -and
    $_.PortName   -notin @('PORTPROMPT:', 'FILE:', 'NUL:')
}

$ports = Get-PrinterPort

# -----------------------------
# Build output
# -----------------------------
$result = foreach ($printer in $printers) {
    $port = $ports | Where-Object { $_.Name -eq $printer.PortName }

    $connectionType = switch -Regex ($printer.PortName) {
        '^USB'  { 'USB' }
        '^IP_'  { 'IP (Standard TCP/IP)' }
        '^\d{1,3}(\.\d{1,3}){3}$' { 'IP (Direct)' }
        '^WSD'  { 'WSD (Auto-discovered Network)' }
        '^COM'  { 'Serial (COM)' }
        '^LPT'  { 'Parallel (LPT)' }
        default { 'Unknown' }
    }

    # IP address (may be null for WSD)
    $ip = $port.PrinterHostAddress
    if (-not $ip -and $printer.PortName -like 'WSD*') {
        $ip = Get-WsdIp $printer.PortName
    }

    [PSCustomObject]@{
        PrinterName    = $printer.Name
        Driver         = $printer.DriverName
        PortName       = $printer.PortName
        ConnectionType = $connectionType
        IPAddress      = $ip
    }
}

# -----------------------------
# Output
# -----------------------------
$result |
    Sort-Object PrinterName |
    Format-Table -AutoSize
