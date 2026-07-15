#Requires -Version 7.0

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------
$TenantId        = ""
$AppId           = ""
$Thumbprint      = ""
$Organization    = "interceramic.com"

$OutputFolder      = "D:\Nuevo\Scripts\Powershell_v7\ReporteFull_O365\Reportes"
$Fecha             = Get-Date -Format "ddMMyyyy_HHmmss"
$CsvPath          = Join-Path $OutputFolder "Reporte_Detallado_M365_$Fecha.csv"

$MailboxUsagePeriod           = "D180"
$EmailActivityPeriod          = "D180"
$GraphReportRetryCount        = 3
$GraphReportRetryDelaySeconds = 8

# ------------------------------------------------------------
# PREPARACIÓN
# ------------------------------------------------------------
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ------------------------------------------------------------
# FUNCIONES DE CONEXIÓN
# ------------------------------------------------------------
function Ensure-ExchangeOnlineConnection {
    param([string]$Organization, [string]$AppId, [string]$Thumbprint)
    $exoConnected = $false
    try {
        $exoSessions = Get-ConnectionInformation -ErrorAction Stop
        if ($exoSessions | Where-Object { $_.TokenStatus -eq "Active" -and $_.State -eq "Connected" }) {
            $exoConnected = $true
            Write-Host "✅ Ya existe conexión activa a Exchange Online." -ForegroundColor Green
        }
    } catch { $exoConnected = $false }

    if (-not $exoConnected) {
        Write-Host "Conectando a Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $Organization -ShowBanner:$false -ErrorAction Stop
        Write-Host "✅ Conexión a Exchange Online establecida." -ForegroundColor Green
    }
}

function Ensure-MgGraphConnection {
    param([string]$TenantId, [string]$AppId, [string]$Thumbprint)
    $mgConnected = $false
    try {
        $mgContext = Get-MgContext
        if ($mgContext) {
            $mgConnected = $true
            Write-Host "✅ Ya existe conexión activa a Microsoft Graph." -ForegroundColor Green
        }
    } catch { $mgConnected = $false }

    if (-not $mgConnected) {
        Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop
        Write-Host "✅ Conexión a Microsoft Graph establecida." -ForegroundColor Green
    }
}

# ------------------------------------------------------------
# FUNCIONES AUXILIARES
# ------------------------------------------------------------
function ConvertTo-MB {
    param($Value)
    if ($null -eq $Value) { return $null }
    try {
        if ($Value -and $Value.Value -and ($Value.Value | Get-Member -Name ToBytes -MemberType Method -ErrorAction SilentlyContinue)) {
            return [math]::Round(($Value.Value.ToBytes() / 1MB), 2)
        }
        if ($Value -and ($Value | Get-Member -Name ToBytes -MemberType Method -ErrorAction SilentlyContinue)) {
            return [math]::Round(($Value.ToBytes() / 1MB), 2)
        }
        $stringValue = $Value.ToString()
        if ($stringValue -match '\(([\d,]+)\sbytes\)') {
            $bytes = [int64](($matches[1] -replace ',', ''))
            return [math]::Round(($bytes / 1MB), 2)
        }
    } catch { return $null }
    return $null
}

function ConvertTo-NullableInt {
    param($Value)
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    try { return [int]$Value } catch { return $null }
}

function Test-DownloadedCsvIsValid {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $firstLines = Get-Content -Path $Path -TotalCount 10 -ErrorAction Stop
        $contentPreview = ($firstLines -join "`n").ToLower()
        if ([string]::IsNullOrWhiteSpace($contentPreview)) { return $false }
        if ($contentPreview -match '<html' -or $contentPreview -match '502 bad gateway' -or $contentPreview -match 'nginx') { return $false }
        return $true
    } catch { return $false }
}

function Invoke-GraphReportDownload {
    param([scriptblock]$Command, [string]$OutFilePath, [string]$ReportName, [int]$RetryCount = 3, [int]$RetryDelaySeconds = 8)
    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            if (Test-Path $OutFilePath) { Remove-Item -Path $OutFilePath -Force -ErrorAction SilentlyContinue }
            & $Command
            if (-not (Test-DownloadedCsvIsValid -Path $OutFilePath)) { throw "Archivo no válido" }
            return $true
        } catch {
            Write-Warning "Intento $attempt/$RetryCount falló para ${ReportName}"
            if ($attempt -lt $RetryCount) { Start-Sleep -Seconds $RetryDelaySeconds }
        }
    }
    return $false
}

# ------------------------------------------------------------
# MAPEO DE LICENCIAS
# ------------------------------------------------------------
$LicenseSkuMap = [ordered]@{
    "Exchange Online Kiosk"                                       = @("EXCHANGEDESKLESS")
    "Exchange Online Plan 1"                                      = @("EXCHANGESTANDARD","EXCHANGE_S_PLAN_1")
    "Exchange Online Plan 2"                                      = @("EXCHANGEENTERPRISE","EXCHANGE_S_PLAN_2")
    "Office 365 E1"                                               = @("STANDARDPACK")
    "Office 365 E3"                                               = @("ENTERPRISEPACK")
    "Microsoft 365 E3"                                            = @("SPE_E3")
    "Aplicaciones de Microsoft 365 para empresas"                = @("OFFICESUBSCRIPTION")
    "Aplicaciones de Microsoft 365 para negocios"                = @("O365_BUSINESS")
    "Gobierno y protección de información de Microsoft 365 E5" = @("M365_INFO_PROTECTION_GOVERNANCE","INFORMATION_PROTECTION_AND_GOVERNANCE","MICROSOFT_365_INFORMATION_PROTECTION_AND_GOVERNANCE","M365_INFORMATION_PROTECTION_AND_GOVERNANCE")
    "Licencia de usuario de Microsoft Copilot Studio"            = @("VIRTUAL_AGENT_USL")
    "Microsoft 365 Copilot"                                       = @("Microsoft_365_Copilot")
    "Microsoft 365 Copilot para empresas"                         = @("MICROSOFT_365_COPILOT_FOR_BUSINESS")
    "Plan 1 de administración avanzada de SharePoint"            = @("SharePoint_advanced_management_plan_1")
    "Power Apps Premium"                                          = @("POWERAPPS_PER_USER")
    "Power BI Premium por usuario"                               = @("PBI_PREMIUM_PER_USER")
    "SharePoint (Plan 1)"                                        = @("SHAREPOINTSTANDARD")
    "Power Automate Premium"                                      = @("POWERAUTOMATE_ATTENDED_RPA","FLOW_PREMIUM")
    "Power BI Pro"                                               = @("POWER_BI_PRO")
    "Microsoft Entra ID"                                          = @("AAD_PREMIUM","AAD_PREMIUM_P1","AAD_PREMIUM_P2")
}

$LicenseColumnsOrdered = $LicenseSkuMap.Keys

# ------------------------------------------------------------
# SELECCIÓN DEL MODO DE USUARIOS
# ------------------------------------------------------------
do {
    Write-Host "`nSelecciona cómo deseas generar el reporte:" -ForegroundColor Cyan
    Write-Host "1. Usuarios desde archivo TXT"
    Write-Host "2. Todos los buzones de Office 365"
    $seleccionModo = Read-Host "Escribe 1 o 2"
} until ($seleccionModo -in @("1", "2"))

$ModoUsuarios = if ($seleccionModo -eq "1") { "Txt" } else { "Todos" }
$ArchivoUsuarios = Join-Path $OutputFolder "usuarios.txt"

if ($ModoUsuarios -eq "Txt") {
    $usarRutaDefault = Read-Host "Deseas usar la ruta por defecto del TXT ($ArchivoUsuarios)? (S/N)"
    if ($usarRutaDefault -notin @("S","s","SI","Si","si")) {
        $rutaIngresada = Read-Host "Ingresa la ruta completa del archivo TXT"
        if (-not [string]::IsNullOrWhiteSpace($rutaIngresada)) { $ArchivoUsuarios = $rutaIngresada.Trim() }
    }
}

# ------------------------------------------------------------
# CONEXIONES Y REPORTES GRAPH
# ------------------------------------------------------------
Ensure-ExchangeOnlineConnection -Organization $Organization -AppId $AppId -Thumbprint $Thumbprint
Ensure-MgGraphConnection -TenantId $TenantId -AppId $AppId -Thumbprint $Thumbprint

Write-Host "Obteniendo SKUs y reportes de Graph..." -ForegroundColor Cyan
$allSkus = Get-MgSubscribedSku -All
$skuPartNumberByGuid = @{}
foreach ($sku in $allSkus) { if ($sku.SkuId) { $skuPartNumberByGuid[$sku.SkuId.ToString()] = $sku.SkuPartNumber } }

# Reporte Uso
$mailboxUsageIndex = @{}
$UsageReportPath = Join-Path $OutputFolder "MailboxUsageDetail_$Fecha.csv"
if (Invoke-GraphReportDownload -Command { Get-MgReportMailboxUsageDetail -Period $MailboxUsagePeriod -OutFile $UsageReportPath } -OutFilePath $UsageReportPath -ReportName "Usage") {
    Import-Csv $UsageReportPath | ForEach-Object { if ($_.'User Principal Name') { $mailboxUsageIndex[$_.'User Principal Name'.ToLower()] = $_ } }
}

# Reporte Actividad
$emailActivityIndex = @{}
$ActivityReportPath = Join-Path $OutputFolder "EmailActivityUserDetail_$Fecha.csv"
if (Invoke-GraphReportDownload -Command { Get-MgReportEmailActivityUserDetail -Period $EmailActivityPeriod -OutFile $ActivityReportPath } -OutFilePath $ActivityReportPath -ReportName "Activity") {
    Import-Csv $ActivityReportPath | ForEach-Object { if ($_.'User Principal Name') { $emailActivityIndex[$_.'User Principal Name'.ToLower()] = $_ } }
}

# ------------------------------------------------------------
# CARGA DE BUZONES Y FORWARDING
# ------------------------------------------------------------
$mailboxes = New-Object System.Collections.Generic.List[object]
if ($ModoUsuarios -eq "Txt") {
    if (-not (Test-Path $ArchivoUsuarios)) { Write-Host "❌ Archivo TXT no encontrado." -ForegroundColor Red; return }
    $usuarios = Get-Content -Path $ArchivoUsuarios | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($u in $usuarios) {
        try { 
            $mbx = Get-EXOMailbox -Identity $u -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,WhenMailboxCreated,ProhibitSendReceiveQuota,Alias,CustomAttribute1,CustomAttribute2,CustomAttribute3 -ErrorAction Stop
            if ($mbx) { $mailboxes.Add($mbx) }
        } catch { Write-Warning "No se encontró buzón para $u" }
    }
} else {
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,WhenMailboxCreated,ProhibitSendReceiveQuota,Alias,CustomAttribute1,CustomAttribute2,CustomAttribute3
}

$forwardingIndex = @{}
try {
    Get-Mailbox -ResultSize Unlimited | Select-Object PrimarySmtpAddress,UserPrincipalName,ForwardingAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward | ForEach-Object {
        if ($_.UserPrincipalName) { $forwardingIndex[$($_.UserPrincipalName.ToLower())] = $_ }
    }
} catch {}

# ------------------------------------------------------------
# PROCESAR REPORTE DETALLADO
# ------------------------------------------------------------
$resultados = New-Object System.Collections.Generic.List[object]
$total = $mailboxes.Count
$curr = 0

foreach ($mbx in $mailboxes) {
    $curr++
    $identityToUse = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName } else { $mbx.PrimarySmtpAddress.ToString() }
    Write-Progress -Activity "Generando reporte" -Status "$curr de $total - $identityToUse" -PercentComplete (($curr/$total)*100)

    # --- DELEGACIONES ---
    $FullAccessFinal = @(); $SendAsFinal = @(); $GroupIdsToExpand = @(); $MembersUPN = @()
    
    if ($null -ne $mbx -and $mbx.RecipientTypeDetails -eq "SharedMailbox") {
        
        $FullAccessFinal = Get-MailboxPermission -Identity $identityToUse | 
            Where-Object { ($_.User -like '*@*' -or $_.User -notlike "NT AUTHORITY*") -and ($_.IsInherited -eq $false) } | 
            ForEach-Object { "$($_.User)" }
        $GroupIdsToExpand += $FullAccessFinal

        $SendAsFinal = Get-RecipientPermission -Identity $identityToUse | 
            Where-Object { $_.Trustee -ne "nt authority\self" } | 
            ForEach-Object {
                $t = "$($_.Trustee)"
                if ($t -match "^[0-9a-fA-F-]{36}$") {
                    $AzureObj = Get-MgGroup -GroupId $t -ErrorAction SilentlyContinue
                    if (-not $AzureObj) { $AzureObj = Get-MgUser -UserId $t -ErrorAction SilentlyContinue }
                    if ($AzureObj) { $t = $AzureObj.DisplayName }
                }
                $t
            }
        $GroupIdsToExpand += $SendAsFinal

        foreach ($Entry in ($GroupIdsToExpand | Select-Object -Unique)) {
            $GraphGroup = Get-MgGroup -Filter "DisplayName eq '$Entry' or MailNickname eq '$Entry' or Id eq '$Entry'" -ErrorAction SilentlyContinue
            if ($GraphGroup) {
                $Members = Get-MgGroupMember -GroupId $GraphGroup.Id -All -ErrorAction SilentlyContinue
                foreach ($M in $Members) {
                    $MemberDetail = Get-MgUser -UserId $M.Id -Property "UserPrincipalName" -ErrorAction SilentlyContinue
                    if ($MemberDetail.UserPrincipalName) { $MembersUPN += $MemberDetail.UserPrincipalName }
                }
            } else {
                $ExoMembers = Get-DistributionGroupMember -Identity $Entry -ErrorAction SilentlyContinue
                foreach ($Exm in $ExoMembers) { 
                    if ($Exm.PrimarySmtpAddress) { $MembersUPN += $Exm.PrimarySmtpAddress.ToString() }
                }
            }
        }
    }

    # --- DATOS COMPLEMENTARIOS ---
    $stats = Get-EXOMailboxStatistics -Identity $identityToUse -Properties LastLogoffTime,LastLogonTime,LastInteractionTime,LastUserActionTime,TotalItemSize -ErrorAction SilentlyContinue
    $mailboxUsedMB = ConvertTo-MB -Value $stats.TotalItemSize
    $mailboxTotalMB = ConvertTo-MB -Value $mbx.ProhibitSendReceiveQuota
    $usagePercent = if ($mailboxUsedMB -and $mailboxTotalMB -gt 0) { [math]::Round((($mailboxUsedMB / $mailboxTotalMB) * 100), 2) } else { $null }

    $entraUser = Get-MgUser -Filter "userPrincipalName eq '$($identityToUse.Replace("'", "''"))'" -Property "Id,AccountEnabled,SignInActivity,AssignedLicenses" -ErrorAction SilentlyContinue
    
    $licenseCols = [ordered]@{}; foreach ($lic in $LicenseColumnsOrdered) { $licenseCols[$lic] = "No" }
    if ($entraUser.AssignedLicenses) {
        foreach ($lic in $entraUser.AssignedLicenses) {
            $part = $skuPartNumberByGuid[$lic.SkuId.ToString()]
            foreach ($key in $LicenseSkuMap.Keys) { if ($LicenseSkuMap[$key] -contains $part) { $licenseCols[$key] = "Si" } }
        }
    }

    $fwd = $forwardingIndex[$identityToUse.ToLower()]
    
    $finalForwardingAddress = if ($fwd.ForwardingAddress) { $fwd.ForwardingAddress.ToString() } else { $null }
    if ($finalForwardingAddress -and ($finalForwardingAddress -notlike "*@*")) {
        try {
            $resolvedObj = Get-Recipient -Identity $finalForwardingAddress -ErrorAction SilentlyContinue
            if ($resolvedObj) {
                $finalForwardingAddress = if ($resolvedObj.PrimarySmtpAddress) { $resolvedObj.PrimarySmtpAddress.ToString() } else { $resolvedObj.Name }
            }
        } catch {}
    }

# --- CÁLCULO DE DÍAS RESTANTES (MÉTODO ROBUSTO) ---
    $diasRestantes = $null
    
    # Solo procesamos si ambos tienen contenido
    if (-not [string]::IsNullOrWhiteSpace($mbx.CustomAttribute2) -and -not [string]::IsNullOrWhiteSpace($mbx.CustomAttribute3)) {
        try {
            # Intentamos convertir directamente
            $fechaFin = [DateTime]::Parse($mbx.CustomAttribute3)
            $fechaActual = Get-Date
            $diferencia = $fechaFin - $fechaActual
            $diasRestantes = [math]::Floor($diferencia.TotalDays)
        } catch {
            # Si el formato no es válido, no hacemos nada (el valor permanece como $null)
            $diasRestantes = $null
        }
    }

    # --- CONSTRUCCIÓN DEL OBJETO FINAL ---
    $row = [ordered]@{
        DisplayName                = $mbx.DisplayName
        UserPrincipalName          = $mbx.UserPrincipalName
        PrimarySmtpAddress         = if ($mbx.PrimarySmtpAddress) { $mbx.PrimarySmtpAddress.ToString() } else { $null }
        RecipientTypeDetails       = $mbx.RecipientTypeDetails       
        FullAccess                 = ($FullAccessFinal | Select-Object -Unique) -join [char]10
        SendAs                     = ($SendAsFinal | Select-Object -Unique) -join [char]10
        MembersOf                  = ($MembersUPN | Select-Object -Unique) -join [char]10
        
        CustomAttribute1           = $mbx.CustomAttribute1
        CustomAttribute2           = $mbx.CustomAttribute2
        CustomAttribute3           = $mbx.CustomAttribute3
        "Días restantes"           = $diasRestantes

        WhenMailboxCreated         = $mbx.WhenMailboxCreated
        EntraAccountStatus         = if ($null -ne $entraUser.AccountEnabled) { if ($entraUser.AccountEnabled) { "Enabled" } else { "Disabled" } } else { $null }
        EntraLastSuccessfulSignInDateTimeUtc = if ($entraUser.SignInActivity) { $entraUser.SignInActivity.LastSuccessfulSignInDateTime } else { $null }       
        Total_Envios               = ConvertTo-NullableInt -Value $emailActivityIndex[$identityToUse.ToLower()].'Send Count'
        Total_Recepcion            = ConvertTo-NullableInt -Value $emailActivityIndex[$identityToUse.ToLower()].'Receive Count'
        Total_Leidos               = ConvertTo-NullableInt -Value $emailActivityIndex[$identityToUse.ToLower()].'Read Count'
        Total_Meets_Creados        = ConvertTo-NullableInt -Value $emailActivityIndex[$identityToUse.ToLower()].'Meeting Created Count'
        Total_Meets_Interac        = ConvertTo-NullableInt -Value $emailActivityIndex[$identityToUse.ToLower()].'Meeting Interacted Count'       
        Buzon_Ocupado_MB           = $mailboxUsedMB
        Buzon_Total_MB             = $mailboxTotalMB
        Porcentaje_Uso             = $usagePercent
        Ultima_Actividad           = $mailboxUsageIndex[$identityToUse.ToLower()].'Last Activity Date'       
        LastLogoffTime             = if ($stats) { $stats.LastLogoffTime } else { $null }
        LastLogonTime              = if ($stats) { $stats.LastLogonTime } else { $null }
        LastInteractionTime        = if ($stats) { $stats.LastInteractionTime } else { $null }
        LastUserActionTime         = if ($stats) { $stats.LastUserActionTime } else { $null }       
        ForwardingAddress          = $finalForwardingAddress
        ForwardingSmtpAddress      = if ($fwd.ForwardingSmtpAddress) { $fwd.ForwardingSmtpAddress.ToString() } else { $null }
        DeliverToMailboxAndForward = if ($null -ne $fwd.DeliverToMailboxAndForward) { $fwd.DeliverToMailboxAndForward } else { $null }
    }

    foreach ($lic in $LicenseColumnsOrdered) { $row[$lic] = $licenseCols[$lic] }
    $resultados.Add([PSCustomObject]$row)
}

# ------------------------------------------------------------
# EXPORTACIÓN FINAL
# ------------------------------------------------------------
$resultados | Sort-Object DisplayName | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding utf8BOM -Delimiter ";"

Write-Host "`n✅ Reporte generado correctamente: $CsvPath" -ForegroundColor Green
Write-Host "✅ Total registros: $($resultados.Count)" -ForegroundColor Green