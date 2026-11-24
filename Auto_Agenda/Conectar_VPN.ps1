# Script de Conexao Automatica da VPN Engelmig (FortiClient)
$ErrorActionPreference = "SilentlyContinue"

# 1. Verifica se a VPN ja esta conectada
$vpnAtiva = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*Fortinet*" -and $_.Status -eq "Up" }
if ($vpnAtiva) {
    Write-Host "[OK] VPN Engelmig ja esta conectada e ativa!"
    exit 0
}

Write-Host "Verificando FortiClient..."
$fortiDir = "C:\Program Files\Fortinet\FortiClient"
$fortiPath = Join-Path $fortiDir "FortiClient.exe"

if (-not (Test-Path $fortiPath)) {
    Write-Host "[ERRO] FortiClient nao encontrado em: $fortiPath"
    exit 1
}

# Define o PATH para incluir as dependencias nativas do FortiClient (.node e DLLs C++)
$env:PATH = "$fortiDir;" + $env:PATH

# Se nao estiver aberto, abre com WorkingDirectory correto para evitar erro de JS
$procForti = Get-Process FortiClient -ErrorAction SilentlyContinue
if (-not $procForti) {
    Write-Host "Iniciando FortiClient com WorkingDirectory configurado..."
    Start-Process -FilePath $fortiPath -WorkingDirectory $fortiDir
    Start-Sleep -Seconds 4
} else {
    Write-Host "FortiClient ja em execucao, focando janela..."
}

# 3. Foca na janela e pressiona Enter no botao Conectar
$wshell = New-Object -ComObject WScript.Shell
for ($i = 0; $i -lt 8; $i++) {
    if ($wshell.AppActivate("FortiClient")) {
        Start-Sleep -Milliseconds 800
        $wshell.SendKeys("{ENTER}")
        Write-Host "Comando Conectar enviado!"
        break
    }
    Start-Sleep -Seconds 1
}

# 4. Aguarda a conexao ser estabelecida (ate 25s)
Write-Host "Aguardando conexao da VPN..."
for ($i = 0; $i -lt 25; $i++) {
    Start-Sleep -Seconds 1
    $vpnStatus = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*Fortinet*" -and $_.Status -eq "Up" }
    if ($vpnStatus) {
        Write-Host "[OK] VPN Engelmig conectada com sucesso!"
        exit 0
    }
}

Write-Host "[AVISO] Tempo limite de conexao esgotado. Verifique o FortiClient."
exit 0
