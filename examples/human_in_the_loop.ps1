param(
    [string]$BaseUrl = "http://127.0.0.1:8000",
    [string]$Token = $env:SIMLAB_API_TOKEN
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Token)) {
    throw "请先设置 SIMLAB_API_TOKEN，或通过 -Token 传入。"
}

$headers = @{ Authorization = "Bearer $Token" }
$configPath = Join-Path $PSScriptRoot "api_service_center.json"
$config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json

$sessionBody = @{
    config = $config
} | ConvertTo-Json -Depth 20
$session = Invoke-RestMethod `
    -Method Post `
    -Uri "$BaseUrl/v1/sessions" `
    -Headers $headers `
    -ContentType "application/json" `
    -Body $sessionBody

Write-Host "会话：$($session.session_id)，版本：$($session.workflow_version)"

$proposalBody = @{
    operation_id = "manual-$([guid]::NewGuid().ToString('N'))"
    expected_version = $session.workflow_version
    action = @{
        action = "set_parameter"
        path = "simulation.stations.1.capacity"
        value = 3
    }
} | ConvertTo-Json -Depth 10
$proposal = Invoke-RestMethod `
    -Method Post `
    -Uri "$BaseUrl/v1/sessions/$($session.session_id)/proposals" `
    -Headers $headers `
    -ContentType "application/json" `
    -Body $proposalBody

Write-Host "待审批提案：$($proposal.proposal_id)"
Write-Host ($proposal.action | ConvertTo-Json -Depth 10)
$decision = Read-Host "输入 APPROVE 批准，其余输入将拒绝"

$decisionBody = @{
    operation_id = "decision-$([guid]::NewGuid().ToString('N'))"
    expected_version = $proposal.updated_version
    reason = "PowerShell 人工演示"
} | ConvertTo-Json

if ($decision -ceq "APPROVE") {
    $proposal = Invoke-RestMethod `
        -Method Post `
        -Uri "$BaseUrl/v1/sessions/$($session.session_id)/proposals/$($proposal.proposal_id):approve" `
        -Headers $headers `
        -ContentType "application/json" `
        -Body $decisionBody
    Write-Host "已批准并应用，新状态：$($proposal.status)"
} else {
    $proposal = Invoke-RestMethod `
        -Method Post `
        -Uri "$BaseUrl/v1/sessions/$($session.session_id)/proposals/$($proposal.proposal_id):reject" `
        -Headers $headers `
        -ContentType "application/json" `
        -Body $decisionBody
    Write-Host "已拒绝，状态：$($proposal.status)"
    exit 0
}

$current = Invoke-RestMethod `
    -Method Get `
    -Uri "$BaseUrl/v1/sessions/$($session.session_id)" `
    -Headers $headers
$runBody = @{
    operation_id = "run-$([guid]::NewGuid().ToString('N'))"
    expected_version = $current.workflow_version
    workers = 1
} | ConvertTo-Json
$run = Invoke-RestMethod `
    -Method Post `
    -Uri "$BaseUrl/v1/sessions/$($session.session_id)/runs" `
    -Headers $headers `
    -ContentType "application/json" `
    -Body $runBody

Write-Host "运行完成：$($run.run_id)"
Write-Host "结果目录：$($run.output_dir)"
