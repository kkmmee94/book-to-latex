$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$modelName = "book-latex-qwen3:8b"
$visionModelName = "book-latex-qwen35-vision:9b"
$templatePath = Join-Path $projectRoot "models\qwen3-book-latex.Modelfile"
$generatedPath = Join-Path $projectRoot "models\qwen3-book-latex.generated.Modelfile"
$visionTemplatePath = Join-Path $projectRoot "models\qwen35-vision-book-latex.Modelfile"
$visionGeneratedPath = Join-Path $projectRoot "models\qwen35-vision-book-latex.generated.Modelfile"

if (-not (Get-Command ollama -ErrorAction SilentlyContinue)) {
    throw "Ollama is not installed or is not available in PATH. Install Ollama and restart PowerShell."
}
$template = Get-Content -LiteralPath $templatePath -Raw
Set-Content -LiteralPath $generatedPath -Value $template -Encoding UTF8

Write-Host "Pulling Ollama's Windows-compatible Q4_K_M package for the same base model..."
ollama pull qwen3:8b
if ($LASTEXITCODE -ne 0) {
    throw "Ollama could not obtain its Windows-compatible Qwen3-8B runtime model."
}

Write-Host "Creating $modelName with the Book-to-LaTeX system instructions..."
ollama create --file $generatedPath $modelName
if ($LASTEXITCODE -ne 0) {
    throw "Ollama could not create $modelName"
}

Write-Host "Testing the model..."
$testBody = @{
    model = $modelName
    messages = @(@{ role = "user"; content = "Reply with exactly MODEL READY" })
    options = @{ temperature = 0; num_predict = 32 }
    think = $false
    stream = $false
} | ConvertTo-Json -Depth 6
$testResponse = Invoke-RestMethod -Uri "http://127.0.0.1:11434/api/chat" -Method Post -ContentType "application/json" -Body $testBody -TimeoutSec 180
if ($testResponse.message.content.Trim() -ne "MODEL READY") {
    throw "The model was created but returned an unexpected test response: $($testResponse.message.content)"
}

Write-Host "$modelName is ready."

Write-Host "Pulling the vision model used for mathematical and visual PDFs..."
ollama pull qwen3.5:9b
if ($LASTEXITCODE -ne 0) {
    throw "Ollama could not obtain the Qwen3.5 vision runtime model."
}
Copy-Item -LiteralPath $visionTemplatePath -Destination $visionGeneratedPath -Force
Write-Host "Creating $visionModelName with the visual document instructions..."
ollama create --file $visionGeneratedPath $visionModelName
if ($LASTEXITCODE -ne 0) {
    throw "Ollama could not create $visionModelName"
}

Write-Host "All local models are ready. Restart the app or select Refresh local models."
