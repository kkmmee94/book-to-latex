param(
    [string]$Weights = $env:BOOK_TO_LATEX_QWEN_WEIGHTS
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Weights)) {
    throw "Provide the Hugging Face model folder with -Weights or set BOOK_TO_LATEX_QWEN_WEIGHTS."
}
$weights = (Resolve-Path -LiteralPath $Weights).Path
$modelRoot = Split-Path -Parent $weights
$f16Gguf = Join-Path $modelRoot "Qwen3-8B-local-F16.gguf"
$q4Gguf = Join-Path $modelRoot "Qwen3-8B-local-Q4_K_M.gguf"
$llamaRoot = Join-Path $projectRoot "tools\llama.cpp"
$converter = Join-Path $llamaRoot "convert_hf_to_gguf.py"
$modelTemplate = Join-Path $projectRoot "models\qwen3-local-uncensored-label.Modelfile"
$modelFile = Join-Path $projectRoot "models\qwen3-local-uncensored-label.generated.Modelfile"
$modelName = "book-latex-qwen3-local-uncensored:8b"

Set-Location -LiteralPath $projectRoot
if (-not (Test-Path -LiteralPath $weights -PathType Container)) {
    throw "The local Qwen3 folder was not found: $weights"
}
if (-not (Get-Command ollama -ErrorAction SilentlyContinue)) {
    throw "Ollama is not installed."
}
if (-not (Get-Command llama-quantize -ErrorAction SilentlyContinue)) {
    throw "llama.cpp is not installed. Install the winget package ggml.llamacpp."
}

if (-not (Test-Path -LiteralPath $q4Gguf -PathType Leaf)) {
    if (-not (Test-Path -LiteralPath $converter -PathType Leaf)) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $llamaRoot) -Force | Out-Null
        git clone --depth 1 https://github.com/ggml-org/llama.cpp.git $llamaRoot
    }
    python -m pip install -r (Join-Path $llamaRoot "requirements\requirements-convert_hf_to_gguf.txt")
    python $converter $weights --outfile $f16Gguf --outtype f16 --model-name "Qwen3-8B local exact weights"
    if ($LASTEXITCODE -ne 0) { throw "Safetensors-to-GGUF conversion failed." }
    llama-quantize $f16Gguf $q4Gguf Q4_K_M
    if ($LASTEXITCODE -ne 0) { throw "GGUF quantization failed." }
    if (Test-Path -LiteralPath $f16Gguf -PathType Leaf) {
        Remove-Item -LiteralPath $f16Gguf -Force
    }
    python -m pip install -r (Join-Path $projectRoot "requirements-dev.txt")
}
if ((Test-Path -LiteralPath $q4Gguf -PathType Leaf) -and (Test-Path -LiteralPath $f16Gguf -PathType Leaf)) {
    Remove-Item -LiteralPath $f16Gguf -Force
    Write-Host "Removed the temporary 16 GB F16 GGUF after successful Q4 quantization."
}

$ollamaPath = $q4Gguf.Replace("\", "/")
$template = Get-Content -LiteralPath $modelTemplate -Raw
Set-Content -LiteralPath $modelFile -Value $template.Replace("__GGUF_PATH__", $ollamaPath) -Encoding UTF8

ollama create --file $modelFile $modelName
if ($LASTEXITCODE -ne 0) { throw "Ollama model creation failed." }

$testBody = @{
    model = $modelName
    messages = @(@{ role = "user"; content = "Reply with exactly LOCAL UNCENSORED MODEL READY" })
    options = @{ temperature = 0; num_predict = 32; num_ctx = 8192 }
    think = $false
    stream = $false
} | ConvertTo-Json -Depth 6
$response = Invoke-RestMethod -Uri "http://127.0.0.1:11434/api/chat" -Method Post -ContentType "application/json" -Body $testBody -TimeoutSec 300
if ($response.message.content.Trim() -ne "LOCAL UNCENSORED MODEL READY") {
    throw "Unexpected test response: $($response.message.content)"
}
Write-Host "$modelName is installed, tested, and ready for the app."
