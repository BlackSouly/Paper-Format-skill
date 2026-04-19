$ErrorActionPreference = "Stop"
$env:PYTHONPATH = Join-Path $PSScriptRoot "src"

python -c "from paper_format_normalizer.cli import main; main()" normalize-batch
