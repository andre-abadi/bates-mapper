Remove-Item -LiteralPath "$PSScriptRoot\incorrect.docx" -Force -ErrorAction SilentlyContinue
Expand-Archive -LiteralPath "$PSScriptRoot\incorrect.zip" -DestinationPath $PSScriptRoot -Force
