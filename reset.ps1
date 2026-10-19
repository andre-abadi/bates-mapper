Get-ChildItem -Path . -Filter *.zip | ForEach-Object {

    Expand-Archive -Path $_.FullName -DestinationPath . -Force

}