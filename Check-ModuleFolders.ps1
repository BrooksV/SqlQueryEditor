ForEach ($path in ($env:PSModulePath -split ';')) {
    If (([System.IO.DirectoryInfo]($path)).Exists) {
        Get-ChildItem -Path $path -Directory
    }
}
