Class Logger {
    [string] $_LogFilePath;
    [string] $_LogName;

    Logger([string] $logName) {
        $this._LogName = $logName;
        $LogsFolder = Join-Path (Join-Path (Join-Path (Join-Path ([Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData)) "CompuTec") "ProcessForce") "Scripts") "Logs"
        $LogNameDatePart = (Get-Date).ToString('O').Replace('-', '_').Replace(':', '_').Replace('.', '_') + '.log'
        $LogNameFileName = $logName + "_" + $LogNameDatePart;
        if ((Test-Path $LogsFolder) -eq $false) {
            New-Item -ItemType Directory -Force -Path $LogsFolder
        }
        $this._LogFilePath = Join-Path $LogsFolder $LogNameFileName
    }
    
    [void] WriteLog($msg) {
        $LogPath = $this._LogFilePath;
        Add-Content -Path $LogPath -Value $msg
        return;
    }
}




