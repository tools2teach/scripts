try {
    Start-Process -FilePath 'TigerConnect-Setup.exe' -ArgumentList ('/s','TigerConnect-Setup.exe') -Wait -Passthru
    Exit 0
} catch { $Exception = $error[0].Exception.Message + "`nAt Line " + $error[0].InvocationInfo.ScriptLineNumber
    Write-Error $Exception
    Exit -1
}