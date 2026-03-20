#default variable definition file - include-Var-Define.ps1
$DefaultVarDefFile = $ScriptPath + "\include\include-Var-Define.ps1"
if ($VariableDefinitionFile) {
    if (Test-Path -Path $VariableDefinitionFile) {
        $IncFile_Var_Define = $VariableDefinitionFile
        write-Host "Using parameter variable definition file: $VariableDefinitionFile"
    }
}
    else {
        if ($env:DEFAULT_VARDEFINITIONFILE -and (Test-Path -Path $env:DEFAULT_VARDEFINITIONFILE)) {
            $IncFile_Var_Define = $env:DEFAULT_VARDEFINITIONFILE
            write-Host "Using variable definition file from environment variable: $env:DEFAULT_VARDEFINITIONFILE"
        }
        else {
            if (Test-Path -Path $DefaultVarDefFile) {
                $IncFile_Var_Define = $DefaultVarDefFile
                write-Host "Using static default variable definition file: $DefaultVarDefFile"
            }
            else {
                Write-Host "Default variable definition file '$DefaultVarDefFile' not found. Cannot proceed."
                exit 1
            }
        }
    }    
write-Host "Variable definition file: $IncFile_Var_Define"
. $IncFile_Var_Define

write-host "Var init file: $IncFile_Var_Init"
write-host "Functions common file: $IncFile_Functions_Common"

. $IncFile_Var_Init
. $IncFile_Functions_Common
