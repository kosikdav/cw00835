#default variable definition file - include-Var-Define.ps1
$DefaultVarDefFile = "d:\scripts-m365\cezdata\include-var-define-CEZDATA.ps1"

if ($VariableDefinitionFile) {
    if (Test-Path -Path $VariableDefinitionFile) {
        $IncFile_Var_Define = $VariableDefinitionFile
    }
}
else {
    if ($env:DEFAULT_VARDEFINITIONFILE -and (Test-Path -Path $env:DEFAULT_VARDEFINITIONFILE)) {
        $IncFile_Var_Define = $env:DEFAULT_VARDEFINITIONFILE
    }
    else {
        if (Test-Path -Path $DefaultVarDefFile) {
            $IncFile_Var_Define = $DefaultVarDefFile
        }
        else {
            exit
        }
    }
}

#write-host $IncFile_Var_Define -ForegroundColor Green
. $IncFile_Var_Define
#write-host $IncFile_Var_Init -ForegroundColor Green
. $IncFile_Var_Init
#write-host $IncFile_Functions_Common -ForegroundColor Green
. $IncFile_Functions_Common
