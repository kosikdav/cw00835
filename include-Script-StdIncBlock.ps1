#default variable definition file - include-Var-Define.ps1
$DefaultVarDefFile = $ScriptPath + "\include\include-Var-Define.ps1"
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

. $IncFile_Var_Define
. $IncFile_Var_Init
. $IncFile_Functions_Common
