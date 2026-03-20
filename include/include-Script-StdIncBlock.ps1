#default variable definition file - include-Var-Define.ps1
$IncFile_Var_Define = $ScriptPath + "\include-Var-Define.ps1"
if ($VariableDefinitionFile) {
    if (Test-Path -Path $VariableDefinitionFile) {
        $IncFile_Var_Define = $VariableDefinitionFile
    }    
}

. $IncFile_Var_Define
. $IncFile_Var_Init
. $IncFile_Functions_Common
