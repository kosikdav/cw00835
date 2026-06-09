#######################################################################################################################
#######################################################################################################################
# INCLUDE-FUNCTIONS-COMMON
#######################################################################################################################
#######################################################################################################################
#
#
#
function Get-TargetT2TUser {
	[CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$SrcIdentity,
        [string]$SrcAccessToken,
		[string]$DstAccessToken,
		[pscredential]$DstADCredential
    )
	# main function body ##################################
	$UriResource = "users/$SrcIdentity"
	$UriSelect = "id,userPrincipalName,onpremisesSamAccountName,$($Src_PN_attr)"
	$Uri = New-GraphUri -Resource $UriResource -Select $UriSelect -Version "v1.0"
	$SrcUser = Get-GraphOutputREST -Uri $Uri -AccessToken $SrcAccessToken -ContentType $ContentTypeJSON
	if ($SrcUser) {
		#try standard user with mapping via PN
		if ($SrcUser.$Src_PN_attr) {
			$Filter = "$($Dst_Mapping_attr) -eq '$($SrcUser.$Src_PN_attr)'"
			$DstUser = Get-ADUser -Filter $Filter -Properties $Dst_Mapping_attr -Credential $DstADCredential -ErrorAction Stop
			if ($DstUser) {
				return $DstUser
			}
			else {
				return $null
			}
		}
		else {
			#try app user with mapping via UPN
			$Filter = "$($Dst_Mapping_attr) -eq '$($SrcUser.userPrincipalName)'"
			$DstUser = Get-ADUser -Filter $Filter -Properties $Dst_Mapping_attr -Credential $DstADCredential -ErrorAction Stop
			if ($DstUser) {
				return $DstUser
			}
			else {
				return $null
			}
		}
	}
	else {
		return $null
	}
}
