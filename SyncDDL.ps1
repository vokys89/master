<#
 Este script SyncDDLtoSG sincroniza las listas dinámicas que empiezan con DDL y actualiza 
 (reemplaza) los miembros en los Grupos de Seguridad/correo correspondientes cuyo Alias 
 es lo mismo que la lista dinámica pero sin "DDL_".
  Febrero 2019 - © Christian Cevallos



$credenciales = Get-Credential

$credenciales | Export-Clixml -Path C:\O365\key\credenciales.xml


#>

$cred = Import-Clixml -Path C:\O365\key\credenciales.xml


Import-Module MsOnline
#Conectar con credenciales previamente encriptadas


  
#Conectarse a Exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

$DDLs = Get-DynamicDistributionGroup | ? {($_.Alias -like "DDL_*")}
ForEach ($DDL in $DDLs)
{
   $Members = Get-Recipient -RecipientPreviewFilter $DDL.RecipientFilter | Select DisplayName,PrimarySmtpAddress
   $Group = $DDL.Alias.Replace("DDL_","")
   Update-DistributionGroupMember -Identity $Group -Members $Members.PrimarySmtpAddress -Confirm:$false
   $Members, $Group = $null   

}


#------------------Excepciones-----------#

Add-DistributionGroupMember -Identity "AGQCorporateIT" -Member zdominguez@agqlabs.com -Confirm:$false
Add-DistributionGroupMember -Identity "ResponsablesATCAgronomia" -Member abresolin@agqlabs.com -Confirm:$false
Add-DistributionGroupMember -Identity "ResponsablesATCAlimentaria" -Member paloma.franco@agqlabs.com -Confirm:$false
Add-DistributionGroupMember -Identity "ResponsablesATCAlimentaria" -Member jjmarchena@agqlabs.com -Confirm:$false
Add-DistributionGroupMember -Identity "ResponsablesATCAlimentaria" -Member egarrido@agqlabs.com -Confirm:$false
Add-DistributionGroupMember -Identity "Tecnologia" -Member bcolmenero@agqlabs.com -Confirm:$false


#-----------------------------------------#
Get-PSSession | Remove-PSSession
