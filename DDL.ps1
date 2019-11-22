<# Script que sirve para consultar y modificar
tanto las DDL como los grupo de seguridad asociados #>


###########################################
#                                         #
# Miramos si los módulos están instalados #
#                                         #
###########################################


$MSOnlineCheck = Get-Module -ListAvailable MsOnline

    if ($SQLModuleCheck -eq $null)
    {
        
        # Sino está instalado seleccionamos el repositorio y cambiamos la políticade instalación
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

        # Instalamos el Módulo SqlServer
        Install-Module -Name SqlServer –Scope AllUsers -Confirm:$false -AllowClobber
    }


##########################
#                        #
# Importamos los Módulos #
#                        #
##########################    

    Import-Module MsOnline

#############################
#                           #
# Definimos otras variables #
#                           #
#############################

    
    $Departamento = "ATC Medio Ambiente" 
    #$Departamento2 = "Calidad"       

    $Company = "AGQ Corporate" 
    #$Company2 = "AGQ Morocco"  

 
    $title1 = "Gerente Comercial"
    $title2 = "Gerente Departamento"
    $title3 = "Jefe de Laboratorio"
    $title4 = "Jefe Técnico inspección y muestreos"


    $AliasDDL = "DDL_ResponsablesATCMinería"
    $DynamicGroupName = "DDL_Responsables Laboratorio Organico Alimentaria"
    $groupName = "AGQ España - ATC Alimentaria"

###########################
#                         #
# Definimos las funciones #
#                         #
###########################

    #funcion para conectar a Office365
    function Conectar-O365{
        
        #Credenciales para conectarnos a O365
        $cred = Import-Clixml -Path C:\O365\key\credenciales.xml

        try{

            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
            Import-PSSession $exchangeSession -DisableNameChecking
            Connect-MsolService -Credential $cred

            }

        catch{

            $log = "$fecha --> Error intentando conectar a Office365: $Error"
            Escribir-Log -log $log -Path $Path 
            $Error.Clear()

        }            
    }


    function AgregaDDL{


        Set-DynamicDistributionGroup -Identity $DynamicGroupName -RecipientFilter "(((department -eq '$Departamento') -or (customAttribute15 -eq '$Departamento') -or (customAttribute14 -eq '$Departamento')) -and ((title -eq '$title1') -or (title -eq '$title2')  -or (title -eq '$title3')  -or (title -eq '$title4'))) -or ((company -eq '$company') -and ((department -eq '$Departamento') -or (customAttribute15 -eq '$Departamento') -or (customAttribute14 -eq '$Departamento')))"
    
    }



    function ConsultarDDL {


        #param ([String]$DynamicGroupName)
        $var = Get-DynamicDistributionGroup $DynamicGroupName     

        $Xs = Get-Recipient -RecipientPreviewFilter $var.RecipientFilter | Format-Table PrimarySmtpAddress, Title, Company, Department, customattribute14, customattribute15

        if ($Xs.Count -eq 0 ){
        
            write-output ("Número de recipientes en el grupo '"+ $DynamicGroupName + "': " + $Xs.Count)

        }else{

            write-output ("Número de recipientes en el grupo '"+ $DynamicGroupName + "': " + ($Xs.Count-4))
            Write-Output $Xs
        }

    }



    
    function SyncDDL{

        #param ([String]$AliasDDL)    
        $AliasDDL = "DDL_ResponsablesATCMineria"
        $DDL = Get-DynamicDistributionGroup | ? Alias -eq $AliasDDL
        $Members = Get-Recipient -RecipientPreviewFilter $DDL.RecipientFilter | Select DisplayName,PrimarySmtpAddress
        $Group = $DDL.Alias.Replace("DDL_","")
        Update-DistributionGroupMember -Identity $Group -Members $Members.PrimarySmtpAddress -Confirm:$false
        $Members, $Group, $DDL, $AliasDDL = $null

    }

   
    function ConsultaGrupo {

        #Grupo de seguridad
        
        #param([String]$groupName)

        $Xs = Get-DistributionGroupMember -Identity $groupName | Format-Table PrimarySmtpAddress, Title, Company, Department, customattribute14, customattribute15

        if($Xs.Count -eq 0){

            write-output ("Número de recipientes en el grupo '"+ $groupName + "': " + $Xs.Count)

        }else{

            write-output ("Número de recipientes en el grupo '"+ $groupName + "': " + ($Xs.Count-4))

        }
        Write-Output $Xs
}

#------------------------ Sincronizamos Todos Grupos en O365 --------------------------#

    function SyncDDLs {

        $DDLs = Get-DynamicDistributionGroup | ? {($_.Alias -like "DDL_*")}

        ForEach ($DDL in $DDLs)
        {
            $Members = Get-Recipient -RecipientPreviewFilter $DDL.RecipientFilter | Select DisplayName,PrimarySmtpAddress
            $Group = $DDL.Alias.Replace("DDL_","")
            Update-DistributionGroupMember -Identity $Group -Members $Members.PrimarySmtpAddress -Confirm:$false
            $Members, $Group = $null
        }

    }


    


#Establecemos el propietario y la cuenta de correo predeterminada

    function CreaGrupo {

        
        $alias = "AGQEspanaLaboratorioOrganicoMedioambiental"
        $displayname = "AGQ España - Laboratorio Orgánico Medioambiental"
        $correo = $alias+"@agqlabs.com"

        New-DistributionGroup -Alias $alias -DisplayName $displayname -Name $displayname -ManagedBy admin@agqlabs.onmicrosoft.com -PrimarySmtpAddress $correo -Type Security

    }



<#Creamos un DDL desde powershell estableciendo la cuenta de 
correo predeterminada#>
    
    
    function CreaDDL{
    
        New-DynamicDistributionGroup -Alias $alias -DisplayName $displayname -Name $nombre -PrimarySmtpAddress $correo

    }

    function Desconectar{

        Get-PSSession | Remove-PSSession

    }



