<#
.SYNOPSIS
    Con este script recorremos los buzones de la compañía y verificamos el uso de dicho buzón.

.DESCRIPTION
   Con este script recorremos los buzones de la compañía y verificamos el uso de dicho buzón. Si el tamaño es muy grande enviamos un correo de advertencia.


.NOTES

    File Name      : Advertencia-Buzon.ps1
    Author         : Efrain Daza (efrain.daza@agqlabs.com)
    Prerequisite   : PowerShell V5
    Copyright 2019 - Efrain Daza/Agq Labs.
    Version        : 1.0.0
#>


###########################################
#                                         #
# Miramos si los módulos están instalados #
#                                         #
###########################################

    $MSOnlineCheck = Get-Module -ListAvailable MsOnline
    
    if ($MSOnlineCheck -eq $null)
    {
        
        # Sino está instalado seleccionamos el repositorio y cambiamos la políticade instalación
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

        # Instalamos el Módulo MsOnline
        Install-Module -Name MsOnline –Scope AllUsers -Confirm:$false -AllowClobber
    }

##########################
#                        #
# Importamos los Módulos #
#                        #
##########################

    Import-Module MsOnline


###########################
#                         #
# Definimos las variables #
#                         #
###########################

    #Si hay algún error no controlado se para el Script
    $ErrorActionPreference = "Stop"    
         
    #Credenciales del buzón
    $CredencialesBuzon= Import-Clixml -Path C:\O365\key\credenciales_buzon.xml   
 
    #Configuramos el correo
    $Desde="noreply@agqlabs.com"
    $Para="notificacionesit@agqlabs.com"
    $Asunto="Buzon casi lleno"    
    $ServidorCorreo="smtp.office365.com"

    $Path = "C:\O365\Logs\Advertencia-Buzon.txt"
    $fecha = Get-Date -Format g

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

    #Función para enviar emails
    function Enviar-Email
    {   
        param($CredencialesBuzon,$Desde,$Para, $Asunto,$CuerpoEmail,$ServidorCorreo) 
        Try
        {   
         
            if($Para -ne $null)
            {
                [string[]]$To=$Para.Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
        
                Send-MailMessage -From $Desde -To $To -Subject $Asunto -Body $CuerpoEmail -BodyAsHtml -SmtpServer $ServidorCorreo -Credential $CredencialesBuzon -UseSsl -Port 587
            }    
        }
        catch [System.Exception]
        {
               $log = Write-Host $_.Exception.ToString()            
               Escribir-Log -log $log -Path $Path        
            
        } 
    }

    #Funcion Para escribir en los logs
    function Escribir-Log{

        param($log,$Path)

        if(($log -ne $null) -and ($Path -ne $null)){

            $log>>$Path

        }
    }

    #Función para cerrar las conexiones a Office365
    Function Desconectar-O365{

       Get-PSSession | Remove-PSSession

    }

    #Función Principal del Script
    function Comprobar-Usuarios{
        
        param($Usuarios)        

        foreach($Usuario in $Usuarios){

            #Guardamos el nombre del buzón en la variable
            $Nombre = $Usuario.UserPrincipalName

            try{

                $Statistics = @(Get-Mailbox -Identity $Nombre | Get-MailboxStatistics)
                $QuotaLimite = @(Get-Mailbox -Identity $Nombre)

            }catch{

                $log = "Se ha encontrado un error intentado obtener las propiedades del usuario $Nombre - $Error"
                Escribir-Log -log $log -Path $Path
                $Error.Clear()
            }

            $Licencias = $Usuario.Licenses

            foreach($Licencia in $Licencias){

                $Lic = $Licencia.AccountSkuId

                if(($Lic -eq "reseller-account:O365_BUSINESS_ESSENTIALS") -or ($Lic -eq "reseller-account:O365_BUSINESS_PREMIUM") -or ($Lic -eq "reseller-account:STANDARDPACK")){

                    #Guardamos el tamaño usado en el buzón
                    $Quota = $Statistics.TotalItemSize.Value.ToString().Split("(")[0]

                    #Guardamos el límite del buzón
                    $Limite = $QuotaLimite.ProhibitSendQuota.Split("(")[0]

                    #Troceamos la cadena en 2
                    $Quota = $Quota.split(" ")

                    #Troceamos la cadena en 2
                    $Limite = $Limite.split(" ")

                    #Si la unidad del tamaño usado es GB
                    if($Quota[1] -eq "GB"){

                        #Calculamos el tamaño disponible
                        $Disponible = $Limite[0] - $Quota[0]                

                        if($Disponible -lt 2){

                            $CuerpoEmail = "<p><h2>Buzon de usuario casi lleno</h2></p></br><p>El buzón del usuario $Nombre está casi lleno</p><p>Le quedan: $Disponible GB</p><p>Por favor contactar con el usuario</p>"

                            Enviar-Email -Desde $Desde -Para $Para -Asunto $Asunto -CuerpoEmail $CuerpoEmail -ServidorCorreo $ServidorCorreo -CredencialesBuzon $CredencialesBuzon

                            $log = "$fecha --> Al usuario $Nombre le quedan: $Disponible GB"
                            Escribir-Log -log $log -Path $Path

                        }
                    }
                }            
            }            
        }
    }

    Desconectar-O365

    Conectar-O365

    try{

        #Guardamos en un array todos los usuarios que estén licenciados
        $Usuarios = Get-MsolUser -MaxResults 1000 | where {$_.isLicensed -eq $true} | Select UserPrincipalName,Licenses
        Comprobar-Usuarios -Usuarios $Usuarios

    }catch{

        $log = "$fecha --> Se ha registrado un error guandando los parametros de los usuarios licenciados: $Error"
        Escribir-Log -log $log -Path $Path
        $Error.Clear()
    }    

    Desconectar-O365

