###########################################
#                                         #
# Miramos si los módulos están instalados #
#                                         #
###########################################

    $SQLModuleCheck = Get-Module -ListAvailable ADSync
    $MSOnlineCheck = Get-Module -ListAvailable MsOnline

    if ($SQLModuleCheck -eq $null)
    {
        
        # Sino está instalado seleccionamos el repositorio y cambiamos la políticade instalación
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

        # Instalamos el Módulo SqlServer
        Install-Module -Name ADSync –Scope AllUsers -Confirm:$false -AllowClobber
    }


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
    
    Import-Module ADSync 
    Import-Module MsOnline

    $cred = Import-Clixml -Path C:\O365\key\credenciales.xml


    #Conectarse a Exchange Online
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
    Import-PSSession $exchangeSession -DisableNameChecking
    Connect-MsolService -Credential $cred

    Start-ADSyncSyncCycle -PolicyType Delta
    Start-ADSyncSyncCycle -PolicyType Initial

    Get-PSSession | Remove-PSSession