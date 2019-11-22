<#
.SYNOPSIS
    Con este script asignamos de forma automática las licencias disponibles en O365

.DESCRIPTION
    Este Script usa la conexión a una base de datos para ver que usuarios se han dado de alta en Besafer y se ha creado el usuario en Active Directory
    para después asignarle una licencia de correo o quitarselas en caso que sea una baja.


.NOTES

    File Name      : Asignar_def.ps1
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

    $SQLModuleCheck = Get-Module -ListAvailable SqlServer
    $MSOnlineCheck = Get-Module -ListAvailable MsOnline

    if ($SQLModuleCheck -eq $null)
    {
        
        # Sino está instalado seleccionamos el repositorio y cambiamos la políticade instalación
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

        # Instalamos el Módulo SqlServer
        Install-Module -Name SqlServer –Scope AllUsers -Confirm:$false -AllowClobber
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
    
    Import-Module SqlServer 
    Import-Module MsOnline

###########################################################
#                                                         #
# Definimos las variables de conexión a la base de datos  #
#                                                         #
###########################################################
    
    #Establecemos el servidor de base de datos
    $SQLInstance = "192.168.0.28"

    #Seleccionamos la Base de datos
    $SQLDatabase = "O365"

    #Guradamos la tabla Usuarios
    $SQLTable    = "Usuarios"

    #Guardamos la tabla Países
    $SQLTable2   = "Paises"

    #Guardamos la tabla Licencias
    $SQLTable3   = "Licencias"

    #Guardamos la tabla Historicos
    $SQLTable4   = "Hist_Usuarios"
    #Credenciales de la Base de datos
    $CredencialesSQL = Import-Clixml -Path C:\O365\key\credenciales_SQL.xml
    
#############################
#                           #
# Definimos otras variables #
#                           #
#############################

    #Si hay algún error no controlado se para el Script
    $ErrorActionPreference = "Stop"    
         
    #Credenciales del buzón
    $CredencialesBuzon= Import-Clixml -Path C:\O365\key\credenciales_buzon.xml   
 
    #Configuramos el correo
    $Desde="noreply@agqlabs.com"
    $Para="notificacionesit@agqlabs.com"
    $Asunto="Faltan Licencias"    
    $ServidorCorreo="smtp.office365.com"

    #Establecemos el Valor a 2 de la columna App de la BBDD Hist-Usuarios
    $App = 2
    
    #Establecemos el valos a 0 de la columna de IdPersonal de la BBDD Hist-Usuarios
    $IdPersonal = 0   

    #Guardamos la fecha
    $fecha = Get-Date -Format g  
    
    #Establecemos la ruta de escritura del log de las licencias
    $Path = "C:\O365\Logs\LOG_Asignar-Licencias.txt"

    #Variable que usamos para mantener un control sobre el comportamiento del Script
    $Control = 0
    
    #Variable en la que establecemos si existe el usuario o no (0 --> exite, 1 --> No existe)
    $UsuarioNoExiste = $false

    $Error.Clear()


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

    #Funcion para ejecutar comandos SQL
    function Ejecutar-SQL{

        param($SQLQuery,$SQLInstance,$CredencialesSQL)

        try{
            
            if(($SQLQuery -ne $null) -and ($SQLInstance -ne $null) -and ($CredencialesSQL -ne $null)){

                Invoke-Sqlcmd -query $SQLQuery -ServerInstance $SQLInstance -Credential $CredencialesSQL
            }

        }
        catch [System.Management.Automation.ValidationMetadataException]{
            
           if($log -ne $null){

                $log = "$fecha --> Error al iniciar sesión en la BBDD"
                Escribir-Log -log $log -Path $Path
                $Error.Clear()
           }
        }
    }

    #Funcion para ejecutal las query de SQl
    function Ejecutar-Query{

     param(     
         [String] $Operacion,     
         [String] $Column,     
         [String] $Valor)
    
    switch($Operacion){

        "UPDATE"{

            switch($Column){

                "Sync"{

                    $Query = "UPDATE $($SQLTable) SET Sync = $Valor WHERE Nombre = '$Nombre'"

                    try{
                    
                        Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                    }catch{

                        $log = "$fecha --> No se ha podido actualizar el valor Sync: $Error"
                        Escribir-Log -log $log -Path $Path
                        $Error.Clear()

                    }

                    break

                }

                "Notificacion"{

                    $Query = "UPDATE $($SQLTable) SET Notificacion = $Valor WHERE Nombre = '$Nombre'"

                    try{

                        Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                    }catch{

                        $log = "$fecha --> No se ha podido actualizar el valor Notificación: $Error"
                        Escribir-Log -log $log -Path $Path
                        $Error.Clear()

                    }
                    break

                }

                "FechaSync"{

                    $Query = "UPDATE $($SQLTable) SET FechaSync = GETDATE() WHERE Nombre = '$Nombre'"

                    try{

                        Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                    }catch{

                        $log = "$fecha --> No se ha podido actualizar el valor FechaSync: $Error"
                        Escribir-Log -log $log -Path $Path
                        $Error.Clear()

                    }
                    break
                }
            }        

            break
        }

        "INSERT"{

            switch($Valor){

                "alta"{

                    $Query = "INSERT INTO $($SQLTable4) (FechaHistorico,Id,App,IdPersonal,Operacion,Alta) VALUES(GETDATE(),'$Id','$App','$IdPersonal',4,'$Alta')"

                    try{

                        Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                    }catch{

                        $log = "$fecha --> No se ha podido insertar el valor de alta en la BBDD: $Error"
                        Escribir-Log -log $log -Path $Path
                        $Error.Clear()

                    }
                    break

                }

                "baja"{

                    $Query = "INSERT INTO $($SQLTable4) (FechaHistorico,Id,App,IdPersonal,Operacion,Alta) VALUES(GETDATE(),'$Id','$App','$IdPersonal',5,'$Alta')"

                    try{

                        Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                    }catch{

                        $log = "$fecha --> No se ha podido insertar el valos de baja en la BBDD: $Error"
                        Escribir-Log -log $log -Path $Path
                        $Error.Clear()

                    }
                    break

                }

                "update"{

                    $Query = "INSERT INTO $($SQLTable4) (FechaHistorico,Id,App,IdPersonal,Operacion,Alta) VALUES(GETDATE(),'$Id','$App','$IdPersonal',6,'$Alta')"

                    try{

                        Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                    }catch{

                        $log = "$fecha --> No se ha podido insertar el valor de update en la BBDD: $Error"
                        Escribir-Log -log $log -Path $Path
                        $Error.Clear()

                    }

                    break
                }
            }

            break
        }
    }
   }

    #Funcion Principal del Script
    function Asignar-Licencias{

        param($Registros)
        
        #Comprobamos que haya registros que sincronizar
        if($Registros -ne $null){
            
            #Hacemos un recorrido por todos los registros guardados en Registros
            foreach($Registro in $Registros){
                
                #Asignamos el valor falso a la variable de control
                $Control = $false

                #Variable que utilizamos para ver si un usuario existe (0 --> El usuario exite, 1 --> el usuario NO existe)
                $UsuarioNoExiste = $false

                #Guardamos el nombre del usuario
                $Nombre = $Registro.Nombre

                #Guardamos el ID del Pais
                $IdPais = $Registro.IdPais

                #Guardamos el ID de la Licencia
                $IdLicencia = $Registro.IdLicencia

                #Guardamos la OU
                $OU = $Registro.NombreOU

                #Guardamos el ID
                $Id = $Registro.Id

                #Guardamos el estado del alta
                $Alta = $Registro.Alta
                
                try{

                    #Guardamos las propiedades del usuario
                    $Sync = Get-MsolUser -UserPrincipalName $Nombre

                }catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException]{
                    
                    #Guardamos el error en el fichero de logs
                    $log = "$fecha --> No se ha podido encontrar las propiedades del usuario $Nombre - $Error"
                    Escribir-Log -log $log -Path $Path
                    
                    #Indicamos en el log que el usuario no existe
                    $log = "$fecha --> El usuario $Nombre no existe en Office 365"
                    Escribir-Log -log $log -Path $Path

                    #Ponemos el valor Sync a 1 en la Base de datos
                    Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor 1

                    #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                    Ejecutar-Query -Operacion "INSERT" -Valor "update"

                    #Establecemos la variable a 1 para indicar que el usuario NO Existe
                    $UsuarioNoExiste = $true

                    $Error.Clear()

                 }

                 switch ($Alta){

                    'True'{

                        #El usuario se tiene que dar de ALTA
                        $log = "$fecha --> procederemos a dar de alta al usuario: $Nombre"
                        Escribir-Log -log $log -Path $Path

                        switch($UsuarioNoExiste){

                            'True'{

                               #El usuario $Nombre NO existe en Office 365
                               $log = "$fecha --> El usuario $Nombre no existe en Office 365. Esperamos hasta la siguiente asignación"
                               Escribir-Log -log $log -Path $Path

                               #Ponemos el valor Sync a 0 en la Base de datos
                               Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "0"

                               #Ponemos el valor de notificacion a 0
                               Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "0"

                               #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                               Ejecutar-Query -Operacion "INSERT" -Valor "update"

                               break

                             }

                            'False'{

                               #El usuario $Nombre SI existe en Office 365
                               switch($Sync.IsLicensed){

                                    'True'{
                                        
                                        #El usuario YA está licenciado
                                        $log = "$fecha --> El usuario $Nombre ya está licenciado, no hacemos ningún cambio"
                                        Escribir-Log -log $log -Path $Path

                                        #Ponemos el valor Sync a 1 en la Base de datos
                                        Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "1"

                                        #Ponemos el valor de notificacion a 1
                                        Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "1"

                                        #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                        Ejecutar-Query -Operacion "INSERT" -Valor "update"

                                        break

                                    }

                                    'False'{

                                        #El usuario NO está licenciado

                                        #Buscamos el país según el código del registro
                                        $Query = "USE $SQLDatabase
                                                SELECT * FROM $SQLTable2 WHERE IdPais = '$IdPais'"
                                        
                                        $Pais = Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL

                                        #Seleccionamos el tipo de licencia O365
                                        $Query = "USE $SQLDatabase
                                                SELECT * FROM $SQLTable3 WHERE IdLicencia = $IdLicencia"

                                        #Gaurdamos el tipo de licencia en la variable
                                        $AccountSkuId = Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL       

                                        #Comprobamos el numero de licencias del tipo AccountSkuId
                                        $Licencias = Get-MsolAccountSku | Where {$_.AccountSkuId -eq $AccountSkuId.Licencia}

                                        $Disponibles = $Licencias.ActiveUnits - $Licencias.ConsumedUnits

                                        if($Disponibles -gt 0){

                                            try{

                                                #Asignamos la localización por defecto en España
                                                Set-MsolUser -UserPrincipalName $Nombre -UsageLocation $Pais.CodAlfa2

                                                #Asignamos la licencia al correo seleccionado
                                                Set-MsolUserLicense -UserPrincipalName $Nombre -AddLicenses $AccountSkuId.Licencia.ToString()                                               

                                            }catch{

                                                $log = "$fecha --> Se ha registrado un error intentando asignar la licencia: $Error"
                                                Escribir-Log -log $log -Path $Path
                                                $Error.Clear()

                                                #Ponemos el valor de notificacion a 0
                                                Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "0"

                                                #Ponemos el valor de Sync a 1
                                                Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "0"
                                                               

                                                #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                                Ejecutar-Query -Operacion "INSERT" -Valor "update"

                                            }                                            

                                            #Ponemos el valor Sync a 1 en la Base de datos
                                            Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "1"

                                            #Ponemos el valor de notificacion a 1
                                            Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "1"

                                            #Escribimos la fecha de sincronizacion en la base de datos
                                            Ejecutar-Query -Operacion "UPDATE" -Column "FechaSync"

                                            #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                            Ejecutar-Query -Operacion "INSERT" -Valor "alta"
                        
                                            #Guardamos el tipo de licencia como cadena de texto
                                            $Lic = $AccountSkuId.Licencia.ToString()

                                            $log = "$fecha --> Se ha asignado la licencia $Lic al usuario $Nombre"
                                            Escribir-Log -log $log -Path $Path

                                            break

                                        }
                                        else{

                                            #Guardamos el tipo de licencia
                                            $Lic = $AccountSkuId.Licencia

                                            #Establecemos en el cuerpo email el tipo de licencia que falta
                                            $CuerpoEmail="<body><h2>Se ha intentado asignar una licencia</h2><p>Faltan licencias del tipo $Lic al usuario $Nombre</p></body>"

                                            #Enviamos el email
                                            Enviar-Email -CredencialesBuzon $CredencialesBuzon -Desde $Desde -Para $Para -Asunto $Asunto -CuerpoEmail $CuerpoEmail -ServidorCorreo $ServidorCorreo
        
                                            #Ponemos el valor de notificacion a 1
                                            Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "1"                  
                            
                
                                            #Escribimos la fecha de sincronizacion en la base de datos de Usuarios
                                            Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "0"

                                            #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                            Ejecutar-Query -Operacion "INSERT" -Valor "update"

                                        }
                                    }
                               }
                            }
                        }
                           
                     break
                    }
                    
                    'False'{
                        
                        #El usuario se tiene que dar de BAJA
                        $log = "$fecha --> procederemos a dar de baja al usuario: $Nombre"
                        Escribir-Log -log $log -Path $Path

                        switch($UsuarioNoExiste){

                            'True'{

                                #El usuario NO existe en Office 365
                                $log = "$fecha --> El usuario $Nombre no existe en Office 365. No tenemos que hacer ningún cambio"
                                Escribir-Log -log $log -Path $Path

                                #Ponemos el valor Sync a 1 en la Base de datos
                                Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "1"

                                #Ponemos el valor de notificacion a 1
                                Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "1"

                                #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                Ejecutar-Query -Operacion "INSERT" -Valor "update"

                                break

                            }

                            'False'{

                                #El usuario SI existe en Office 365
                                $log = "$fecha --> El usuario $Nombre si existe en Office 365. procedemos a darlo de baja"
                                Escribir-Log -log $log -Path $Path

                                switch($Sync.IsLicensed){

                                    'True'{

                                        #El usuario TIENE licencias asignadas
                                                #Guardamos todas las licencie tiene asignadas el usuario
                                        $LicenciasO365 = Get-MsolUser -UserPrincipalName $Nombre | Select Licenses

                                        #Por cada licencia que tenga se la quitamos.
                                        foreach($Licencia in $LicenciasO365){

                                            try{

                                                #Quitamos las licencias que tenga el usuario
                                                Set-MsolUserLicense -UserPrincipalName $Nombre -RemoveLicenses $Licencia.Licenses.AccountSkuId

                                                $Lic = $Licencia.Licenses.AccountSkuId.ToString()

                                                $log = "$fecha --> Se ha quitado la licencia $Lic al usuario $Nombre"
                                                Escribir-Log -log $log -Path $Path

                                                #Escribimos la fecha de sincronizacion en la base de datos
                                                Ejecutar-Query -Operacion "UPDATE" -Column "FechaSync"

                                                #Ponemos el valor de notificacion a 1
                                                Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "1"

                                                #Ponemos el valor de Sync a 1
                                                Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "1"
                                                               

                                                #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                                Ejecutar-Query -Operacion "INSERT" -Valor "baja"

                                       
                                            }catch{

                                                $log = "$fecha --> Se ha registrado un error intentado quitar las licencias: $Error"
                                                Escribir-Log -log $log -Path $Path
                                                $Error.Clear()

                                                #Ponemos el valor de notificacion a 0
                                                Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "0"

                                                #Ponemos el valor de Sync a 0
                                                Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "0"
                                                               

                                                #Escribimos la fecha de sincronizacion en la base de datos de Historico de cambios
                                                Ejecutar-Query -Operacion "INSERT" -Valor "update"

                                            }
   
                                            break
                                        }
                                    }

                                    'False'{

                                        #El usuario NO TIENE licencias asignadas
                                        $log = "$fecha --> El usuario $Nombre no tiene ninguna licencia asignada, por lo que no hacemos nada con él"
                                        Escribir-Log -log $log -Path $Path

                                        #Escribimos la fecha de sincronizacion en la base de datos
                                        Ejecutar-Query -Operacion "UPDATE" -Column "FechaSync"

                                        #Ponemos el valor de notificacion a 1
                                        Ejecutar-Query -Operacion "UPDATE" -Column "Notificacion" -Valor "1"

                                        #Ponemos el valor de Sync a 1
                                        Ejecutar-Query -Operacion "UPDATE" -Column "Sync" -Valor "1"

                                        break

                                    }
                                }
                            }
                        }
                    }
                    
                  }                
            }

        }else{

            $log = "$fecha --> No hay usuarios para sincronizar"
            Escribir-Log -log $log -Path $Path

        }
    }

#######################################
#                                     #
# Funcionamiento Principal del Script #
#                                     #
#######################################

    #Seleccionamos los usuarios que no han sincronizado
    $Query = "USE $SQLDatabase
              SELECT * FROM $SQLTable WHERE Sync = 0"
    
    #Guardamos los registros de la base de datos que tienen el valos Sync = 0
    $QueryOutPut = Ejecutar-SQL -SQLQuery $Query -SQLInstance $SQLInstance -CredencialesSQL $CredencialesSQL
  
    Desconectar-O365

    Conectar-O365
  
    Asignar-Licencias -Registros $QueryOutPut

    Desconectar-O365