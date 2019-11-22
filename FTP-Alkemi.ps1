<#
.SYNOPSIS
    Con este Script se buscan ficheros que se dejan en un FTP y se manipulan según unas condiciones dadas

.DESCRIPTION
    Con este Script se buscan ficheros que se dejan en un FTP y se manipulan según unas condiciones dadas


.NOTES

    File Name      : FTP_Alkemi.ps1
    Author         : Efrain Daza (efrain.daza@agqlabs.com)
    Prerequisite   : PowerShell V5
    Copyright 2019 - Efrain Daza/Agq Labs.
    Version        : 1.0.0
#>

###########################
#                         #
# Definimos las variables #
#                         #
###########################

  $FechaActual = Get-Date -Format g
  

  #Guardamos el directorio donde se encontrarán los archivos
  $Directorio = Get-Item -Path \\192.168.100.243\Alkemi\alkemi\ClientesXML

  #Guardamos todos los ficheros existentes
  $Ficheros = $Directorio | Get-ChildItem -File

  [System.Collections.ArrayList]$FicherosaEnviar= @()

  #Establecemos los Path con los que trabajaremos para mover los archivos
  $PathOK = "\\192.168.100.243\Alkemi\alkemi\ClientesXML\OK\"
  $PathKO = "\\192.168.100.243\Alkemi\alkemi\ClientesXML\KO\"

  $Path = "C:\O365\Logs\LOG_FTP-Alkemi.txt"  

  $EstadoOk = "OK"
  $EstadoKo = "NOK"
  
  #Si hay algún error no controlado se para el Script
  $ErrorActionPreference = "Stop"    
         
  #Credenciales del buzón
  $CredencialesBuzon= Import-Clixml -Path C:\O365\key\credenciales_buzon.xml   
 
  #Configuramos el correo
  $Desde="noreply@agqlabs.com"
  $Para="notificacionesit@agqlabs.com"
  $Asunto="Fallo de ficheros en FTP Alkemi"
  $Mensaje = "Los ficheros adjuntos dieron error y se han movido a $PathKO"   
  $ServidorCorreo="smtp.office365.com"
  
  #Limpiamos los posibles errores que hubiera en la pila
  $Error.Clear()

###########################
#                         #
# Definimos las funciones #
#                         #
###########################

    #Funcion Para escribir en los logs
    function Escribir-Log{

        param($log,$Path)

        if(($log -ne $null) -and ($Path -ne $null)){

            $log>>$Path

        }
    }
    
    #Función para enviar emails
    function Enviar-Email
    {   
        param($Desde,$Para,$Asunto,$CuerpoEmail,$Fichero,$ServidorCorreo,$CredencialesBuzon) 
        Try
        {   
         
            if($Para -ne $null)
            {
                [string[]]$To=$Para.Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
        
                Send-MailMessage -From $Desde -To $To -Subject $Asunto -Body $CuerpoEmail -BodyAsHtml -Attachments $Fichero -SmtpServer $ServidorCorreo -Credential $CredencialesBuzon -UseSsl -Port 587
            }    
        }
        catch [System.Exception]
        {
               $log = Write-Host $_.Exception.ToString()            
               Escribir-Log -log $log -Path $Path        
            
        } 
    }

    Function Comprobar-Ficheros{

        foreach($Fichero in $Ficheros){

           $NombreFichero = $Fichero.Name
           
           $EstadoFichero = $NombreFichero.split(" ")

           if($EstadoFichero[0] -eq $EstadoOk){
               
               Copy-Item -Path $Fichero.FullName -Destination $PathOK -Force
               $log = "$FechaActual --> Se ha encontrado el fichero: $NombreFichero y se ha movido a la carpeta: OK"
               Escribir-Log -log $log -Path $Path
               
               #Remove-Item -Path $Fichero.FullName -Force

           }           
           elseif($EstadoFichero[0] -eq $EstadoKo){
            
               Copy-Item -Path $Fichero.FullName -Destination $PathKO
               $log = "$FechaActual --> Se ha encontrado el fichero co errores: $NombreFichero y se ha movido a la carpeta: KO"
               Escribir-Log -log $log -Path $Path

               $FicheroTemp = $PathKO+$Fichero.Name
               $FicherosaEnviar.Add($FicheroTemp)
               
               #Remove-Item -Path $Fichero.FullName -Force

           }else{

                $log = "$FechaActual --> No se han encontrado ficheros para procesar"
                Escribir-Log -log $log -Path $Path

           }
        }
    }    
   
   Comprobar-Ficheros
   
   Enviar-Email -Desde $Desde -Para $Para -Asunto $Asunto -CuerpoEmail $Mensaje -Fichero $FicherosaEnviar -ServidorCorreo $ServidorCorreo -CredencialesBuzon $CredencialesBuzon



