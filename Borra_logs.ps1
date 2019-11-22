    <#
.SYNOPSIS
    Con este script creamos un histórico de los logs y posteriormente los borramos

.DESCRIPTION
    Este script gestiona los logs que crea el script de asignacion de licencias

.NOTES

    File Name      : Borra_logs.ps1
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
    
        $Archivo1            = "C:\O365\Logs\LOG_Asignar-Licencias.txt"
        $Archivo2            = "C:\O365\Logs\Historicos\LOG_Asignar-Licencias.txt"
        $Path                = "C:\O365\Logs\"
        $Path2               = "C:\O365\Logs\Historicos\"
        $fecha               = Get-Date -Format g
        $ArchivoHistorico    = $fecha.ToString().Replace(" ","_").ToString().Replace("/","-").ToString().Insert(16,".txt").ToString().Insert(0,"LOG_").ToString().Replace(":","-")

        $ExiteArchivo = Get-ChildItem -Path $Archivo1

        #Comprobamos que el archivo Archivo1 Existe
        if($ExiteArchivo -ne $null){
    
            Copy-Item –Path $Archivo1 –Destination $Archivo2 -Confirm:$false -Force                    
            Rename-Item -Path $Archivo2 -NewName $ArchivoHistorico -Confirm:$false -Force
            Remove-Item -Path $Archivo1 -Confirm:$false -Force  
            
         }
         
            
           



   
        
        


        
        


  