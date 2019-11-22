$Usuarios = Get-ChildItem -Path \\192.168.0.50\perfiles -Directory
$domain = "AGQ\"

foreach($Usuario in $Usuarios){

    
    #$Directorios.Add($Usuario.Name)
    #$permissions = Get-Acl $Usuario.FullName
    
    if(($Usuario.Name.Split(".")[1]) -ne "V2"){
        
        $Nombre = $Usuario.Name.Split(".")[0]+"."+$Usuario.Name.Split(".")[1]
        $Ruta = $Usuario.FullName 
        $Nombre = $domain+$Nombre       
        <#[string]$Nombre = $domain+$Nombre        
        $permissions = Get-Acl $Ruta
        $userpermissions = New-Object System.Security.AccessControl.FileSystemAccessRule($Nombre,“FullControl”, “ContainerInherit, ObjectInherit”, “None”, “Allow”)
        $permissions.AddAccessRule($userpermissions)
        Set-Acl $Ruta $permissions
        #takeown /f  $Usuario.Fullname /R /U $Nombre
        #>
        .\icacls.exe "$Ruta" /setowner "$Nombre"
        .\icacls.exe "$Ruta" /reset /T
        Write-Host("Se concede permisos al usuario: $Nombre en la ruta: $Usuario.FullName")

    }else{

        $Nombre = $Usuario.Name.Split(".")[0]
        $Nombre = $domain+$Nombre 
        $Ruta = $Usuario.FullName
        .\icacls.exe "$Ruta" /setowner "$Nombre"
        .\icacls.exe "$Ruta" /reset /T
        <#        
        [string]$Nombre = $domain+$Nombre        
        $permissions = Get-Acl $Ruta
        $userpermissions = New-Object System.Security.AccessControl.FileSystemAccessRule($Nombre,“FullControl”, “ContainerInherit, ObjectInherit”, “None”, “Allow”)
        $permissions.AddAccessRule($userpermissions)
        Set-Acl $Ruta $permissions
        #>
        Write-Host("Se concede permisos al usuario: $Nombre en la ruta: $Usuario.FullName")
    }

}

