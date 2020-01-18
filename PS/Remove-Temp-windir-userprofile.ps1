#############################################
#
# Suppression de fichiers temporaires sur Windows 10
# avec arrêt/redémarrage du service Office
#
#############################################





#Verification et execution en tant qu'admin
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}
#Fin verification


clear-host
write-host "Script de suppression des fichiers temporaires"
write-host "##############################################"
write-host "##############################################"

$ServiceName = 'ClickToRunSvc'
$arrService = Get-Service -Name $ServiceName

while ($arrService.Status -ne 'Stopped')
{

    Stop-Service $ServiceName
    write-host "Le service $ServiceName s'arrete"
   Start-Sleep -seconds 2
   $arrService.Refresh()
   write-host $arrService.status
     
    if ($arrService.Status -eq 'Stopped')
    {
        write-host "Suppression des fichiers temps"
        remove-item $env:windir\temp\* -Recurse
        remove-item $env:USERPROFILE\AppData\local\Temp\* -Recurse
    }

}

write-host "Le service $ServiceName redémarre ..."
Start-Service $ServiceName
$arrService.Refresh()
Start-Sleep -seconds 2
write-host $arrService.Status

#Suppression fichiers temp dans C:\Windows
#remove-item $env:windir\temp\* -Recurse

#Suppression fichiers temp dans repertoire profil
#remove-item $env:USERPROFILE\AppData\local\Temp\* -Recurse

#Redemarrage service clicktorum
#Get-Service -name ClickToRunSvc | Start-Service

pause

