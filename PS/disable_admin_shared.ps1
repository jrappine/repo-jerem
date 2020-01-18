##################################################
#                                                #
# Désactiver le partage administratif Windows 10 #
#                                                #
##################################################


##### Stopper et désactiver le service ####

get-service -Name lanmanserver | stop-service -Force

##### Désactiver le service au démarrage #####

set-service -Name lanmanserver -StartupType Disabled


##### Désactiver dans le registre #####

$RegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"

new-itemproperty -Path $RegistryPath -Name "AutoShareWks" -Property "Dword" -Value 0