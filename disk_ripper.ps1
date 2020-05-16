Import-Module -Name PoSHue

$HUE_API_KEY = <YOUR HUE API KEY>
$HUE_IP = <YOUR HUE IP>
$HUE_ROOM = <THE ROOM WHERE THE LIGHTS WILL FLASH>

$MIN_TITLE_LENGTH = 900
$OUTPUT_DIR = <YOUR OUTPUT DIRECTORY>
$DISK_DRIVE = <CD_ROM_DRIVE LETTER>
$TITLE = "all"
$PROFILE = ".\default.mmcp.xml"

function EjectDisk {
    $drives = Get-WmiObject Win32_Volume -Filter "DriveType=5"
        if ($null -eq $drives) {
            Write-Warning "Your computer has no CD drives to eject."
                return
        }
    $drives | ForEach-Object {
        (New-Object -ComObject Shell.Application).Namespace(17).ParseName($_.Name).InvokeVerb("Eject")
    }
}

function WaitForDiskLoad {
    Write-Output "Waiting for disk load"
        do {
            $cdrom = Get-WMIObject Win32_CDROMDrive
                if ($cdrom.MediaLoaded -eq "True") {
                    Write-Output "Media loaded"
                        return
                }
            Start-Sleep -Seconds 1
        } While ($true)

}

function FlashLights {
    param (
            [HueGroup]$LightGroup
          )
        Write-Output "Lights!"
        $OriginalX = $LightGroup.XY.x
        $OriginalY = $LightGroup.XY.y
        $OrignalBright = $LightGroup.Brightness

        for ($i = 0; $i -lt 4; $i++) {
            $LightGroup.SetHueGroup(150, 0.1, 0.1)
                Start-Sleep -Milliseconds 500
                $LightGroup.SetHueGroup(150, 0.65, 0.31)
                Start-Sleep -Milliseconds 500

        }

    $LightGroup.SetHueGroup($OrignalBright, $OriginalX, $OriginalY)
}

Write-Output "Starting"
$ActiveGroup = [HueGroup]::New($HUE_ROOM, $HUE_IP, $HUE_API_KEY)


if ((Get-WMIObject -Class Win32_CDROMDrive -Property *).MediaLoaded -eq $False) {
    Write-Output "Please insert disk..."
        EjectDisk
}

do {
    WaitForDiskLoad
        New-Item -ItemType Directory $OUTPUT_DIR\tmp
        makemkvcon --noscan --messages=-stdout --progress=-stderr --debug=$OUTPUT_DIR\Debug\debug.log --profile=$PROFILE `
        --minlength=$MIN_TITLE_LENGTH mkv dev:$DISK_DRIVE $TITLE $OUTPUT_DIR\tmp

        $movieName = (Get-ChildItem $OUTPUT_DIR\tmp)[0] | Select-Object -Property BaseName

        $duplicateCounter = 0
        while (Test-Path $OUTPUT_DIR\$movieName) {
            $movieName = $movieName + $duplicateCounter++
        }
    Rename-Item $OUTPUT_DIR\tmp $OUTPUT_DIR\$movieName
        Write-Output "Disk ripped to: " + $OUTPUT_DIR + $movieName

        FlashLights($ActiveGroup)
        Write-Output "Done"
        EjectDisk
} while($true)
