# Set-Wallpaper function borrowed from here https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/
Function Set-WallPaper {
 
<#
 
    .SYNOPSIS
    Applies a specified wallpaper to the current user's desktop
    
    .PARAMETER Image
    Provide the exact path to the image
 
    .PARAMETER Style
    Provide wallpaper style (Example: Fill, Fit, Stretch, Tile, Center, or Span)
  
    .EXAMPLE
    Set-WallPaper -Image "C:\Wallpaper\Default.jpg"
    Set-WallPaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
  
#>
 
param (
    [parameter(Mandatory=$True)]
    # Provide path to image
    [string]$Image,
    # Provide wallpaper style that you would like applied
    [parameter(Mandatory=$False)]
    [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
    [string]$Style
)
 
$WallpaperStyle = Switch ($Style) {
  
    "Fill" {"10"}
    "Fit" {"6"}
    "Stretch" {"2"}
    "Tile" {"0"}
    "Center" {"0"}
    "Span" {"22"}
  
}
 
If($Style -eq "Tile") {
 
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force
 
}
Else {
 
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force
 
}
 
Add-Type -TypeDefinition @" 
using System; 
using System.Runtime.InteropServices;
  
public class Params
{ 
    [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
    public static extern int SystemParametersInfo (Int32 uAction, 
                                                   Int32 uParam, 
                                                   String lpvParam, 
                                                   Int32 fuWinIni);
}
"@ 
  
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02
  
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
  
    $ret = [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $fWinIni)
}

$Date = Get-Date -Format "yyyyMMdd"
$ImgPath = "C:\Users\" + $env:UserName + "\Pictures\BingImageOfTheDay_" + $Date
if ($ImgPath + ".*" | Test-Path) {
    Write-Host "Already got today's image"
} else {
    Write-Host "Getting new Bing image of the day"
    
    # Grab image location from the API
    $BingUrlBase = "https://www.bing.com"
    $ImgDataUrl = $BingUrlBase + "/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-GB"
    $ImgDataResponse = Invoke-RestMethod -URI $ImgDataUrl
    $ParsedURI = [System.Uri]($BingUrlBase + $ImgDataResponse.images[0].url)

    # Strip out unecessary qs params and get the file extension
    $Query = [System.Web.HttpUtility]::ParseQueryString($ParsedURI.Query)
    $ImgId = $Query.Get('id')
    $ImageUrl = $BingUrlBase + $ParsedURI.AbsolutePath + '?id=' + $ImgId
    $Ext = $ImageUrl.split('.')[-1]
    $FinalFilename = $ImgPath + "." + $Ext

    Invoke-WebRequest -Uri $ImageUrl -OutFile $FinalFilename
    Set-WallPaper -Image $FinalFilename -Style "Fill"
}