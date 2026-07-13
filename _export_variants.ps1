$ErrorActionPreference = "Stop"
$dir = "C:\Users\user\Documents\ai\osg-prod\ppt\_variants"
$ppt = New-Object -ComObject PowerPoint.Application
Get-ChildItem -Path $dir -Filter "s05_v*.pptx" | ForEach-Object {
    $src = $_.FullName
    $png = [System.IO.Path]::ChangeExtension($src, ".png")
    try {
        $pres = $ppt.Presentations.Open($src, $true, $false, $false)
        $pres.Slides.Item(1).Export($png, "PNG", 1600, 900)
        $pres.Close()
        Write-Output ("exported " + $_.Name)
    } catch { Write-Output ("FAILED " + $_.Name + " : " + $_.Exception.Message) }
}
$ppt.Quit()
