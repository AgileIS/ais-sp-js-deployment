param(
    [Parameter(Mandatory=$true)][String] $File,
	[Parameter(Mandatory=$true)][String] $ConnectionTemplate
)

function PrepareConnectionFile([Parameter(Mandatory=$true)]$Template, [Parameter(Mandatory=$true)]$Connections) {
    if($Connections.Length -gt 0){
        Write-Output "- preparing connection file for $($File)"
        $connectionFile = Get-Content $Template
        foreach($con in $Connections){
            $connectionFile = $connectionFile | % { $_ -replace "__$($con.ListName)__View__" , $con.ViewId -replace "__$($con.ListName)__List__" , $con.ListId -replace  "__$($con.ListName)__RootFolderURL__" , $con.ListRootFolderUrl -replace "__$($con.ListName)__WebUrl__" , $con.WebUrl }
           
        }
        Set-Content "$($File).connection.xml" -Value $connectionFile -Force
    }
}

function UpdateXlsxConnectionFile([Parameter(Mandatory=$true)]$File){
    $windowsBase = [Reflection.Assembly]::LoadWithPartialName("WindowsBase");
    Write-Output "- update xlsx file connection in file $($File)"
    $xlsxFile = [System.IO.Packaging.Package]::Open($(Resolve-Path $File))
    $newConnectionFileInfo = Get-Item "$($File).connection.xml"
    if($xlsxFile.PartExists("/xl/connections.xml") -and $newConnectionFileInfo) {
         $connectionItem = $xlsxFile.GetPart("/xl/connections.xml")
         $newConnectionFileInfo = Get-Item "$($File).connection.xml"
         $newConnectionFileStream = $newConnectionFileInfo.Open([System.IO.FileMode]::Open)
         $newConnectionFileStream.CopyTo($connectionItem.GetStream())
         $newConnectionFileStream.Close()
         $newConnectionFileStream.Dispose()
    }
    $xlsxFile.Close()
    $xlsxFile.Dispose();
}

Write-Output "updating file connection $($file)"
$cons = (Get-Content "$($file).json") -join "`n" | ConvertFrom-Json
if($cons) {
    PrepareConnectionFile -Template $ConnectionTemplate -Connections $cons
    UpdateXlsxConnectionFile -File $File
    Remove-Item "$($File).connection.xml"
    Write-Output "updated file connection $($file)"
} else {
    Write-Error "- cannot find connection file for $($file)"
}

