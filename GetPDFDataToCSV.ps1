[Reflection.Assembly]::LoadFrom("D:\Scripts\itextsharp.dll") | Out-Null

#Store directory name
$folderName = pwd | Select-Object | %{$_.ProviderPath.Split("\")[-1]}

# Get Location of this path
$location = Get-Location

# Get Name File Path
$pdfPath = Get-ChildItem -Path $location -File -Recurse *.pdf | Select-Object -ExpandProperty FullName

# Get File Name 
$pdfFileName = (Get-ChildItem -Path $location *pdf -recurse).BaseName  


#Storage Data
$data = @()

#Storage Additional Information
$sum = 0
$num = 0

foreach ($pdfPath in $pdfPath){
    $pdfFiles = New-Object iTextSharp.text.pdf.PdfReader($pdfPath)
    $pdfCount = $pdfFiles.NumberOfPages
    $data += [PSCustomObject]@{
	  No = 0
        Nama_File = 0
        Jumlah = $pdfCount
        }
    $sum+= $pdfFiles.NumberOfPages
}



foreach ($pdfFileName in $pdfFileName){
    $data[$num].Nama_File = $pdfFileName
    $data[$num].No = $num + $pdfFileName.Count
    $num += $pdfFileName.Count
}

$data += [PSCustomObject]@{
        Nama_File = "Total"
        Jumlah = $sum
        }

$dataName = "$folderName.csv"

$data | Export-Csv $dataName -NoTypeInformation