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

function plDua{
param($num)
$num +2
}

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


Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$xl = New-Object -ComObject Excel.Application
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = $folderName


$pdfFileName | ForEach-Object {
    $data[$num].Nama_File = $_
    $data[$num].No = $num + $_.Count

    $ws.Cells.Item($num+2, 1) = $data[$num].No
    $ws.Cells.Item($num+2, 2) = $data[$num].Nama_File
    $ws.Cells.Item($num+2, 3) = $data[$num].Jumlah

    $ws.Cells.Item(1, 1) = "No"
    $ws.Cells.Item(1, 2) = "Nama File"
    $ws.Cells.Item(1, 3) = "Jumlah"

    $num += $_.Count
}


$data += [PSCustomObject]@{
        Nama_File = "Total"
        Jumlah = $sum
        }

$plusOne = $num + 1
$plusTwo = $num + 2

$ws.Cells.Item($plusTwo, 2) = $data[$num].Nama_File
$ws.Cells.Item($num+2, 3).Formula = "=sum(C2:C$plusOne)"

$range = "A1:C$plusTwo"
$ws.Range($range).Borders.LineStyle = 1

$ws.Columns.EntireColumn.AutoFit()

$fileName = "$location\$folderName.xlsx"
$wb.SaveAs($fileName)
$xl.Quit()




    