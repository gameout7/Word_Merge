#
#  Made by Anatolii Zhilin
#  From Engineer to Engineers
#
$inifile = ".\Word_List.ini"
$ini = Get-Content -Path $inifile
$Word_List = @()

$section = "none"  #section name in ini file
#Ini File parser
foreach ($line in $ini)
    {
    if ($line -ne "" -and $line.startswith(";") -ne $true )
        {
# Checking section name
            if ($line.StartsWith("[Word_list]") -eq $True) {$section = 'Word_list'; continue}
            
# Section Word_list                
            if ($section -eq 'Word_list')
            {
                $Word_list += $Line.Trim('')
            }
        }
    }

$WordPath = (get-location).path
#List of documents with pathes
$Word_List_Path = @()

for ($i = 0; $i -lt $Word_List.Count; $i++) {
    $Word_List_Path += $WordPath + "\Templates\" + $Word_List[$i]
}

#----------------------------------------------
#First document
$FirstDoc = $Word_List_Path[0]
#All other documents
$OtherDocs = $Word_List_Path[1..$($Word_List_Path.Length - 1)]

$resultDocument =  $WordPath + "\" + "Merged.docx"  # the output file

#--------------------------------------------------
$word = New-Object -ComObject Word.Application
$word.Visible = $true

$worddoc   = $word.Documents.Open($FirstDoc)
foreach ($doc in $OtherDocs)
    {
        $range = $worddoc.Range()
        # move to the end of the first document
        $range.Collapse(0)     # wdCollapseEnd see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcollapsedirection
        # insert a page break so the second document starts on a new page
        $range.InsertBreak(7)  # wdPageBreak   see https://learn.microsoft.com/en-us/office/vba/api/word.wdbreaktype
        $range.InsertFile($doc)

    }


$worddoc.SaveAs($resultDocument)
# quit Word and cleanup the used COM objects
$word.Quit()

$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worddoc)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
#-------------------------------------------