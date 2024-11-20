Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
Add-Type -AssemblyName System.Windows.Forms
function CorrectSpelling_word {
    param (
        [string[]]$Proofread_texts,  # Accepts an array of strings (lines of text)
        [switch]$tostring
    )
try{
$Word = New-Object -COM Word.Application
$Word.Visible = $false  # Set to true for debugging if you want to see Word open
$Document = $Word.Documents.Add()
$Textrange = $Document.Range(0)
$Textrange.LanguageID = [Microsoft.Office.Interop.Word.WdLanguageID]::wdEnglishUS  # Set language to English (US)
$newwords=@()
$excludes=get-content C:\Users\106094-DUTN\Desktop\Auto\plume_frv\iText\exclude.txt
$newdict=get-content C:\Users\106094-DUTN\Desktop\Auto\plume_frv\iText\newwords.txt

  if($Proofread_texts.gettype().name -eq "String"){
    $Proofread_texts = $Proofread_texts.split("`n")|  Where-Object {($_|out-string).Trim() -gt 0}
    }

foreach($Proofread_text in $Proofread_texts){
    $Proofread_text=$Proofread_text.replace(",","，").replace("^","")
    $combinded = $false
    $splitwords = $Proofread_text.Split(" ").split("-").split(".").split("`n").split("(").split(")").split("[").split("]") | Where-Object {($_|out-string).Trim() -gt 0}  # Split words and remove extra spaces
    $i = $splitwords.Count -1
    while ($i -lt $splitwords.Count -and $i -ge 0 ) {     
        $wd = ($splitwords[$i]|Out-String).trim()
           #$Textrange.Text = $wd
                   # Skip if the word contains any digits
           if(($i -eq $splitwords.Count -1 -and $wd  -match '\d') -or $wd -in $excludes )   {
                    $i--
                    continue
                }
          
         try{ 
         $wd2=$($splitwords[$i - 1]|Out-String).trim()
         }
         catch{
         $wd2=$false
         }
         try{ 
         $wd3=$($splitwords[$i - 2]|Out-String).trim()
         }
         catch{
         $wd3=$false
         }

        if ($wd3 -and $i + 2 -le $splitwords.Count -and  !($wd2 -match '\d')  -and !($wd3 -match '\d') -and ($wd3 -notin $excludes) ) {
            
           $nextWord ="$wd3$($wd2)$($wd)"
           $oldword = "$wd3 $($wd2) $($wd)"
           $pattern="$($wd3)\s$($wd2)\s$($wd)"  
           
           $Textrange.Text = $nextWord
           #Start-Sleep -Milliseconds 200  # Delay for Word to process the text
            if (($Textrange.SpellingErrors.Count -eq 0 -or $nextWord -in $newdict) -and $nextWord -notin $excludes  -and $Proofread_text -match $pattern) {
                $combinded=$true
                $Proofread_text = $Proofread_text.Replace("$oldword", "$nextWord")
                $i -= 2
            }
        }
        # Check if there's a next word to combine with
        if ($wd2 -and $i + 1 -le $splitwords.Count -and !$combinded -and !($wd2 -match '\d') -and ($wd2 -notin $excludes)) {
            $nextWord = "$($wd2)$($wd)"
            $oldword = "$($wd2) $($wd)"
            $pattern = "$($wd2)\s$($wd)"        
            $Textrange.Text =$nextWord
            #Start-Sleep -Milliseconds 200             
            if (($Textrange.SpellingErrors.Count -eq 0 -or $nextWord -in $newdict ) -and $nextWord -notin $excludes -and $Proofread_text -match $pattern ) {
                $combinded=$true
                $Proofread_text = $Proofread_text.Replace("$oldword", "$nextWord")
                $i--
            }
        }
        $combinded = $false
        $i--  # Move to the next word
    }
    $newwords+=@($Proofread_text)
}

if ($tostring) {
$newwordlines = $newwords -join "`n"
return $newwordlines
}
else {
    return $newwords
}

} finally {  # Cleanup Word COM objects
    if ($Textrange) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Textrange) | Out-Null
    }
    if ($Document) {
        $Document.Close([ref]0) # 0 means "don't save"
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
    }
    if ($Word) {
        $Word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null
    }
    $Textrange = $null
    $Document = $null
    $Word = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
 }
}

if($PSScriptRoot.length -eq 0){
    $scriptRoot="C:\Users\106094-DUTN\Desktop\Auto\iText\"
    }
    else{
    $scriptRoot=$PSScriptRoot
    }
        
    $zipname=(Get-ChildItem $scriptRoot |Where-Object{$_.name -eq "modules.zip"}).FullName
    #$zipname
    $unzippath="$scriptRoot\modules"
    #$unzippath
    if (!(Test-Path $unzippath) ){        
     Expand-Archive $zipname -DestinationPath $unzippath -Force
    }
    $assemblies=Get-ChildItem -path $scriptRoot -r |Where-Object{$_.name -match ".dll"}
    foreach($assembly in $assemblies){
        Unblock-File $assembly.fullname 
    try{
        add-type -Path $assembly.fullname
        #Write-Host "PASS:"
        #$assembly.fullname 

    }
    catch{
        #Write-Host "FAIL:"
        #$assembly.fullname
    }
    }

$mainpath="C:\tmp\pdftest\"
$pdfs=Get-ChildItem -path $mainpath -filter *frv*
$anafolder= $mainpath  |Join-Path -ChildPath "ana"

if (!(test-path $anafolder)){
    new-item -ItemType Directory -path $anafolder|Out-Null
}
foreach($pdf in $pdfs){
  
  $pdfPath=$pdf.FullName
  $anaPath=($pdf.basename -split "frv")[1]
  $anafolder=$mainpath|Join-Path -ChildPath "ana/$($anaPath)"
  if(!(test-path  $anafolder)){
    new-item -ItemType Directory -Path $anafolder|out-null
  }
  $csvpath=$anafolder |Join-Path -ChildPath "casetable$($anaPath).csv"
  #$csvpath2=$anafolder |Join-Path -ChildPath "casetable$($anaPath)_keys.csv"
  $menutxt=$anafolder |Join-Path -ChildPath "pdfmenu$($anaPath).csv"
  $contenttxt=$anafolder |Join-Path -ChildPath "pdfcontent$($anaPath).txt"

$reader= New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $pdfPath
$menutext = @()
$contents = @()
# Loop through each page of the PDF
for ($page =  1; $page -le $reader.NumberOfPages; $page++) {
    # Extract the text from the page
    #$strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
    $strategy = New-Object iTextSharp.text.pdf.parser.LocationTextExtractionStrategy
    $currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy)
    $lines=($currentText.split("`n") | Select-Object -skip  1) 
    $line1=($currentText.split("`n") | Select-Object -first  1) 

    #$line2=($currentText.split("`n") | Select-Object -Skip 1) 
    if (($line1.replace(" ","")) -match "documentation") {
        $menutext+=@($lines)
    }
    else{
       $contents+=@($lines)
    }
   
}

# Close the PdfReader
$reader.Close()

# Output the text

new-item $menutxt -Force|Out-Null
$newlines=CorrectSpelling_word -Proofread_texts $menutext -tostring
add-content $menutxt -value $newlines
new-item $contenttxt -Force|Out-Null
add-content $contenttxt -value $contents

$sections2=New-Object System.Collections.ArrayList
$section=$sections=$caseids=$pylines=$noaddcaseids=@()
$caseid =  $null
$i=0
$j=0
$k=999
$m=9999
$q=9999
$r=9999
$caseidlines=@()
$pylinestep=@()
$expectline=@()
$preconline=@()
$n=0
foreach ($line in $contents) {
   
        $perc=[math]::Round($i/($contents.Count)*100,0)
        Write-Progress -Activity "$($pdfPath) pdf reading" -Status "$($perc)% Complete:" -PercentComplete $perc

    $i++
    $j++
    $k++
    $n++
    $m++
    $q++
    $r++
    if($i -lt $contents.count){
     $nextline=(($contents[$i]).replace(" ","")|Out-String).trim()
    }
    $linetrim=($line.replace(" ","")|Out-String).trim()
    #if((($linetrim -eq "CaseID" -and $nextline -match "^C\d+") -or ($line -in $line1s -and $nextline -match "^C\d+") -or $i -eq ($contents.count))){
   
     if($j -eq 1 -and $i -ne 1){
                $caseid=$line.replace(" ","")
                $caseids+=@($caseid)
     }
    else{

        if($linetrim -match "^precondition"){
            $r=0
            $m=9999
            $q=9999
        } 
        if ($linetrim -eq "steps"){
            $r=9999
        }

       if($linetrim -match "^Automation"){
            $k=0
            $r=9999
            $m=9999
            $q=9999
        } 

        if($linetrim -eq "Expectedresult" -or $linetrim.trim() -eq "Expectedresult"){
            $r=9999
            $m=9999
            $q=0
            $expectline+=@($stepid)
        } 
        if(($linetrim -match "^\d+\.$" -or $linetrim -match "^\d+\.\d+\.$" ) -and ($nextline -eq "step") -and $r -ge 9999 ){
            $m=0
            $q=9999
            $r=9999
            $stepid=$linetrim
        }
        if($linetrim -match "^\d+\.\d+\.\d+"){
            $k=999
            $q=9999
        }
        if($k -gt 0 -and $k -lt 5){
           $pylines+=$line
           $pylines=$pylines.trim()
        }
        if($m -ge 0 -and $m -ne 1 -and $m -lt 9999 ){
            $pylinestep+=@($line.replace("^","")) 
         }
         if($q -gt 0 -and $q -lt 9999){
            $expectline+=@($line)
            if($q -gt 50){
            $expectline+=@("...(The output is too long, please refer document)")
            $q = 9999
            }
         }
         if($r -gt 0 -and $r -lt 9999){
            $preconline+=@($line)
         }
        $section+=@($line)
    }

         
    if(($linetrim -eq "CaseID" -and ($nextline -match "^\d{6,}\b" -and !($nextline -match "\."))) -or ($i -eq $contents.count) ){
        $caseidlines+=@($n)
        $j=0
        $q=9999
        #$caseid
        #$section
            if(!$pylines -or $pylines.length -eq 0){
            $pylinesnew="manual"
            }
            else{
            if($pylines -match "pytest tests"){
            $pylinesnew="pytest tests"+ ($pylines -split "pytest tests")[1] 
            }
            elseif($pylines -match "tests\/" -and !($pylines -match "pytest tests")){
                $pylinesnew="pytest tests/"+ ($pylines -split "tests/")[1] 
            }
            else{
            $pylinesnew=$pylines|out-string
            }
            }
            try{
             $pylinestepnew=CorrectSpelling_word -Proofread_texts $pylinestep -tostring
              }
              catch{
                write-host "check $($caseid): `n $pylinestep"
                 }
               try{
                $sectionObject = [PSCustomObject]@{
                    CaseID = $caseid.trim()
                    #Contents  = $section|Out-String
                    PyLines  = $pylinesnew.trim()
                    precondition=$preconline|out-string
                    steps=$pylinestepnew
                    expects=$expectline|out-string
                   }                
                [void]$sections2.Add($sectionObject)
                $sections+=@($section)
                $pylinestep=@()
                $expectline=@()
                $preconline=@()
                $section = $null
                $pylines=$null
                $pylinesnew=$null
               }
                catch{
                  if($caseid){
                    write-host "$caseid logs not added"
                    write-host "$_.Exception.Message"
                    $noaddcaseids+=@($caseid)
                    }            

                }
        
       }
}
$sections2|export-csv $csvpath -NoTypeInformation -Encoding UTF8

$csvdata=import-csv $csvpath
foreach($csvline in $csvdata){
    $csvline."expects"=($csvline."expects").replace(",","，").replace("?","").replace("^^","")
    $csvline."precondition"=($csvline."precondition").replace(",","，").replace("?","").replace("^^","")
    $csvline."steps"=($csvline."steps").replace(",","，").replace("?","").replace("^^","")
}
$csvdata|export-csv $csvpath -NoTypeInformation -Encoding UTF8
<#
#collect keywords
$casecontent=import-csv $csvpath
$keywords=get-content C:\tmp\plume\Plumewords.txt

foreach($caseadd in $casecontent){
    foreach($keyword in $keywords){
        $caseadd | Add-Member -MemberType NoteProperty -Name $keyword -Value ""
    }
}

$casecontent|export-csv $csvpath2 -NoTypeInformation


$newcsvcontent=import-csv $csvpath2
$caseids=$newcsvcontent."Caseid"
$newcsv=@()
foreach($line1 in $caseidlines){
    if($line1 -notin $line1s){
     $nextindex=$caseidlines.IndexOf($line1)+1
    try{
        $contentcase=$contents[$line1..($caseidlines[$nextindex]-2)]
    }
    catch{
        $contentcase=$contents[$line1..($contents.count -1)]
    }
    $checkcaseid=$contentcase[0]
    $csvchecks=$newcsvcontent|Where-Object{($_."CaseID") -eq  $checkcaseid}
    if($csvchecks){
    foreach($line2 in $contentcase){
        $line3=$line2.trim().replace(" ","")
        foreach($keyword in $keywords){
        if($line3 -like "*$keyword*"){
           [int32]($csvchecks.$keyword) +=1
           <#
           if ($keyword -eq "bridge"){
            $line2
            start-sleep -s 10
           }
          }
        }
      if($line3 -like "*automationexecution*"){
        break
      }
    }
    #write-host "check  $checkcaseid ok"
    $newcsv+=$csvchecks
   }
 }

}
$newcsv|export-csv $csvpath2 -NoTypeInformation
#>
 Write-Progress -Activity "$($pdfPath) PDF read" -Completed
}

$outputCsv= "$($mainpath)ana\casetable_all.csv"
$csvFiles = Get-ChildItem -Path "$($mainpath)ana\"  -r -Filter "casetable*.csv"

$data = @()

foreach ($csvFile in $csvFiles) {
    $currentData = Import-Csv -Path $csvFile.FullName -Encoding UTF8    
    $data += $currentData
}
$uniqueRows = @()
$finalUniqueData = @()

foreach ($row in $data) {
    $caseid=$row.CaseID
    #$rowString = ($row.PSObject.Properties.Value -join "|")
    if (-not $uniqueRows.Contains($caseid)) {
        $uniqueRows+=@($caseid)
        $finalUniqueData += $row
    }
}

$finalUniqueData | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

