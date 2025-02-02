Function Export-EmlToWord {
    param(
        $SourceDirectory,
        $OutputDirectory="c:\temp\"

    )
    BEGIN {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $emlFiles = Get-ChildItem -Path $SourceDirectory -Filter *.eml
    }
    
    PROCESS {
        foreach ($emlFile in $emlFiles) {
            write-host "Processing -- $($emlFile.BaseName)"
            $OutputDirectory=$OutputDirectory.TrimEnd("\")
            $docPath = "$OutputDirectory\$($emlFile.BaseName).docx"

            #the functino below is from https://github.com/PsCustomObject/PowerShell-Functions/blob/master/Convert-EmlFile.ps1 and needs to be loaded prior to running this script
            $email=Convert-EmlFile $emlFile.FullName
            $imageFiles=@{}
            foreach ($bodyPart in $email.BodyPart.BodyParts) {
                if ($bodyPart.FileName -ne "") {
                    $bodyPartFile="$($env:TEMP)\$($bodyPart.FileName)"
                    $bodyPart.SaveToFile("$bodyPartFile")
                    $imageFiles.add($bodyPart.FileName,$bodyPartFile)
                }
            }

            #create new tmp file & write html body to temp file
            $tmpFile=New-TemporaryFile
            $email.HTMLBody | out-file $tmpFile -Encoding utf8 -Force

            #instantiate document range and insert the html file content
            $doc = $word.Documents.Add()
            $range=$doc.Range()
            $range.Insertfile($tmpFile)
            
            #fix embedded images
            foreach ($inlineShape in $doc.InlineShapes) {
                $inlineShapeFileName=$inlineShape.LinkFormat.SourceName
                if ($inlineShapeFileName) {
                    $inlineShapeFileName=$inlineShapeFileName.Substring(0,$inlineShapeFileName.IndexOf("@"))
                    write-host "Processing ---- $inlineShapeFileName"
                    #if inlineshape we want, selected it
                    $sourceFile=$imageFiles[$inlineShapeFileName]
                    $inlineShape.Select()
                    $word.Selection.InlineShapes.AddPicture("$sourceFile",$false,$true)
                }
            }

            write-host "Saving -- $docPath"
            $doc.SaveAs([ref] $docPath)
            $doc.Close()

            #cleanup temp files
            Remove-Item $tmpFile
            foreach ($imageFile in $imageFiles.Values) {
                Remove-Item $imageFile
            }
        }
    }
    END {
        
        #Close Word
        $word.Quit()
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    }
    
}