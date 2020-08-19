# Doc-PDF
Converts a Doc to a PDF and moves them to designated folders


# 08/19/2020
#
# SarahJanieC
#
#
#
# Purpose:  Converts the DOC into a PDF and sends it to outbound folder.
#           Finally, the program removes DOC files from inbound folder


#Set Folder Paths
$docPath = [Environment]::GetFolderPath("Desktop") + '\Folder Here'
$releasePath = [Environment]::GetFolderPath("Desktop") + '\Folder Here'


# Convert DOC to PDF
function PDFcreate()
{
        # Load Applications
        $word = New-Object -ComObject Word.Application
        $word.visible = $false

        ### Gets the most recent doc files in the folder specified by $docPath
        $docs = Get-ChildItem -Path $docPath | 
                    Where-Object {$_.Extension -like ".doc*" } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -first 10

        foreach($doc in $docs)
        {
            $convert = $word.Documents.Open($doc.Fullname)
            $pdf = ($convert.Fullname).replace("docx","pdf")
            $convert.SaveAs($pdf,17)
        }

        $word.quit()
        $word = $null
        [GC]::Collect()
        Stop-Process -Name "winword"
}

# Moves PDFs to new folder
function PDFmove()
{
        ### Gets the most recent pdf files in the folder specified by $docPath
        $pdfdocs = Get-ChildItem -Path $docPath | 
                    Where-Object {$_.Extension -like ".pdf" } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -first 10

         foreach($pdfdoc in $pdfdocs)
         {
            Move-Item -Path $pdfdoc.fullname -Destination $releasePath
         }
}

# Deletes DOCs from folder after converting them to PDFs
function Cleanse()
{
         foreach($doc in $docs)
         {
            Remove-Item -Path $doc.FullName
         }

}
