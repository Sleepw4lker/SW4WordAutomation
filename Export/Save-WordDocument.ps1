<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Save-WordDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        # To-Do: Verify against allowed Extensions
        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path ((New-Object System.IO.FileInfo $_).Directory.FullName)})]
        [String]
        $File,

        [Parameter(Mandatory=$False)]
        [Switch]
        $EmbedFonts = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $AsPdf = $False
    )

    process {

        Write-Verbose -Message "Saving Document as $File"

        If ($EmbedFonts.IsPresent) {

            # https://docs.microsoft.com/en-us/office/vba/api/word.document.embedtruetypefonts
            $App.ActiveDocument.EmbedTrueTypeFonts = $True

            # https://docs.microsoft.com/en-us/office/vba/api/word.document.donotembedsystemfonts
            $App.ActiveDocument.DoNotEmbedSystemFonts = $True 

        }

        If ($AsPdf.IsPresent) {

            # https://docs.microsoft.com/en-us/office/vba/api/word.saveas2
            # See https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
            $App.ActiveDocument.SaveAs2(
                $File,
                [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
            )

        }
        Else {

            $App.ActiveDocument.SaveAs($File)

        }

    }

}