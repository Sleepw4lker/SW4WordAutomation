<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Update-WordDocumentFields {

    # You must pass a "Word.Application" Object      

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App
    )

    process {

        # Fields before ToC, otherwise the ToC will not honor the Lists of Figures and Tables correctly
        Write-Verbose -Message "Updating Document Fields"

        [void]$App.ActiveDocument.Fields.Update()

        Write-Verbose -Message "Updating Table(s) of Contents"

        # https://docs.microsoft.com/en-us/office/vba/api/word.tableofcontents
        $App.ActiveDocument.TablesOfContents | ForEach-Object {

            [void]$_.Update()
            
        }

    }

}