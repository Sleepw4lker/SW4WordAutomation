<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Update-WordDocumentFields {

    # You must pass a "Word.Application" Object      

    [cmdletbinding()]
    Param (
        [Parameter(
            Mandatory=$True,
            ParameterSetName="CallByApp"
        )]
        [Alias("WordApp")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(
            Mandatory=$True,
            ParameterSetName="CallByDoc"
        )]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc
    )

    process {

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

        # Fields before ToC, otherwise the ToC will not honor the Lists of Figures and Tables correctly
        Write-Verbose -Message "Updating Document Fields"

        [void]$Doc.Fields.Update()

        Write-Verbose -Message "Updating Table(s) of Contents"

        # https://docs.microsoft.com/en-us/office/vba/api/word.tablesofcontents
        $Doc.TablesOfContents | ForEach-Object {

            # https://docs.microsoft.com/en-us/office/vba/api/word.tableofcontents
            [void]$_.Update()
            
        }

    }

}