<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Remove-WordSelection {

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

        $Selection = $Doc.ActiveWindow.Selection

        # https://docs.microsoft.com/en-us/office/vba/api/Word.Selection.Delete
        [void]$Selection.Delete()

        # https://docs.microsoft.com/de-de/office/vba/api/word.selection.typebackspace
        $Selection.TypeBackSpace()

    }

}