<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Remove-WordSelection {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc
    )

    process {

        $Selection = $Doc.ActiveWindow.Selection

        # https://docs.microsoft.com/en-us/office/vba/api/Word.Selection.Delete
        [void]$Selection.Delete()

        # https://docs.microsoft.com/de-de/office/vba/api/word.selection.typebackspace
        $Selection.TypeBackSpace()

    }

}