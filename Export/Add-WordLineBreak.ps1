<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordLineBreak {

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

        $Selection.TypeParagraph()

    }

}