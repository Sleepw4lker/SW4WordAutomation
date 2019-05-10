<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordLineBreak {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App
    )

    process {

        # Can this be changed to $Doc.Selection so that we can call either by $App or $Doc?
        $Selection = $App.Selection

        $Selection.TypeParagraph()

    }

}