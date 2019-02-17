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

        $Selection = $App.Selection

        $Selection.TypeParagraph()

    }

}