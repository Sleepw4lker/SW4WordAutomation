<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordPageBreak {

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

        # https://docs.microsoft.com/en-us/office/vba/api/word.selection.insertbreak
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdbreaktype
        $Selection.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

    }

}