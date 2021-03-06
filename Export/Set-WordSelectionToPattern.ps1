<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordSelectionToPattern {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Pattern,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoWrap = $False
    )

    process {

        $Selection = $Doc.ActiveWindow.Selection

        $Selection.Find.ClearFormatting() 
        $Selection.Find.Forward = $True

        If ($NoWrap -eq $False) {
            $Selection.Find.Wrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue
        }
        Else {
            $Selection.Find.Wrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindStop
        }

        $Selection.Find.Text = $Pattern

        [void]$Selection.Find.Execute()

        $Selection.Find.Found

    }

}