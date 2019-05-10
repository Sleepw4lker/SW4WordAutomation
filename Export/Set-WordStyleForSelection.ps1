<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordStyleForSelection {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Style = $Null
    )

    process {

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

        $Selection = $Doc.ActiveWindow.Selection

        $NewStyle = $Doc.Styles($Style)
        
        If ($NewStyle) {
            $Selection.Range.Style = $NewStyle
        } 

    }

}