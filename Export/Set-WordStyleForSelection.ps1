<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordStyleForSelection {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Style = $Null
    )

    process {

        $Selection = $Doc.ActiveWindow.Selection

        $NewStyle = $Doc.Styles($Style)
        
        If ($NewStyle) {
            $Selection.Range.Style = $NewStyle
        } 

    }

}