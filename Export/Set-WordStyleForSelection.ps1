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

    $Selection = $App.Selection

    $NewStyle = $App.ActiveDocument.Styles | Where-Object { $_.NameLocal -eq $Style }
    
    If ($NewStyle) {
        $Selection.Range.Style = $NewStyle
    } 

}