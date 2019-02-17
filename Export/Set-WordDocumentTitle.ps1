<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Set-WordDocumentTitle {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Title
    )

    Write-Verbose "Setting Document Title to $Title"

    $App.ActiveDocument.BuiltInDocumentProperties("Title") = $Title

}