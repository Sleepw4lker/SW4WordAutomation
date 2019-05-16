<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Set-WordDocumentTitle {

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
        $Title
    )

    Write-Verbose "Setting Document Title to $Title"

    $Doc.BuiltInDocumentProperties("Title") = $Title

}