<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Set-WordDocumentTitle {

    [cmdletbinding()]
    Param (
        [Parameter(
            Mandatory=$True,
            ParameterSetName="CallByApp"
        )]
        [Alias("WordApp")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(
            Mandatory=$True,
            ParameterSetName="CallByDoc"
        )]
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

    # Assuming that the Function was called via the $App Parameter,
    # we take the currently active Document as the Document to process
    If (-not $Doc) {
        $Doc = $App.ActiveDocument
    }

    $Doc.BuiltInDocumentProperties("Title") = $Title

}