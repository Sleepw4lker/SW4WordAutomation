<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Close-WordDocument {

    # You must pass a "Word.Application" Object

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
        $Doc
    )

    process {

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

        Write-Verbose -Message "Closing current Document"

        # Check version of Word installed and discard changes
        If ($(Get-WordVersion) -eq 14) {
            $Doc.Close([ref]$False)
        }
        Else {
            # Office 2013 or newer
            $Doc.Close($False)  
        }

    }

}