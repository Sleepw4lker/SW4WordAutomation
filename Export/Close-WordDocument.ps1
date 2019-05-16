<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Close-WordDocument {

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