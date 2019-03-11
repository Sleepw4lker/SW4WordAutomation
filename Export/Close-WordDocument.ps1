<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Close-WordDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App
    )

    process {

        Write-Verbose -Message "Closing current Document"

        # Check version of Word installed and discard changes
        If ($(Get-WordVersion) -eq 14) {
            $App.ActiveDocument.Close([ref]$False)
        }
        Else {
            # Office 2013 or newer
            $App.ActiveDocument.Close($False)  
        }

    }

}