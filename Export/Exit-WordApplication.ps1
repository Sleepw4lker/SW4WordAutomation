<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Exit-WordApplication {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App
    )

    process {

        If ($App.Application.Documents.Count -gt 0) {
            Close-WordDocument -App $App
        }

        Write-Verbose -Message "Exiting Word Application"

        $App.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($App)

    }

}