<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Above = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Below = $False
    )

    process {

        $Selection = $App.Selection

        If ($Above) {
            Set-WordSelectionToTopOfDocument -App $App
        }

        If ($Below) {
            Set-WordSelectionToBottomOfDocument -App $App
        }

        Write-Verbose -Message "Inserting $File"

        # Append the Document to the Base Document
        # See https://technet.microsoft.com/en-us/library/ee692877.aspx
        $Selection.InsertFile($File)

    }

}