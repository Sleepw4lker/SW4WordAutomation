<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Edit-WordPattern {

    # You must pass a "Word.Application" Object

    # ToDo: Include Text Markers (Yellow and so on)

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Pattern,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Underline = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Italic = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Bold = $False
    )

    process {

        Write-Verbose -Message "Editing Pattern ""$Pattern"""

        # We must search without wrapping to avoid an endless loop
        Set-WordSelectionToTopOfDocument -App $App

        Do {

            $Found = Set-WordSelectionToPattern -App $App -Pattern $Pattern -NoWrap

            If ($Found -eq $True) {

                $Selection = $App.Selection

                $Selection.Font.Italic = $Italic
                $Selection.Font.Bold = $Bold
                $Selection.Font.Underline = $Underline

            }

        } While ($Found -eq $True)
 
    }

}