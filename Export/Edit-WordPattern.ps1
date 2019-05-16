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
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

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
        Set-WordSelectionToTopOfDocument -Doc $Doc

        Do {

            $Found = Set-WordSelectionToPattern `
                -Doc $Doc `
                -Pattern $Pattern `
                -NoWrap

            If ($Found -eq $True) {

                $Selection = $Doc.ActiveWindow.Selection

                $Selection.Font.Italic = $Italic
                $Selection.Font.Bold = $Bold
                $Selection.Font.Underline = $Underline

            }

        } While ($Found -eq $True)
 
    }

}