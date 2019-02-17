<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Search-WordPatternAndReplaceInDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Pattern,

        # Null or Empty allowed
        [Parameter(Mandatory=$False)]
        [String]
        $ReplaceWith,

        [Parameter(Mandatory=$False)]
        [Switch]
        $IncludeHeaders = $False
    )

    process {

        # Prohibit Function failure when an empty String is passed
        If ((![String]::IsNullOrEmpty($Pattern)) -and (![String]::IsNullOrEmpty($ReplaceWith))) {

            $Selection = $App.Selection
            Search-WordPatternAndReplaceInSelection `
                -Selection $Selection `
                -Pattern $Pattern `
                -ReplaceWith $ReplaceWith

            If ($IncludeHeaders -eq $True) {

                $App.ActiveDocument.Sections | ForEach-Object {

                    $_.Headers | ForEach-Object {
                        Search-WordPatternAndReplaceInSelection `
                            -Selection $_.Range `
                            -Pattern $Pattern `
                            -ReplaceWith $ReplaceWith
                    }

                    $_.Footers | ForEach-Object  {
                        Search-WordPatternAndReplaceInSelection `
                            -Selection $_.Range `
                            -Pattern $Pattern `
                            -Replacewith $ReplaceWith
                    }

                }

            }

        }

    }

}