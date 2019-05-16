<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Search-WordPatternAndReplaceInDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

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

        $Selection = $Doc.ActiveWindow.Selection

        # Prohibit Function failure when an empty String is passed
        If ((-not [String]::IsNullOrEmpty($Pattern)) -and (-not [String]::IsNullOrEmpty($ReplaceWith))) {

            Search-WordPatternAndReplaceInSelection `
                -Selection $Selection `
                -Pattern $Pattern `
                -ReplaceWith $ReplaceWith

            If ($IncludeHeaders -eq $True) {

                $Doc.Sections | ForEach-Object {

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