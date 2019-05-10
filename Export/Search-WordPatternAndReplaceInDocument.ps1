<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Search-WordPatternAndReplaceInDocument {

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

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

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