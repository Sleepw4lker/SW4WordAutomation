<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordHeadersLinkedToSection {

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

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [Int]
        $Section      
    )

    process {

        Write-Verbose "Linking all Headers to the One in Section $Section"

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

        $Doc.Sections | ForEach-Object {

            $SectionIndex++

            $_.Headers | ForEach-Object {

                If ($SectionIndex -gt $Section) {

                    $_.LinktoPrevious = $True

                }

            }

        }

    }

}