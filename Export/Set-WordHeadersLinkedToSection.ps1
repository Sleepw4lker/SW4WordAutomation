<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordHeadersLinkedToSection {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
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