<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordTableStyle {

    # Set Style of a Table or Individual Row

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Style
    )

    process {

        # To Do: Implement a Validation if the Style actually exists
        # https://docs.microsoft.com/en-us/office/vba/api/word.table.style
        # https://docs.microsoft.com/en-us/office/vba/api/word.tablestyle
        # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdstyletype
        # Paragraph, Table ans other Styles are all referenced via the same Style Property
        # Thus, if applying a Table Style to a Table and then a Font Style, the first will be overridden by the first
        Write-Verbose "Setting Table Style to $Style"
        Try {
            $Table.Style = $Style
        }
        Catch {
            
        }

    }

}