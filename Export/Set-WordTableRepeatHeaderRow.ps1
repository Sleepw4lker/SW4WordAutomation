<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordTableRepeatHeaderRow {

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table
    )

    # Repeat as Header Row
    # https://docs.microsoft.com/en-us/office/vba/api/word.row.headingformat
    $Table.Rows(1).HeadingFormat = $True
}