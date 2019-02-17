<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Get-WordTableColumnCount {

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table
    )

    $Table.Range.Columns.Count

}