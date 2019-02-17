<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Get-WordTableRowCount {

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table
    )

    $Table.Range.Rows.Count
    
}