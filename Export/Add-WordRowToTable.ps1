<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordRowToTable {

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table
    )

    [void]$Table.Rows.Add()

}