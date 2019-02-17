<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordSelectionToTable {

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table
    )

    $Table.Select()

}