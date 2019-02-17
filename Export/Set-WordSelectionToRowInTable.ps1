<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordSelectionToRowInTable {

    # Select a Row in a given Table

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Row
    )

    # To Do: Check if the Range is valid

    $Table.Rows($Row).Select()
    
}