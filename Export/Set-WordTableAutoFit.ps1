<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordTableAutoFit {

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table
    )

    $Table.Columns.AutoFit()

}