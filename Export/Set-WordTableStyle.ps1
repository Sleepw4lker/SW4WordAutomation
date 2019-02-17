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

        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [String]
        $Style
    )

    # To Do: Implement a Validation if the Style actually exists

    $Table.Style = $Style

}