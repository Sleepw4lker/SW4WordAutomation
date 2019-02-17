<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordSelectionToCellInTable {

    # Jump to a Cell, identified by Row, Col ID

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Row,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,63)]
        [int]
        $Column
    )

    # To Do: Check if the Range is valid

    $Table.Cell($Row, $Column).Range.Select()

}