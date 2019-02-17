<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function New-WordTable {

    # Create a Table, return a Table Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Rows,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,63)]
        [int]
        $Columns,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Caption
    )

    <#
        To implement:
        - Default Border Style
        - Default Width Rule
        - Default Horizontal Alignment
        - Default Vertical Alignment
    #>

    Write-Verbose "Creating a Table of $Rows x $Columns Dimension."

    # https://msdn.microsoft.com/en-us/vba/word-vba/articles/tables-add-method-word
    $Table = $App.ActiveDocument.Tables.Add(
        $App.Selection.Range,
        $Rows,
        $Columns,
        # https://docs.microsoft.com/en-us/office/vba/api/word.wddefaulttablebehavior
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdautofitbehavior
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    )

    If ($Caption) {
        # Add Caption to Table
        # https://msdn.microsoft.com/en-us/vba/word-vba/articles/selection-insertcaption-method-word
        $Table.Range.InsertCaption(
            # https://docs.microsoft.com/en-us/office/vba/api/word.wdcaptionlabelid
            [Microsoft.Office.Interop.Word.WdCaptionLabelID]::wdCaptionTable,
            ": $Caption",
            $False, 
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcaptionposition?view=word-pia
            [Microsoft.Office.Interop.Word.WdCaptionPosition]::wdCaptionPositionBelow
        )
    }

    $Table

}