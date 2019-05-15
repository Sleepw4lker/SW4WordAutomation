<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordTableCaption {

    # Set Style of a Table or Individual Row

    [cmdletbinding()]
    Param (
        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [Object]
        $Table,

        # To Do: Parameter Validation
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Caption,

        # To Do: Parameter Validation
        [Parameter(Mandatory=$False)]
        [Switch]
        $Above = $False
    )

    process {

        Write-Verbose "Adding Caption $Caption"

        # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcaptionposition?view=word-pia
        If ($Above.IsPresent) {
            $Position = [Microsoft.Office.Interop.Word.WdCaptionPosition]::wdCaptionPositionAbove
        }
        Else {
            $Position = [Microsoft.Office.Interop.Word.WdCaptionPosition]::wdCaptionPositionBelow
        }

        # https://msdn.microsoft.com/en-us/vba/word-vba/articles/selection-insertcaption-method-word
        $Table.Range.InsertCaption(
            # https://docs.microsoft.com/en-us/office/vba/api/word.wdcaptionlabelid
            [Microsoft.Office.Interop.Word.WdCaptionLabelID]::wdCaptionTable,
            ": $Caption",
            $False, 
            $Position
        )

    }

}