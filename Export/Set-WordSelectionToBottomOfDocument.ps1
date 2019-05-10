<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Set-WordSelectionToBottomOfDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(
            Mandatory=$True,
            ParameterSetName="CallByApp"
        )]
        [Alias("WordApp")]
        [Alias("Application")]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(
            Mandatory=$True,
            ParameterSetName="CallByDoc"
        )]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc
    )

    process {

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

        $Selection = $Doc.ActiveWindow.Selection

        # https://technet.microsoft.com/en-us/library/ee692877.aspx
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdmovementtype
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdunits

        # This method returns an integer that indicates the number of characters the selection 
        # or active end was actually moved, or it returns 0 (zero) if the move was unsuccessful. 
        # This method corresponds to functionality of the END key.
        [void]$Selection.EndKey(
            [Microsoft.Office.Interop.Word.WdUnits]::wdStory,
            [Microsoft.Office.Interop.Word.WdMovementType]::wdMove
        )

    }

}