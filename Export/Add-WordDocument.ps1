<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordDocument {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$True)]
        [Alias("Path")]
        [ValidateScript({Test-Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Above = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Below = $False
    )

    process {

        $Selection = $Doc.ActiveWindow.Selection

        If ($Above) {
            Set-WordSelectionToTopOfDocument -Doc $Doc
        }

        If ($Below) {
            Set-WordSelectionToBottomOfDocument -Doc $Doc
        }

        <#
        # https://docs.microsoft.com/en-us/office/vba/api/word.wdinformation
        # Cursor Position greater 1 means that this is not an empty Line
        If ($Selection.Information([Microsoft.Office.Interop.Word.WdInformation]::wdFirstCharacterColumnNumber) -gt 1) {
            # https://docs.microsoft.com/en-us/office/vba/api/word.selection.typeparagraph
            $Selection.TypeParagraph()
        }
        #>

        Write-Verbose -Message "Inserting $File"

        # Append the Document to the Base Document
        # https://technet.microsoft.com/en-us/library/ee692877.aspx
        # https://docs.microsoft.com/en-us/office/vba/api/word.selection.insertfile
        $Selection.InsertFile($File)

    }

}