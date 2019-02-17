<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Set-WordDocumentTemplate {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path $_})]
        [String]
        $File
    )

    Write-Verbose "Setting Document Styles Template to $File"

    # https://docs.microsoft.com/en-us/office/vba/api/word.document.attachedtemplate
    $App.ActiveDocument.AttachedTemplate = $File

    Write-Verbose "Copying Styles from Template"

    # The original Code Sample says to use $App.ActiveDocument.AttachedTemplate.FullName()
    # but as this may return a HTTP URL if the Files are stored on OneDrive and the Option 
    # to use Office to update Documents is selected, and the Path is exactly the same, we use $File instead

    # https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CopyStylesFromTemplate
    $App.ActiveDocument.CopyStylesFromTemplate($File)

}