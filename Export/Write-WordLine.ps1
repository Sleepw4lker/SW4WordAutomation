<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Write-WordLine {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(
            Position = 0,
            Mandatory = $True, 
            ValuefromPipeline = $True
        )]
        [String[]]
        $Line,

        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Font = $Null,

        [Parameter(Mandatory=$False)]
        [ValidateScript({
            Test-WordIsValidStyle -Doc $Doc -Style $_
        })]
        [String]
        $Style,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,72)]
        [int]
        $Size = 0,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,16)]
        [int]
        $Indent = 0,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Underline = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Italic = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Bold = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Bullet = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoNewLine = $False,

        [Parameter(Mandatory=$False)]
        [Switch]
        $Upward = $False
    )

    begin {

        $Selection = $Doc.ActiveWindow.Selection

        If (![String]::IsNullOrEmpty($Font)) {
            $Selection.Font.Name = $Font
        } 

        If ($Size -ne 0) {
            $Selection.Font.Size = $Size
        }

        $Selection.Font.Italic = $Italic
        $Selection.Font.Bold = $Bold
        $Selection.Font.Underline = $Underline

        If ($Style) {
            
            # https://docs.microsoft.com/en-us/office/vba/api/word.style
            $OldStyle = $Selection.Range.Style.NameLocal

            Try {
                $NewStyle = $Doc.Styles($Style)
            }
            Catch {

            }
    
            If ($NewStyle) {
                $Selection.Range.Style = $NewStyle
            } 
            
        }

        # https://msdn.microsoft.com/en-us/VBA/Word-VBA/articles/range-orientation-property-word
        # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdtextorientation?view=word-pia
        If ($Upward) {
            $Selection.Range.Orientation = [Microsoft.Office.Interop.Word.WdTextOrientation]::wdTextOrientationUpward
        }
        Else {
            $Selection.Range.Orientation = [Microsoft.Office.Interop.Word.WdTextOrientation]::wdTextOrientationHorizontal
        }

        If ($Bullet) {
            
            # https://docs.microsoft.com/en-us/office/vba/api/word.listformat.applybulletdefault
            # For compatibility reasons, the default constant is wdWord8ListBehavior , but in new procedures 
            # you should use wdWord9ListBehavior to take advantage of improved Web-oriented formatting with 
            # respect to indenting and multilevel lists.
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wddefaultlistbehavior?view=word-pia
            # Use formatting compatible with Microsoft Word 2002.
            $Selection.Range.ListFormat.ApplyBulletDefault(
                [Microsoft.Office.Interop.Word.WdDefaultListBehavior]::wdWord10ListBehavior
            )

        }

    }

    process {

        $Line | ForEach-Object {
            $Selection.TypeText($_)
            $Selection.TypeParagraph()
        }

    }

    end {

        If ($Indent -ne 0) {

            For ($i = 1; $i -le $Indent; $i++) {
                # https://docs.microsoft.com/en-us/office/vba/api/word.paragraph.indent
                $Selection.Paragraphs(1).Indent()
            }

        }

        If ($NoNewLine) {
            $Selection.TypeBackspace()
        }

        If ($Indent -ne 0) {

            For ($i = 1; $i -le $Indent; $i++) {
                # https://docs.microsoft.com/en-us/office/vba/api/word.paragraph.indent
                $Selection.Paragraphs(1).Outdent()
            }

        }

        If ($Bullet) {

            # ApplyBulletDefault() is just exactly like clicking the bullet 
            # icon so you have to call it to turn it on and turn it off.
            $Selection.Range.ListFormat.ApplyBulletDefault(
                [Microsoft.Office.Interop.Word.WdDefaultListBehavior]::wdWord10ListBehavior
            )

        }

        If ($NewStyle) {

            $Selection.Range.Style = $OldStyle

        }

        
    }

}