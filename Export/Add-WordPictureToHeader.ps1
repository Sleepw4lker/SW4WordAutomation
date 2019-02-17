Function Add-WordPictureToHeader {

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$True)]
        [ValidateRange(1,[int16]::MaxValue)]
        [Int]
        $Section,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Left = 0,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Top = 0,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Width,

        [Parameter(Mandatory=$False)]
        [ValidateRange(1,[int16]::MaxValue)]
        [int]
        $Height
    )

    process {

        Write-Verbose "Inserting Picture $File at current Selection"

        $App.ActiveDocument.Sections($Section).Headers | ForEach-Object {

            # https://docs.microsoft.com/en-us/office/vba/api/word.shapes.addpicture
            $Range = $_.Range
            [void]$App.ActiveDocument.Shapes.AddPicture(
                $File,
                $False,
                $True,
                $Left,
                $Top,
                [math]::Ceiling($Width),
                [math]::Ceiling($Height),
                $Range
            )

        }

    }

}