<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Add-WordDraftWatermark {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Word.ApplicationClass]
        $App
    )

    process {

        $App.Templates.LoadBuildingBlocks()

        # Try to get english ones
        $BuildingBlocks = 
            $App.Templates | 
            Where-Object { (($_.name -eq 'Built-In Building Blocks.dotx') -and ($_.LanguageID -eq 1033)) } | 
            Select-Object -First 1

        # Revert to Default
        If (-not ($BuildingBlocks)) {
            $BuildingBlocks = 
                $App.Templates | 
                Where-Object { (($_.name -eq 'Built-In Building Blocks.dotx')) } | 
                Select-Object -First 1
        }

        If ($BuildingBlocks) {

            Write-Verbose -Message "Adding DRAFT Watermark"

            # 1 ... ASAP 1
            # 2 ... ASAP 2
            # 3 ... CONFIDENTIAL 1
            # 4 ... CONFIDENTIAL 2
            # 5 ... DO NOT COPY 1
            # 6 ... DO NOT COPY 2
            # 7 ... DRAFT 1
            # 8 ... DRAFT 2
            # 9 ... SAMPLE 1
            # 10 ... SAMPLE 2
            # 11 ... URGENT 1
            # 12 ... URGENT 2
            $Watermark = $BuildingBlocks.BuildingBlockEntries.Item(7)

            $SectionIndex++

            $App.ActiveDocument.Sections | ForEach-Object {

                $_.Headers | ForEach-Object {

                    If ($SectionIndex -eq 2) {

                        [void]$Watermark.Insert($_.Range,$True)

                    }

                }

            }

        }

    }

}