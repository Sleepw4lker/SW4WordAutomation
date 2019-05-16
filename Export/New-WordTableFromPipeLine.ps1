<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function New-WordTableFromPipeLine {

    # Create a Table, return a Table Object

    [cmdletbinding()]
    Param (
        [Parameter(
            Position = 0,
            Mandatory = $True, 
            ValuefromPipeline = $True
        )]    
        [psobject]$Object,

        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Caption,

        [Parameter(Mandatory=$False)]
        [ValidateScript({
            Test-WordIsValidStyle -Doc $Doc -Style $_ -Type Table
        })]
        [String]
        $TableStyle,

        [Parameter(Mandatory=$True)]
        [ValidateScript({
            Test-WordIsValidStyle -Doc $Doc -Style $_
        })]
        [String]
        $Style,

        [Parameter(Mandatory=$False)]
        [ValidateScript({
            Test-WordIsValidStyle -Doc $Doc -Style $_
        })]
        [String]
        $HeaderStyle,

        [Parameter(Mandatory=$False)]
        [Switch]
        $RepeatHeader,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoNewLine,

        [Parameter(Mandatory=$False)]
        [Switch]
        $StyleColumnBands,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoStyleFirstColumn,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoStyleHeadingRows,

        [Parameter(Mandatory=$False)]
        [Switch]
        $StyleLastColumn,

        [Parameter(Mandatory=$False)]
        [Switch]
        $StyleLastRow,

        [Parameter(Mandatory=$False)]
        [Switch]
        $NoStyleRowBands
    )

    begin {

        $Selection = $Doc.ActiveWindow.Selection

        $Columns = 1
        $Rows = 1
        $CurrentRow = 1

        [Object[]]$Properties

        # https://msdn.microsoft.com/en-us/vba/word-vba/articles/tables-add-method-word
        $Table = $Doc.Tables.Add(
            $Selection.Range,
            $Rows,
            $Columns,
            # https://docs.microsoft.com/en-us/office/vba/api/word.wddefaulttablebehavior
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            # https://docs.microsoft.com/en-us/office/vba/api/word.wdautofitbehavior
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
        )

    }

    process {

        Write-Verbose "Row $CurrentRow"

        # Before the first Row is written into the Table, we first have to build the Header Row
        # Thus we enumerate the Properties, create the necessary columns and fill them with the Property Names 
        If ($CurrentRow -eq 1) {

            # The first Object that enters the pipeline determines the List of Properties to display
            # It is therefores advised that you explicitly specify a list of properties when pipelining 
            # into the function by using Select-Object
            # https://stackoverflow.com/questions/51894114/powershell-is-there-a-way-to-get-proper-order-of-properties-in-select-object-so?rq=1
            $Properties = ($Object.PSObject.Properties).Name

            # If there are fewer Columns than our Data, add the Columns
            For ($i = $Table.Range.Columns.Count; $i -lt $Properties.Count; $i++) {
                [void]$Table.Columns.Add()
            }

            # Fill the Cells with the Property Names
            $CurrentColumn = 1
            $Properties | ForEach-Object { 

                $CurrentCell = $Table.Cell($Currentrow, $CurrentColumn).Range
                $CurrentCell.Text = ($_ -as [System.string])
                If ($HeaderStyle) {
                    $CurrentCell.Style = $Doc.Styles($HeaderStyle)
                }
                $CurrentColumn++

            }

            Write-Verbose "Table has now $($Table.Range.Columns.Count) Columns)"

        }

        # The first Row that has Data is Row #1, as #1 is the Header Row
        $CurrentRow++

        # Insert a Row for the Data
        [void]$Table.Rows.Add()

        # Fill the Cells with the Data
        $CurrentColumn = 1
        $Properties | ForEach-Object {

            # Objects may have a differing Property Set
            # Thus checking if the Property is present, and if not, we leave the Cell empty
            If ($_ -in ($Object.PSObject.Properties).Name) {

                $CellValue = $Object."$($_)" -as [System.string]
                Write-Verbose "Property $($_) has Value $CellValue in Row $Currentrow, Col $CurrentColumn"
                $CurrentCell = $Table.Cell($Currentrow, $CurrentColumn).Range
                $CurrentCell.Text = $CellValue
                If ($Style) {
                    $CurrentCell.Style = $Doc.Styles($Style)
                }

            }
            $CurrentColumn++

        }

        Write-Verbose "Table has now $($Table.Range.Rows.Count) Rows"

    }

    End {

        If ($Caption) {
            Set-WordTableCaption -Table $Table -Caption $Caption
        }

        If ($TableStyle) {
            Set-WordTableStyle -Table $Table -Style $TableStyle
        }

        If ($RepeatHeader.IsPresent) {
            # https://docs.microsoft.com/en-us/office/vba/api/word.row.headingformat
            $Table.Rows(1).HeadingFormat = $True
        }

        $Table.ApplyStyleColumnBands	= $StyleColumnBands.IsPresent
        $Table.ApplyStyleFirstColumn	= (-not $NoStyleFirstColumn.IsPresent)
        $Table.ApplyStyleHeadingRows	= (-not $NoStyleHeadingRows.IsPresent)
        $Table.ApplyStyleLastColumn	    = $StyleLastColumn.IsPresent
        $Table.ApplyStyleLastRow	    = $StyleLastRow.IsPresent
        $Table.ApplyStyleRowBands	    = (-not $NoStyleRowBands.IsPresent)

        # Set-WordTableAutoFitBehavior
        # Set-WordTableBehavior

        If (-not $NoNewLine.IsPresent) {

            # Move the Selection below the Table when finished and before returning the Object
            $Table.Select()

            $Selection.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

            # If the Table has a Caption, the Cursor will otherwise land at the beginning of the Caption
            If ($Caption) {

                # Move to End of Line
                $Selection.EndKey([Microsoft.Office.Interop.Word.wdUnits]::wdLine)
                
                # Enter Key
                $Selection.TypeParagraph()
                
            }

        }

        $Table

    }

}