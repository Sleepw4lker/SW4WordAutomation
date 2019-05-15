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
        $Doc,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Caption,

        # To Do: Parameter Validation
        [Parameter(Mandatory=$False)]
        [String]
        $Style = "Table Grid"
    )

    begin {

        # Assuming that the Function was called via the $App Parameter,
        # we take the currently active Document as the Document to process
        If (-not $Doc) {
            $Doc = $App.ActiveDocument
        }

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

                $Table.Cell($Currentrow, $CurrentColumn).Range.Text = ($_ -as [System.string])
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
                $Table.Cell($Currentrow, $CurrentColumn).Range.Text = $CellValue

            }
            $CurrentColumn++

        }

        Write-Verbose "Table has now $($Table.Range.Rows.Count) Rows"

    }

    End {

        # Move to Set-WordTableCaption with -Label, -Above and -Below Switch
        If ($Caption) {

            Write-Verbose "Adding Caption $Caption"
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

        # Move to Set-WordTableStyle
        If ($Style) {

            Write-Verbose "Setting Table Style to $Style"
            Try {
                $Table.Style = $Style
            }
            Catch {
                
            }

        }

        # Set-WordTableAutoFitBehavior
        # Set-WordTableBehavior
        # Set-WordTableFontStyle with -HeaderStyle and -ContentStyle
        # Move the Selection below the Table when finished and before returning the Object

        $Table

    }

}