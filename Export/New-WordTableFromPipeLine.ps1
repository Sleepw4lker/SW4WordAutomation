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
            Mandatory = $False, 
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

        If ($CurrentRow -eq 1) {

            # The first Object that enters the pipeline determines the List of Properties to display
            # It is advised that you explicitly specify a list of properties when pipelining into the function
            # by using Select-Object
            $Properties = $Object | Get-Member -MemberType Property,NoteProperty | Select-Object -Property Name

            # If there are fewer Columns than our Data, add the Columns
            For ($i = $Table.Range.Columns.Count; $i -lt $Properties.Count; $i++) {
                [void]$Table.Columns.Add()
            }

            $CurrentColumn = 1

            $Properties | ForEach-Object { 

                $Table.Cell($Currentrow, $CurrentColumn).Range.Text = ($_.Name -as [System.string])
                $CurrentColumn++

            }

            Write-Verbose "Table has now $($Table.Range.Columns.Count) Columns)"

        }

        $CurrentRow++

        # Insert a Row for the Data
        [void]$Table.Rows.Add()

        $CurrentColumn = 1

        $Properties | ForEach-Object { 

            # Objects may have a differing Property Set
            # Thus checking if the Property is present to avoid confusion
            If ($_.Name -in  ($Object | Get-Member -MemberType Property,NoteProperty | Select-Object Name).Name) {

                $CellValue = $Object."$($_.Name)" -as [System.string]
                Write-Verbose "Property $($_.Name) has Value $CellValue in Row $Currentrow, Col $CurrentColumn"
                $Table.Cell($Currentrow, $CurrentColumn).Range.Text = $CellValue

            }
            $CurrentColumn++

        }

        Write-Verbose "Table has now $($Table.Range.Rows.Count) Rows"

    }

    End {

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

        If ($Style) {

            Write-Verbose "Setting Table Style to $Style"
            Try {
                $Table.Style = $Style
            }
            Catch {
                
            }

        }

        $Table

    }

}