Function Test-WordIsValidStyle {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Style,

        [Parameter(Mandatory=$False)]
        [ValidateSet("Table","Paragraph")]
        [String]
        $Type  = "Paragraph"
    )

    begin {



    }
    
    process {

        # This is by far the fastest Method that I found
        # $Doc.Styles | Where-Object { $_.NameLocal -eq $Style } would take ages in comparison
        Try { 
            $StyleObject = $Doc.Styles($Style) | Select-Object NameLocal,Type
        }
        Catch {
            # Not Style found, skip here and exit
            return $False
        }

        Switch ($Type) {

            "Paragraph" {
                [Int]$TypeToSearchFor = [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph
            }

            "Table" {
                [Int]$TypeToSearchFor = [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeTable
            }

        }

        If ($StyleObject.Type -eq $TypeToSearchFor) {
            return $True
        }
        Else {
            return $False
        }

    }
    
}