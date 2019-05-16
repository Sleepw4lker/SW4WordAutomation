<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Set-WordPaperFormat {

    # You must pass a "Word.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Alias("WordDoc")]
        [Alias("Document")]
        [Microsoft.Office.Interop.Word.Document]
        $Doc,

        [Parameter(Mandatory=$True)]
        [ValidateSet(
            "10x14",
            "11x17",
            "Letter",
            "LetterSmall",
            "Legal",
            "Executive",
            "A3",
            "A4",
            "A4Small",
            "A5",
            "B4",
            "B5",
            "CSheet",
            "DSheet",
            "ESheet",
            "FanfoldLegalGerman",
            "FanfoldStdGerman",
            "FanfoldUS",
            "Folio",
            "Ledger",
            "Note",
            "Quarto",
            "Statement",
            "Tabloid",
            "Envelope9",
            "Envelope10",
            "Envelope11",
            "Envelope12",
            "Envelope14",
            "EnvelopeB4",
            "EnvelopeB5",
            "EnvelopeB6",
            "EnvelopeC3",
            "EnvelopeC4",
            "EnvelopeC5",
            "EnvelopeC6",
            "EnvelopeC65",
            "EnvelopeDL",
            "EnvelopeItaly",
            "EnvelopeMonarch",
            "EnvelopePersonal")]
        [String]
        $Format
    )

    process {

        # To Do: Convert to something more elegant? (perhaps a Hashtable)

        Switch ($Format) {
            "10x14" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaper10x14
            }
            "11x17" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaper11x17
            }
            "Letter" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLetter
            }
            "LetterSmall" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLetterSmall
            }
            "Legal" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLegal
            }
            "Executive" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperExecutive
            }
            "A3" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA3
            }
            "A4" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4
            }
            "A4Small" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4Small
            }
            "A5" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA5
            }
            "B4" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperB4
            }
            "B5" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperB5
            }
            "CSheet" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperCSheet
            }
            "DSheet" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperDSheet
            }
            "ESheet" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperESheet
            }
            "FanfoldLegalGerman" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFanfoldLegalGerman
            }
            "FanfoldStdGerman" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFanfoldStdGerman
            }
            "FanfoldUS" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFanfoldUS
            }
            "Folio" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperFolio
            }
            "Ledger" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperLedger
            }
            "Note" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperNote
            }
            "Quarto" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperQuarto
            }
            "Statement" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperStatement
            }
            "Tabloid" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperTabloid
            }
            "Envelope9" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope9
            }
            "Envelope10" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope10
            }
            "Envelope11" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope11
            }
            "Envelope12" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope12
            }
            "Envelope14" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelope14
            }
            "EnvelopeB4" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeB4
            }
            "EnvelopeB5" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeB5
            }
            "EnvelopeB6" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeB6
            }
            "EnvelopeC3" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC3
            }
            "EnvelopeC4" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC4
            }
            "EnvelopeC5" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC5
            }
            "EnvelopeC6" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC6
            }
            "EnvelopeC65" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeC65
            }
            "EnvelopeDL" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeDL
            }
            "EnvelopeItaly" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeItaly
            }
            "EnvelopeMonarch" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopeMonarch
            }
            "EnvelopePersonal" {
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperEnvelopePersonal
            }
            default {
                #A4
                $Size = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4
            }

        }

        Write-Verbose -Message "Setting Paper Format to ""$Format"""

        # https://docs.microsoft.com/en-us/office/vba/api/Word.sections
        $Doc.Sections | ForEach-Object {

            # https://docs.microsoft.com/en-us/office/vba/api/word.pagesetup.papersize
            # https://gallery.technet.microsoft.com/office/Change-Default-Paper-Size-451f74f8
            $_.PageSetup.PaperSize = $Size
        }

    }

}