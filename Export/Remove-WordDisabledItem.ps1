<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Remove-WordDisabledItem {

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path $_})]
        [String]
        $File
    )

    process {

        # https://stackoverflow.com/questions/751048/how-to-programatically-re-enable-documents-in-the-ms-office-list-of-disabled-fil

        $WordVersion = Get-WordVersion

        # Converts the File Name string to UTF16 Hex
        $File_UniHex = ""
        [System.Text.Encoding]::ASCII.GetBytes($File.ToLower()) | ForEach-Object { 
            $File_UniHex += "{0:X2}00" -f $_
        }

        Try {
            # Tests to see if the Disabled items registry key exists
            $DisabledItemsRegistryKey = (Get-Item "HKCU:\Software\Microsoft\Office\${WordVersion}.0\Word\Resiliency\DisabledItems\")
        }
        Catch {
            # Nothing yet
        }

        If ($NULL -ne $DisabledItemsRegistryKey) {

            #Cycles through all the properties and deletes it if it contains the file name.
            Foreach ($DisabledItem in $DisabledItemsRegistryKey.Property) {

                $Value = ""

                ($DisabledItemsRegistryKey | Get-ItemProperty).$DisabledItem | ForEach-Object{
                    $Value += "{0:X2}" -f $_
                }

                If ($Value.Contains($File_UniHex)) {

                    Write-Verbose "Removing $File from the List of Disabled Items."

                    $DisabledItemsRegistryKey | Remove-ItemProperty -name $DisabledItem

                }
            }
        }
    }
}