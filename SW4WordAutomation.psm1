Try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
}
Catch {
    Write-Error -Message "Microsoft Office seems not to be installed."
    break
}

# http://iextendable.com/2018/07/04/powershell-how-to-structure-a-module/

# The first gci block loads all of the functions in the Export and Private directories. 
# The -Recurse argument allows me to group functions into subdirectories as appropriate in larger modules.
Get-ChildItem *.ps1 -path Export,Private -Recurse | ForEach-Object {

    . $_.FullName

}

# The second gci block exports only the functions in the Export directory. 
# Notice the use of the -Recurse argument again.
Get-ChildItem *.ps1 -path Export -Recurse | ForEach-Object {

    Export-ModuleMember $_.BaseName

}