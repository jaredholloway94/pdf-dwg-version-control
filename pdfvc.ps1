Import-Module MergePdf
Add-Type -AssemblyName System.Windows.Forms

$ConvertPdfToDwgScr = Join-Path -Path $HOME -ChildPath ".pdfvc\ConvertPdfToDwg.scr"
$NewMasterDwgScr = Join-Path -Path $HOME -ChildPath ".pdfvc\NewMasterDwg.scr"
$AddXrefScr = Join-Path -Path $HOME -ChildPath ".pdfvc\AddXref.scr"
$UpdateXrefScr = Join-Path -Path $HOME -ChildPath ".pdfvc\UpdateXref.scr"

# <UI & IO> -------------------------------------------------------------------
function Get-Time {
    return (Get-Date -Format "HH:mm:ss")

}
function Get-UserInput {
    param (
        $InitStatements,

        $MainQuestion,

        $Options
    )
    Write-Host
    if($InitStatements) {$InitStatements | ForEach-Object {Write-Host $_} }
    Write-Host
    $Prompt = ($MainQuestion + ( " [{0}]" -f [String]::Join('/',$Options) ))
    $UserInput = Read-Host $Prompt
    if ($Options -contains $UserInput) {<# user input is valid, return it. #>}
    else{# user input is invalid, define recursive _AskAgain func, call it.
        function _AskAgain {
            param ([String[]]$Options)
            $Prompt = ("Please enter one of [{0}]" -f [String]::Join('/',$Options))
            $UserInput = Read-Host $Prompt
            if ($Options -contains $UserInput) {}
            else {$UserInput = _AskAgain $Options}
            return $UserInput
        }
        # this guarantees $UserInput has valid value before returning
        #   (or will eventually hit max rec. depth if user is stupid or mean).
        $UserInput = _AskAgain $Options
    }
    return $UserInput
}
# RETO [switch statement to handle user input]

function Get-UserFolder {
    [CmdletBinding()]
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $null = $FolderDialog.ShowDialog()
    $UserFolder = $FolderDialog.SelectedPath
    return $UserFolder
}
# RETO [?]

function Get-UserFile {
    [CmdletBinding()]
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $null = $FileDialog.ShowDialog()
    $UserFile = $FileDialog.FileName
    return $UserFile
}
# RETO [?]

function Import-Project {
    param (
        $ProjFile
        )
    Write-Host ("Opening {0}..." -f $ProjFile)
    # read in $Project from serialized clixml format
    $Project = Import-CliXml $ProjFile
    Write-Host ("Done Opening {0}." -f $ProjFile)
    return $Project
}
# RETO [?]

function Export-Project {
    param (
        #[PSTypeName('Project')]
        $Project
    )
    $ProjFile = $Project.ProjFile
    Write-Host ("Saving {0}..." -f $ProjFile)
    # save $Project in serialized clixml format
    Export-CliXml -Path $ProjFile -InputObject $Project -Depth 2 -Force
    Write-Host ("Done Saving  {0}." -f $ProjFile)
    return $ProjFile
}
# RETO [?]

function Exit-Project {
    param (
        #[PSTypeName('Project')]
        $Project
    )
    if ($Project) {Export-Project $Project}
    Read-Host "Press Enter to exit..."
    Exit
}
# TERM

# </UI & IO> ------------------------------------------------------------------
# GOTO Entry-Point




# <Entry-Point> ---------------------------------------------------------------

# ->FROM Entry-Point
#   <Main> --------------------------------------------------------------------

#   ->FROM Main
#     <Find-AutoCadCoreConsole> -----------------------------------------------

#     ->FROM Find-AutoCadCoreConsole
#       <Resolve-AcccNotFoundInPath> ------------------------------------------

#       ->FROM Resolve-AcccNotFoundInPath
#         <Resolve-OneAcccFoundInFS> ------------------------------------------
function Resolve-OneAcccFoundInFS {
    Write-Host ("Found it at {0}." -f $AcccPath)
}
#         </Resolve-OneAcccFoundInFS> -----------------------------------------
#       <-RETO Resolve-AcccNotFoundInPath


#       ->FROM Resolve-AcccNotFoundInPath
#         <Resolve-ManyAcccFoundInFS> -----------------------------------------
function Resolve-ManyAcccFoundInFS {
    $UserInput = Get-UserInput `
            @("Multiple versions found:",
                "`r`n{0}`r`n" -f ([String]::Join("`r`n",$AcccPath))
            ) `
            "Which would you like to use?" `
            @(1..($AcccPath.Count))
        $AcccPath = $AcccPath[([int]$UserInput - 1)]
        Write-Host ("Using version at {0}." -f $AcccPath)
}
#         </Resolve-ManyAcccFoundInFS> ----------------------------------------
#       <-RETO Resolve-AcccNotFoundInPath


#       ->FROM Resolve-AcccNotFoundInPath
#         <Resolve-NoAcccFoundInFS> -------------------------------------------
function Resolve-NoAcccFoundInFS {
    # if we still didn't find it, we're done trying to help. exit.
    Write-Host ("Could not find AutoCAD Core Console (accoreconsole.exe)")
    Write-Host ("Please install or locate AutoCAD, then try again.")
    Exit-Project
}
#         </Resolve-NoAcccFoundInFS> ------------------------------------------
#       <-RETO Resolve-AcccNotFoundInPath


#       ->FROM Resolve-AcccNotFoundInPath
#         <Resolve-AutodeskDirNotFound> ---------------------------------------
function Resolve-AutodeskDirNotFound {
    # ask the user to supply the non-standard install dir, or exit.
    $UserInput = Get-UserInput `
        @("Could not find folder 'C:\Program Files\Autodesk'.",
        "This tool requires AutoCAD to be installed.") `
        "Enter 'y' to locate Autodesk folder elsewhere, or 'n' to exit." `
        @('y','n')
    switch ($UserInput) {
        # user wants to supply non-standard AutoCAD install dir
        'y' {$AutodeskDir = Get-UserFolder -WarningAction SilentlyContinue}
        # AutoCAD not installed, user wants to exit.
        'n' {Exit-Project}
    }
    return $AutodeskDir
}
#         </Resolve-AutodeskDirNotFound> --------------------------------------
#       <-RETO Resolve-AcccNotFoundInPath

function Resolve-AcccNotFoundInPath {
    # if it's not in $Path, try to find it in the file system...
    Write-Host ("[{0}] Didn't find it in the `$PATH environment variable..." -f (Get-Time))
    # Start with the default install directory...
    $AutodeskDir = [System.IO.DirectoryInfo]'C:\Program Files\Autodesk'
    # if the default install directory does not exist...
    if ($false -eq (Test-Path $AutodeskDir)) {
        # ask the user to supply the non-standard install dir, or exit.
        $AutodeskDir = Resolve-AutodeskDirNotFound
    }
    # if we're here, we have a dir and are ready to look for accoreconsole
    Write-Host ("Looking in {0}..." -f $AutodeskDir)
    $AcccPath = Get-ChildItem -Recurse -Path $AutodeskDir -Filter 'accoreconsole.exe'
    switch ($AcccPath.Count) {
        # if we still didn't find it, we're done trying to help. exit.
        0 {Resolve-NoAcccFoundInFS}
        # if we found one version, use it.
        1 {Resolve-OneAcccFoundInFS}
        # if we found more than one, ask the user which to use.
        Default {Resolve-ManyAcccFoundInFS}
    }
    # also add it to $Path so we don't have to do this again.
    Write-Host ("Adding it to `$PATH environment variable... Done.")
    & $HOME/.pdfvc/append_user_path.cmd $AcccPath
    return
}
#       </Resolve-AcccNotFoundInPath> -----------------------------------------
#     <-RETO Find-AutoCadCoreConsole


#     ->FROM Find-AutoCadCoreConsole
#       <Resolve-AcccFoundInPath> ---------------------------------------------
function Resolve-AcccFoundInPath {
    Write-Host ("Found it in the `$PATH environment variable.")
    return
}
#       </Resolve-AcccFoundInPath> --------------------------------------------
#     <-RETO Find-AutoCadCoreConsole

function Find-AutoCadCoreConsole {
    Write-Host ("Looking for AutoCAD Core Console (accoreconsole.exe)...")
    # if accoreconsole.exe is in $PATH,
    if ($null -ne (Get-Command 'accoreconsole' -ErrorAction SilentlyContinue)) {
        # all good, return.
        Resolve-AcccFoundInPath
    } else {
        # if it's not in $Path, try to find it in the file system...
        Resolve-AcccNotFoundInPath
    }
    return
}
#     </Find-AutoCadCoreConsole> ----------------------------------------------
#   <-RETO Main
#   ->GOTO Show-ProjectMenu




#   ->FROM Main
#     <Show-ProjectMenu> ------------------------------------------------------

#     ->FROM Show-Project-Menu
#       <Add-Revision> --------------------------------------------------------

#       ->FROM Add-Revision
#         <New-Revision> ------------------------------------------------------
function New-Revision {
    param(
        #[PSTypeName('Project')]
        $Project,

        $RevName,

        $RevDate
    )
    $BaseName = ("{0}_{1}_{2}" -f ($RevDate,$Project.Name,$RevName))
    $PdfPath = (Join-Path $Project.RevDir ($BaseName + '.pdf'))
    $TxtPath = (Join-Path $Project.RevDir ($BaseName + '.txt'))
    $DirPath = (Join-Path -Path $Project.RevDir -ChildPath $BaseName)
    Write-Host ("[{0}] Creating new revision {1}..." -f (Get-Time),$RevName)
    # create new [Revision] obj
    $Revision = [PSCustomObject]@{
        # Type declaration
        PSTypeName = 'Revision'
        # Revision attrs
        Name = $RevName
        Date = $RevDate
        BaseName = $BaseName
        # Revision aggts
        Sheets = {@()}.Invoke()
        SheetGroups = {@()}.Invoke()
        # Project ref
        Project = $Project
        # Paths
        PdfPath = $PdfPath
        TxtPath = $TxtPath
        DirPath = $DirPath
    }
    Write-Host ("[{0}] Done creating new revision {1}." -f (Get-Time),$RevName)
    return $Revision
}
#         </New-Revision> -----------------------------------------------------
#       <-RETO Add-Revision
#       ->GOTO Initialize-Revision


#       ->FROM Add-Revision
#         <Initialize-Revision> -----------------------------------------------

#         ->FROM Initialize-Revision
#           <Split-Revision> ------------------------------------------
function Split-Revision {
    param (
        #[PSTypeName('Revision')]
        $Revision
    )
    Write-Host ("[{0}] Splitting revision pdf {1}.pdf into sheet pdfs..." -f (Get-Time),$Revision.BaseName)
    # split $Revision.pdf file into individual sheets; rm extra file
    $null = pdftk $Revision.PdfPath burst output (Join-Path -Path $Revision.DirPath -ChildPath 'pg_%04d.pdf')
    $null = Get-ChildItem -Path $Revision.DirPath -Filter 'doc_data.txt' | Remove-Item
    # rename individual sheets as per $Revision.txt (sheet names) file
    $SheetPdfs = Get-ChildItem -Path $Revision.DirPath -Filter "*.pdf"
    $SheetNames = Get-Content -Path $Revision.TxtPath
    for ($i = 0; $i -lt $SheetPdfs.Count; $i++) {
        $oldBase = $SheetPdfs[$i].BaseName
        $newBase = '{0}_{1}_{2}' -f $SheetNames[$i],$Revision.Name,$Revision.Date
        $null = Rename-Item `
            $SheetPdfs[$i].FullName `
            $SheetPdfs[$i].FullName.Replace($oldBase,$newBase)
    }
    Write-Host ("[{0}] Done splitting revision pdf {1}.pdf into sheet pdfs." -f (Get-Time),$Revision.BaseName)
    return $Revision
}
#           </Split-Revision> -----------------------------------------
#         <-RETO Initialize-Revision
#         ->GOTO Add-Sheets


#         ->FROM Initialize-Revision
#           <Add-Sheets> ------------------------------------------------------

#           ->FROM Add-Sheets
#             <Add-Sheet> -----------------------------------------------------

#             ->FROM Add-Sheet
#               <New-Sheet> ---------------------------------------------------
function New-Sheet {
    param (
        #[PSTypeName('Revision')]
        $Revision,

        
        $PdfPath
    )
    $BaseName = (Get-Item $PdfPath).BaseName
    $Name, $RevName, $RevDate = $BaseName.Split('_')
    $GroupName, $GroupNumber = $Name.Split('.')
    $Project = $Revision.Project
    Write-Host ("[{0}] Creating new sheet {1}..." -f (Get-Time),$Name)
    $Sheet = [PSCustomObject]@{
        # Type declaration
        PSTypeName = 'Sheet'
        # Sheet attrs
        Name = $Name
        GroupName = $GroupName
        GroupNumber = $GroupNumber
        RevName = $RevName
        RevDate = $RevDate
        BaseName = $BaseName
        # Revision ref
        Revision = $Revision
        #RevSheetGroup
        # Project ref
        Project = $Project
        #ProjSheetGroup
        # Paths
        RevPdfPath = $PdfPath
        #RevDwgPath
        #ScrPath
        #XrefScrPath
        #ProjPdfPath
        #ProjDwgPath
    }
    Write-Host ("[{0}] Done creating new sheet {1}." -f (Get-Time),$Name)
    return $Sheet
}
#               </New-Sheet> --------------------------------------------------
#             <-RETO Add-Sheet
#             ->GOTO Initialize-Sheet


#             ->FROM Add-Sheet
#               <Initialize-Sheet> --------------------------------------------

#               ->FROM Initialize-Sheet
#                 <Set-RevSheetGroup> -----------------------------------------

#                 ->FROM Set-RevSheetGroup
#                   <Get-RevSheetGroup> ---------------------------------------

#                   ->FROM Get-RevSheetGroup
#                     <Add-RevSheetGroup> -------------------------------------

#                     ->FROM Add-RevSheetGroup
#                       <New-RevSheetGroup> -----------------------------------
function New-RevSheetGroup {
    param (     
        #[PSTypeName('Sheet')]
        $Sheet,

        #[PSTypeName('Revision')]
        $Revision
    )
    if ($null -eq $Revision) {$Revision = $Sheet.Revision}
    $GroupName = $Sheet.GroupName
    Write-Host ("[{0}] Creating new revision sheetgroup {1}..." -f (Get-Time),$GroupName)
    $RevSheetGroup = [PSCustomObject]@{
        # Type declaration
        PSTypeName = 'RevSheetGroup'
        # SheetGroup attrs
        Name = $GroupName
        # SheetGroup aggts
        Sheets = {@()}.Invoke()
        # Revision ref
        Revision = $Revision
    }
    Write-Host ("[{0}] Done creating new revision sheetgroup {1}." -f (Get-Time),$GroupName)
    return $RevSheetGroup
}
#                       </New-RevSheetGroup> ----------------------------------
#                     <-RETO Add-RevSheetGroup
#                     ->GOTO Initialize-RevSheetGroup


#                     ->FROM Add-RevSheetGroup
#                       <Initialize-RevSheetGroup> ------------------------------------
function Initialize-RevSheetGroup {
    param (
        #[PSTypeName('RevSheetGroup')]
        $RevSheetGroup
    )
    Write-Host ("[{0}] Initializing new revision sheetgroup {1}..." -f (Get-Time),$RevSheetGroup.Name)
    $RevSheetGroup.Revision.SheetGroups.Add($RevSheetGroup)
    Write-Host ("[{0}] Done initializing new revision sheetgroup {1}." -f (Get-Time),$RevSheetGroup.Name)
    return $RevSheetGroup
}
#                       </Initialize-RevSheetGroup> -----------------------------------
#                     <-RETO Add-RevSheetGroup

function Add-RevSheetGroup {
    param (
        #[PSTypeName('Sheet')]
        $Sheet,

        #[PSTypeName('Revision')]
        $Revision
    )
    if ($null -eq $Revision) {$Revision = $Sheet.Revision}
    Write-Host ("[{0}] Adding new revision sheetgroup {1} to revision {2}..." -f (Get-Time),$Sheet.GroupName,$Revision.Name)
    $RevSheetGroup = New-RevSheetGroup -Sheet $Sheet -Revision $Revision
    $null = Initialize-RevSheetGroup $RevSheetGroup
    Write-Host ("[{0}] Done adding new revision sheetgroup {1} to revision {2}." -f (Get-Time),$Sheet.GroupName,$Revision.Name)
    return $RevSheetGroup
}
#                     </Add-RevSheetGroup> --------------------------------------------
#                   <-RETO Get-RevSheetGroup

function Get-RevSheetGroup {
    param (        
        #[PSTypeName('Sheet')]
        $Sheet,

        #[PSTypeName('Revision')]
        $Revision
    )
    if ($null -eq $Revision) {$Revision = $Sheet.Revision}
    Write-Host ("[{0}] Retrieving revision sheetgroup for sheet {1}..." -f (Get-Time),$Sheet.Name)
    # try to get existing RevSheetGroup from Revision
    $RevSheetGroup = $Revision.SheetGroups | Where-Object {$_.Name -eq $Sheet.GroupName}
    # if this there is no matching RevSheetGroup in the Revision, add it.
    if ($null -eq $RevSheetGroup) {
        Write-Host ("[{0}] Revision sheetgroup {1} does not exist yet." -f (Get-Time),$Sheet.GroupName)
        $RevSheetGroup = Add-RevSheetGroup -Sheet $Sheet -Revision $Revision
    }
    Write-Host ("[{0}] Done retrieving revision sheetgroup for sheet {1}." -f (Get-Time),$Sheet.Name)
    return $RevSheetGroup
}
#                   </Get-RevSheetGroup> ----------------------------------------------
#                 <-RETO Set-RevSheetGroup

function Set-RevSheetGroup {
    param (
        #[PSTypeName('Sheet')]
        $Sheet,

        #[PSTypeName('RevSheetGroup')]
        $RevSheetGroup
    )
    Write-Host ("[{0}] Assigning sheet {1} to a revision sheetgroup..." -f (Get-Time),$Sheet.Name)
    if ($null -eq $RevSheetGroup) {$RevSheetGroup = Get-RevSheetGroup -Sheet $Sheet}
    # set $Sheet.RevSheetGroup = $RevSheetGroup
    $Sheet | Add-Member -NotePropertyName 'RevSheetGroup' -NotePropertyValue $RevSheetGroup
    # add $Sheet to $Rev.SheetGroup[$RevSheetGroup]
    $RevSheetGroup.Sheets.Add($Sheet)
    Write-Host ("[{0}] Done assigning sheet {1} to a revision sheetgroup." -f (Get-Time),$Sheet.Name)
    return $Sheet
}
#                 </Set-RevSheetGroup> ------------------------------------------------
#               <-RETO Initialize-Sheet
#               ->GOTO Convert-PdfToDwg


#               ->FROM Initialize-Sheet
#                 <Convert-PdfToDwg> --------------------------------------------------

#                 ->FROM Convert-PdfToDwg
#                   <Write-Script_Convert-PdfToDwg> -----------------------------------
function Write-Script_ConvertPdfToDwg {
    param(
        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Writing script to convert sheet {1} from Pdf to Dwg..." -f (Get-Time),$Sheet.Name)
    # compose path and set $Sheet.ScrPath
    $ScrPathName = ("PdfToDwg_{0}.scr" -f $Sheet.BaseName)
    $ScrPath = Join-Path -Path $Sheet.Revision.DirPath -ChildPath $ScrPathName
    $Sheet | Add-Member -NotePropertyName 'ScrPath' -NotePropertyValue $ScrPath
    # modify ConvertPdfToDwg.scr for $this Sheet; save to $ScrPath
    $scr = New-Item `
        -ItemType File `
        -Path $Sheet.Revision.DirPath `
        -Name $ScrPathName `
        -Value (Get-Content $ConvertPdfToDwgScr -Raw
            ).Replace('_pdfpath_',$Sheet.RevPdfPath
            ).Replace('_dwgpath_',$Sheet.RevDwgPath
            )
    Write-Host ("[{0}] Done writing script to convert sheet {1} from Pdf to Dwg." -f (Get-Time),$Sheet.Name)
    return $scr
}
#                   </Write-Script_Convert-PdfToDwg> ------------------------------------
#                 <-RETO Convert-PdfToDwg

function Convert-PdfToDwg {
    param (
        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Converting sheet {1} from Pdf to Dwg..." -f (Get-Time),$Sheet.Name)
    # compose path and set $Sheet.RevDwgPath
    $RevDwgPath = Join-Path -Path $Sheet.Revision.DirPath -ChildPath ($Sheet.BaseName + '.dwg')
    $Sheet | Add-Member -NotePropertyName 'RevDwgPath' -NotePropertyValue $RevDwgPath
    # convert (copy) $Sheet.RevPdf to $Sheet.RevDwg
    $scr = Write-Script_ConvertPdfToDwg $Sheet
    Write-Host ("[{0}] Executing script {1}..." -f (Get-Time),$scr.BaseName)
    & accoreconsole -s $scr
    Write-Host ("[{0}] Done converting sheet {1} from Pdf to Dwg." -f (Get-Time),$Sheet.Name)
    return $Sheet
}
#                 </Convert-PdfToDwg> -------------------------------------------------
#               <-RETO Initialize-Sheet

function Initialize-Sheet {
    param (
        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Initializing sheet {1}..." -f (Get-Time),$Sheet.Name)
    $null = Set-RevSheetGroup -Sheet $Sheet
    $null = Convert-PdfToDwg -Sheet $Sheet
    $Sheet.Revision.Sheets.Add($Sheet)
    Write-Host ("[{0}] Done initializing sheet {1}." -f (Get-Time),$Sheet.Name)
    return $Sheet
}
#               </Initialize-Sheet> ---------------------------------------------------
#             <-RETO Add-Sheet

function Add-Sheet {
    param (
        #[PSTypeName('Revision')]
        $Revision,

        
        $PdfPath
    )
    $BaseName = (Get-Item $PdfPath).BaseName
    $SheetName, $RevName, $RevDate = $BaseName.Split('_')
    Write-Host ("[{0}] Adding sheet {1} to revision {2}..." -f (Get-Time),$SheetName,$Revision.Name)
    $Sheet = New-Sheet -Revision $Revision -PdfPath $PdfPath
    $null = Initialize-Sheet $Sheet
    Write-Host ("[{0}] Done adding sheet {1} to revision {2}..." -f (Get-Time),$SheetName,$Revision.Name)
    return $Sheet
}
#             </Add-Sheet> ------------------------------------------------------------
#           <-RETO Add-Sheets

function Add-Sheets {
    param (
        #[PSTypeName('Revision')]
        $Revision
    )
    Write-Host ("[{0}] Adding sheets to revision {1}..." -f (Get-Time),$Revision.Name)
    # populate $Revision.Sheets with [Sheet]$Sheet.pdf for $Sheet.pdf in $Revision.Dir
    $null = Get-ChildItem -Path $Revision.DirPath | ForEach-Object {
        Add-Sheet -Revision $Revision -PdfPath $_.FullName
    }
    Write-Host ("[{0}] Adding sheets to revision {1}..." -f (Get-Time),$Revision.Name)
    return $Revision
}
#           </Add-Sheets> -------------------------------------------------------------
#         <-RETO Initialize-Revision

#       ->FROM Add-Revision
function Initialize-Revision {
    param (
        #[PSTypeName('Revision')]
        $Revision,

        $PdfPath,

        $TxtPath
    )
    Write-Host ("[{0}] Initializing new revision {1}..." -f (Get-Time),$Revision.Name )
    # copy $Rev.pdf and $Rev.txt to $Project.RevDir; mkdir $Rev.Dir
    $null = Copy-Item -Path $PdfPath -Dest $Revision.PdfPath -InformationAction SilentlyContinue
    $null = Copy-Item -Path $TxtPath -Dest $Revision.TxtPath -InformationAction SilentlyContinue
    $null = New-Item -ItemType Directory -Path $Revision.DirPath -InformationAction SilentlyContinue
    $null = Split-Revision $Revision
    $null = Add-Sheets $Revision
    # add $Rev to $Proj.Revs
    $Revision.Project.Revisions.Add($Revision)
    Write-Host ("[{0}] Done Initializing new revision {1}." -f (Get-Time),$Revision.Name )
    return $Revision
}
#         </Initialize-Revision> ----------------------------------------------------
#       <-RETO Add-Revision
#       ->GOTO Publish-Revision


#       ->FROM Add-Revision
#         <Publish-Revision> --------------------------------------------------------

#         ->FROM Publish-Revision
#           <Get-ProjSheetGroup> ----------------------------------------------------

#           ->FROM Get-ProjSheetGroup
#             <Add-ProjSheetGroup> --------------------------------------------------

#             ->FROM Add-ProjSheetGroup
#               <New-ProjSheetGroup> ------------------------------------------------
function New-ProjSheetGroup {
    param (
        #[PSTypeName('$RevSheetGroup')]
        $RevSheetGroup
    )
    Write-Host ("[{0}] Creating new project sheetgroup {1}..." -f (Get-Time),$RevSheetGroup.Name)
    $Revision = $RevSheetGroup.Revision
    $Project = $Revision.Project
    $Name = $RevSheetGroup.Name
    $ProjSheetGroup = [PSCustomObject]@{
        # Type declaration
        PSTypeName = 'ProjSheetGroup'
        # ProjSheetGroup attrs
        Name = $Name
        # ProjSheetGroup aggts
        Current = {@()}.Invoke()
        Archive = {@()}.Invoke()
        # Revision ref
        Revision = $Revision
        # Project ref
        Project = $Project
        # Paths
        #PdfPath
        #DwgPath
        #ScrPath
    }
    Write-Host ("[{0}] Done creating new project sheetgroup {1}." -f (Get-Time),$RevSheetGroup.Name)
    return $ProjSheetGroup
}
#               </New-ProjSheetGroup> -----------------------------------------------
#             <-RETO Add-ProjSheetGroup
#             ->GOTO Initialize-ProjSheetGroup


#             ->FROM Add-ProjSheetGroup
#               <Initialize-ProjSheetGroup> -----------------------------------------

#               ->FROM Initialize-ProjSheetGroup
#                 <New-MasterDwg> ---------------------------------------------------

#                 ->FROM New-MasterDwg
#                   <Write-Script_NewMasterDwg> ---------------------------------------
function Write-Script_NewMasterDwg {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup
    )
    Write-Host ("[{0}] Writing script to create new master dwg {1}..." -f (Get-Time),$ProjSheetGroup.Name)
    # compose and set $ProjSheetGroup.ScrPath
    $ScrPathName = ("NewMasterDwg_{0}.scr" -f $ProjSheetGroup.Name)
    $ScrPath = Join-Path -Path $ProjSheetGroup.Project.ScrDir -ChildPath $ScrPathName
    $ProjSheetGroup | Add-Member -NotePropertyName 'ScrPath' -NotePropertyValue $ScrPath
    # write the scipt needed to create a new _Master.dwg for $this ProjSheetGroup
    $scr = New-Item `
        -ItemType File `
        -Path $ProjSheetGroup.Project.ScrDir `
        -Name $ScrPathName `
        -Value (Get-Content $NewMasterDwgScr -Raw).Replace('_dwgpath_',$ProjSheetGroup.DwgPath)
    Write-Host ("[{0}] Done writing script to create new master dwg {1}." -f (Get-Time),$ProjSheetGroup.Name)
    return $scr
}
#                   </Write-Script_NewMasterDwg> --------------------------------------
#                 <-RETO New-MasterDwg

function New-MasterDwg {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup
    )
    Write-Host ("[{0}] Creating new master dwg {1}..." -f (Get-Time),$ProjSheetGroup.Name)
    # compose and set $ProjSheetGroup.DwgPath
    $DwgPath = Join-Path -Path $Project.DwgDir -ChildPath ("_{0}.dwg" -f $ProjSheetGroup.Name)
    $ProjSheetGroup | Add-Member -NotePropertyName 'DwgPath' -NotePropertyValue $DwgPath
    # write the scipt needed to create a new _Master.dwg for $this ProjSheetGroup
    $scr = Write-Script_NewMasterDwg $ProjSheetGroup
    Write-Host ("[{0}] Executing script {1}..." -f (Get-Time),$scr.BaseName)
    # create new _Master.dwg; save it to $ProjSheetGroup.DwgPath
    & accoreconsole -s $scr
    Write-Host ("[{0}] Done creating new master dwg {1}." -f (Get-Time),$ProjSheetGroup.Name)
    return 
}
#                 </New-MasterDwg> --------------------------------------------------
#               <-RETO Initialize-ProjSheetGroup

function Initialize-ProjSheetGroup {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup
    )
    Write-Host ("[{0}] Initializing new project sheetgroup {1}..." -f (Get-Time),$ProjSheetGroup.Name)
    $null = New-MasterDwg $ProjSheetGroup
    $Project = $ProjSheetGroup.Project
    $Project.SheetGroups.Add($ProjSheetGroup)
    Write-Host ("[{0}] Done initializing new project sheetgroup {1}." -f (Get-Time),$ProjSheetGroup.Name)
    return $ProjSheetGroup
}
#               </Initialize-ProjSheetGroup> ----------------------------------------
#             <-RETO Add-ProjSheetGroup

function Add-ProjSheetGroup {
    param (
        #[PSTypeName('RevSheetGroup')]
        $RevSheetGroup,

        #[PSTypeName('Project')]
        $Project
    )
    if ($null -eq $Project) {$Project  = $RevSheetGroup.Revision.Project}
    Write-Host ("[{0}] Adding new project sheet group {1} to project {2}..." -f (Get-Time),$RevSheetGroup.Name,$Project.Name)
    $ProjSheetGroup = New-ProjSheetGroup $RevSheetGroup
    $null = Initialize-ProjSheetGroup $ProjSheetGroup
    Write-Host ("[{0}] Done adding new project sheet group {1} to project {2}." -f (Get-Time),$RevSheetGroup.Name,$Project.Name)
    return $ProjSheetGroup
}
#             </Add-ProjSheetGroup> -------------------------------------------------
#           <-RETO Get-ProjSheetGroup

function Get-ProjSheetGroup {
    param (
        #[PSTypeName('RevSheetGroup')]
        $RevSheetGroup,

        #[PSTypeName('Project')]
        $Project
    )
    Write-Host ("[{0}] Assigning revision sheetgroup {1} to a project sheetgroup..." -f (Get-Time),$RevSheetGroup.Name)
    if ($null -eq $Project) {$Project = $RevSheetGroup.Revision.Project}
    # try to get existing ProjSheetGroup from Project...
    $ProjSheetGroup = $Project.SheetGroups | Where-Object {$_.Name -eq $RevSheetGroup.Name}
    # if there is no matching ProjSheetGroup in the Project, add it.
    if ($null -eq $ProjSheetGroup) {
        Write-Host ("[{0}] Project sheetgroup {1} does not exist yet." -f (Get-Time),$RevSheetGroup.Name)
        $ProjSheetGroup = Add-ProjSheetGroup -RevSheetGroup $RevSheetGroup
    }
    Write-Host ("[{0}] Done assigning revision sheetgroup {1} to a project sheetgroup." -f (Get-Time),$RevSheetGroup.Name)
    return $ProjSheetGroup
}
#           </Get-ProjSheetGroup> ---------------------------------------------
#         <-RETO Publish-Revision
#         ->GOTO Update-ProjSheetGroup


#         ->FROM Publis-Revision
#           <Publish-Sheet> ---------------------------------------------------

#           ->FROM Publish-Sheet
#             <Add-Xref> ---------------------------------------------------

#             ->FROM Add-Xref
#               <Write-Script_AddXref> -------------------------------------
function Write-Script_AddXref {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup,

        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Writing script to add new xref {1} to master dwg {2}..." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    $ScrPathName = ("AddXref_{0}.scr" -f $Sheet.Name)
    $scr = New-Item `
        -ItemType File `
        -Path $Sheet.Project.ScrDir `
        -Name $ScrPathName `
        -Value (Get-Content $AddXrefScr -Raw
            ).Replace('_SheetDwgPath_',$Sheet.ProjDwgPath
            ).Replace('_SheetBaseName_',$Sheet.BaseName
            ).Replace('_SheetName_',$Sheet.Name
            ).Replace('_MasterDwgPath_',$ProjSheetGroup.DwgPath
            )
    Write-Host ("[{0}] Done writing script to add new xref {1} to master dwg {2}." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    return $scr
}
#               </Write-Script_AddXref> ------------------------------------
#             <-RETO Add-Xref

function Add-Xref {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup,

        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Adding new xref {1} to master dwg {2}..." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    $dwg = $ProjSheetGroup.DwgPath
    $scr = Write-Script_AddXref $ProjSheetGroup $Sheet
    Write-Host ("[{0}] Executing script {1}..." -f (Get-Time),$scr.BaseName)
    & accoreconsole -i $dwg -s $scr
    Write-Host ("[{0}] Done adding new xref {1} to master dwg {2}." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    return $dwg
}
#             </Add-Xref> --------------------------------------------------
#           <-RETO Publish-Sheet


#           ->FROM Publish-Sheet
#             <Update-Xref> ---------------------------------------------------

#             ->FROM Update-Xref
#               <Write-Script_UpdateXref> -------------------------------------
function Write-Script_UpdateXref {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup,

        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Writing script to update existing xref {1} in master dwg {2}..." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    $ScrPathName = ("UpdateXref_{0}.scr" -f $Sheet.BaseName)
    $scr = New-Item `
        -ItemType File `
        -Path $Sheet.Project.ScrDir `
        -Name $ScrPathName `
        -Value (Get-Content $UpdateXrefScr -Raw
            ).Replace('_XrefName_',$Sheet.Name
            ).Replace('_XrefPath_',$Sheet.ProjDwgPath
#           ).Replace('_MasterDwgPath_',$ProjSheetGroup.DwgPath
            )
    Write-Host ("[{0}] Done writing script to update existing xref {1} in master dwg {2}." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    return $scr
}
#               </Write-Script_UpdateXref> ------------------------------------
#             <-RETO Update-Xref

function Update-Xref {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup,

        #[PSTypeName('Sheet')]
        $Sheet
    )
    Write-Host ("[{0}] Updating existing xref {1} in master dwg {2}..." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    $dwg = $ProjSheetGroup.DwgPath
    $scr = Write-Script_UpdateXref $ProjSheetGroup $Sheet
    Write-Host ("[{0}] Executing script {1}..." -f (Get-Time),$scr.BaseName)
    & accoreconsole -i $dwg -s $scr
    Write-Host ("[{0}] Done updating existing xref {1} in master dwg {2}." -f (Get-Time),$Sheet.Name,$ProjSheetGroup.Name)
    return $dwg
}
#             </Update-Xref> --------------------------------------------------
#           <-RETO Publish-Sheet

function Publish-Sheet {
    param (
        #[PSTypeName('Sheet')]
        $NewSheet,

        #[PSTypeName('Sheet')]
        $OldSheet,

        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup
    )
    Write-Host ("[{0}] Publishing sheet {1}..." -f (Get-Time),$NewSheet.Name)
    # compose $ProjPdfPath
    $ProjPdfPath = Join-Path -Path $NewSheet.Project.PdfDir -ChildPath ($NewSheet.BaseName + '.pdf')
    # copy Sheet.df from $Sheet.RevPdfPath to $ProjPdfPath ($null to absorb return value)
    $null = Copy-Item -Path $NewSheet.RevPdfPath -Dest $ProjPdfPath
    # add $Sheet.ProjPdfPath attr, set to $ProjPdfPath
    $NewSheet | Add-Member -NotePropertyName 'ProjPdfPath' -NotePropertyValue $ProjPdfPath
    # compose $ProjDwgfPath
    $ProjDwgPath = Join-Path -Path $NewSheet.Project.DwgDir -ChildPath ($NewSheet.BaseName + '.dwg')
    # copy Sheet.dwg from $Sheet.RevDwgPath to $ProjDwgPath ($null to absorb return value)
    $null = Copy-Item -Path $NewSheet.RevDwgPath -Dest $ProjDwgPath
    # add $Sheet.ProjDwgPath attr, set to $ProjDwgPath
    $NewSheet | Add-Member -NotePropertyName 'ProjDwgPath' -NotePropertyValue $ProjDwgPath
    # update xref path in _Master.dwg for $Sheet.Name to $Sheet.ProjDwgPath
    if ($null -eq $OldSheet) {$null = Add-Xref -ProjSheetGroup $ProjSheetGroup -Sheet $NewSheet
    } else {$null = Update-Xref -ProjSheetGroup $ProjSheetGroup -Sheet $NewSheet}
    # add $ProjSheetGroup to Sheet
    $NewSheet | Add-Member -NotePropertyName 'ProjSheetGroup' -NotePropertyValue $ProjSheetGroup
    # add $Sheet to $ProjectSheetGroup
    $ProjSheetGroup.Current.Add($NewSheet)
    Write-Host ("[{0}] Done publishing sheet {1}." -f (Get-Time),$NewSheet.Name)
return $NewSheet
}
#           </Publish-Sheet> --------------------------------------------------
#         <-RETO Publish-Revision
#         ->GOTO Unpublish-Sheet


#         ->FROM Publish-Revision
#           <Unpublish-Sheet> -------------------------------------------------
function Unpublish-Sheet {
    param (
        #[PSTypeName('Sheet')]
        $OldSheet,

        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup
    )
    Write-Host ("[{0}] Archiving sheet {1}..." -f (Get-Time),$OldSheet.Name)
    $PdfArchivePath = Join-Path `
        -Path $ProjSheetGroup.Project.PdfArchiveDir `
        -ChildPath ($OldSheet.BaseName + '.pdf')
    $null = Move-Item -Path $OldSheet.ProjPdfPath -Dest $PdfArchivePath
    $OldSheet.ProjPdfPath = $PdfArchivePath
    $DwgArchivePath = Join-Path `
        -Path $ProjSheetGroup.Project.DwgArchiveDir `
        -ChildPath ($OldSheet.BaseName + '.dwg')
    $null = Move-Item -Path $OldSheet.ProjDwgPath -Dest $DwgArchivePath
    $OldSheet.ProjDwgPath = $DwgArchivePath
    $ProjSheetGroup.Current.Remove($OldSheet)
    $ProjSheetGroup.Archive.Add($OldSheet)
    Write-Host ("[{0}] Done archiving sheet {1}." -f (Get-Time),$OldSheet.Name)
    return $OldSheet
}
#           </Unpublish-Sheet> ------------------------------------------------
#         <-RETO Publish-Revision
#         ->GOTO Update-MasterPdf


#         ->FROM Publish-Revision
#           <Update-MasterPdf> ------------------------------------------------
function Update-MasterPdf {
    param (
        #[PSTypeName('ProjSheetGroup')]
        $ProjSheetGroup,

        #[PSTypeName('RevSheetGroup')]
        $RevSheetGroup
    )
    Write-Host ("[{0}] Updating master pdf {1}..." -f (Get-Time),$ProjSheetGroup.Name)
    $PSG = $ProjSheetGroup
    $RSG = $RevSheetGroup
    if ($null -eq $PSG.PdfPath) {
        # this is the first rev of this psg, need to init vars
        $PdfPath = Join-Path -Path $PSG.Project.PdfDir -ChildPath ("_{0}.pdf" -f $PSG.Name)
        $PSG | Add-Member -NotePropertyName 'PdfPath' -NotePropertyValue $PdfPath
    } else {
        # this is not the first rev of this psg, need to archive old pdf
        $ArchivePath = Join-Path `
            -Path $PSG.Project.PdfArchiveDir `
            -ChildPath ( "_{0}_{1}_{2}.pdf" -f $PSG.Name,$PSG.Revision.Name,$PSG.Revision.Date )
        $null = Move-Item -Path $PSG.PdfPath -Dest $ArchivePath
    }
    $PSG.Revision = $RSG.Revision
    $PSG.Current | ForEach-Object {$_.ProjPdfPath} | Merge-Pdf -OutputPath $PSG.PdfPath
    Write-Host ("[{0}] Done updating master pdf {1}." -f (Get-Time),$ProjSheetGroup.Name)
    return $PSG
}
#           </Update-MasterPdf> -----------------------------------------------
#         <-RETO Publish-Revision

function Publish-Revision {
    param(
        #[PSTypeName('Revision')]
        $Revision
    )
    Write-Host ("[{0}] Publishing revision {1}..." -f (Get-Time),$Revision.Name)
    foreach ($RevSheetGroup in $Revision.SheetGroups) {
        $ProjSheetGroup = Get-ProjSheetGroup $RevSheetGroup
        foreach ($NewSheet in $RevSheetGroup.Sheets) {
            $OldSheet = $ProjSheetGroup.Current | Where-Object {$_.Name -eq $NewSheet.name}
            $null = Publish-Sheet -NewSheet $NewSheet -OldSheet $OldSheet -ProjSheetGroup $ProjSheetGroup
            if ($null -ne $OldSheet) {
                $OldSheet | Add-Member -NotePropertyName 'SupersededBy' -NotePropertyValue $NewSheet
                $null = Unpublish-Sheet $OldSheet $ProjSheetGroup
            }
        }
        $null = Update-MasterPdf $ProjSheetGroup $RevSheetGroup
        Write-Host ("[{0}] Done publishing revision {1}." -f (Get-Time),$Revision.Name)
    }
    return $Revision
}
#         </Publish-Revision> -----------------------------------------------------
#       <-RETO Add-Revision

function Add-Revision {
    param (
        #[PSTypeName('Project')]
        $Project,

        #[Parameter(mandatory)]
        $RevName,

        #[Parameter(mandatory)]
        $RevDate
    )
    if ($null -eq $RevDate) {
        $RevName = Read-Host "Please enter a name for the revision"
    }
    if ($null -eq $RevDate) {
        $RevDate = Read-Host "Please enter the date of the revision, in the format YYYY-MM-DD"
    }
    Write-Host ("[{0}] Adding revision {1} to project {2}..." -f (Get-Time),$RevName,$Project.Name)
    $null = Write-Host ("You will now be prompted to select the revision PDF file.")
    $null = Read-Host ("Press Enter to continue...")
    $PdfPath = Get-UserFile -WarningAction SilentlyContinue
    $null = Write-Host ("You will now be prompted to select the sheet names TXT file.")
    $null = Read-Host ("Press Enter to continue...")
    $TxtPath = Get-UserFile -WarningAction SilentlyContinue
    # $PdfPath = Read-Host "Please enter the full path to the revision Pdf file"
    # $TxtPath = Read-Host "Please enter the full path to the sheet names Txt file"
    $Revision = New-Revision -Project $Project -RevName $RevName -RevDate $RevDate
    $null = Initialize-Revision -Revision $Revision -PdfPath $PdfPath -TxtPath $TxtPath
    $null = Publish-Revision -Revision $Revision
    # save $Project to ~.projxml
    $null = Export-Project $Project
    Write-Host ("[{0}] Done adding revision {1} to project {2}." -f (Get-Time),$RevName,$Project.Name)
    return $Revision
}
#       </Add-Revision> -------------------------------------------------------
#     <-RETO Show-ProjectMenu




#     ->FROM Show-ProjectMenu
#       <Add-Project> ---------------------------------------------------------

#       ->FROM Add-Project
#         <New-Project> -------------------------------------------------------------
function New-Project {
    param (
        #[Parameter(mandatory)]
        $Name
    )
    Write-Host ("[{0}] Creating new project {1}..." -f (Get-Time),$Name)
    $ProjDir = $pwd.ToString()
    $Project = [PSCustomObject]@{
        # Type declaration
        PSTypeName = 'Project'
        # Project attrs
        Name = $Name
        # Project aggts
        Revisions = {@()}.Invoke()
        SheetGroups = {@()}.Invoke()
        # Paths
        ProjDir = $ProjDir
        #RevDir = $RevDir
        #PdfDir = $PdfDir
        #DwgDir = $DwgDir
        #ScrDir = $ScrDir
        #ProjFile = (Join-Path $ProjDir ($Name + '.projxml'))
    }
    Write-Host ("[{0}] Done creating new project {1}." -f (Get-Time),$Name)
    return $Project
}
#         </New-Project> ------------------------------------------------------------
#       <-RETO Add-Project
#       ->GOTO Initialize-Project


#       ->FROM Add-Project
#         <Initialize-Project> ------------------------------------------------------

#         ->FROM Initialize-Project
#           <Show-AskAddRevision> ---------------------------------------------------------
function Show-AskAddRevision {
    param (
        #[PSTypeName('Project')]
        $Project
    )
    $AskAddRevision = Get-UserInput `
        @("Project created: {0}" -f $Project.Name) `
        "Would you like to add a Revision to the new Project?" `
        @('y','n')
    switch ($AskAddRevision) {
        'y' {Add-Revision $Project}
        'n' {Write-Host ("OK. Returning to Project Menu...")}
    }
    return
}
#           </Show-AskAddRevision> --------------------------------------------------------
#         <-RETO Initialize-Project

function Initialize-Project {
    param (
        #[PSTypeName('Project')]
        $Project
    )
    Write-Host ("[{0}] Initializing new project {1}..." -f (Get-Time),$Project.Name)
    # add $Project.RevDir attr
    $RevDir = (Join-Path -Path $Project.ProjDir -ChildPath '_Revisions')
    if ((Test-Path $RevDir) -ne $true) {$null = mkdir $RevDir}
    $Project | Add-Member -NotePropertyName 'RevDir' -NotePropertyValue $RevDir
    # add $Project.PdfDir attr
    $PdfDir = (Join-Path -Path $Project.ProjDir -ChildPath '_CurrentPdfs')
    if ((Test-Path $PdfDir) -ne $true) {$null = mkdir $PdfDir}
    $Project | Add-Member -NotePropertyName 'PdfDir' -NotePropertyValue $PdfDir
    # add $Project.PdfArchiveDir attr
    $PdfArchiveDir = (Join-Path -Path $Project.PdfDir -ChildPath '_Archive')
    if ((Test-Path $PdfArchiveDir) -ne $true) {$null = mkdir $PdfArchiveDir}
    $Project | Add-Member -NotePropertyName 'PdfArchiveDir' -NotePropertyValue $PdfArchiveDir
    # add $Project.DwgDir attr
    $DwgDir = (Join-Path -Path $Project.ProjDir -ChildPath '_CurrentDwgs')
    if ((Test-Path $DwgDir) -ne $true) {$null = mkdir $DwgDir}
    $Project | Add-Member -NotePropertyName 'DwgDir' -NotePropertyValue $DwgDir
    # add $Project.DwgArchiveDir attr
    $DwgArchiveDir = (Join-Path -Path $Project.DwgDir -ChildPath '_Archive')
    if ((Test-Path $DwgArchiveDir) -ne $true) {$null = mkdir $DwgArchiveDir}
    $Project | Add-Member -NotePropertyName 'DwgArchiveDir' -NotePropertyValue $DwgArchiveDir
    # add $Project.ScrDir attr
    $ScrDir = (Join-Path -Path $Project.ProjDir -ChildPath '_Scripts')
    if ((Test-Path $ScrDir) -ne $true) {$null = mkdir $ScrDir}
    $Project | Add-Member -NotePropertyName 'ScrDir' -NotePropertyValue $ScrDir
    # add $Project.Projfile attr
    $ProjFile = (Join-Path -Path $Project.ProjDir -ChildPath ($Project.Name + '.projxml'))
    $Project | Add-Member -NotePropertyName 'ProjFile' -NotePropertyValue $ProjFile
    Write-Host ("[{0}] Done initializing new project {1}." -f (Get-Time),$Project.Name)
    #Show-AskAddRevision $Project
    return $Project
}
#         </Initialize-Project> -----------------------------------------------------
#       <-RETO Add-Project

function Add-Project {
    param (
        $Name
    )
    if ($null -eq $Name) {$Name = Read-Host "Please enter a name for the project"}
    Write-Host ("[{0}] Adding new project {1}..." -f (Get-Time),$Name)
    $Project = New-Project $Name
    $null = Initialize-Project $Project
    # save $Project to ~.projxml
    $null = Export-Project $Project
    Write-Host ("[{0}] Done adding new project {1}." -f (Get-Time),$Name)
    return $Project
}
#       </Add-Project> --------------------------------------------------------
#     <-RETO Show-ProjectMenu




#     ->FROM Show-ProjectMenu
#       <Show-ProjectDetails> -----------------------------------------------------------
function Show-ProjectDetails {
    param (
        $Project
    )
    Write-Host
    return $Project
}
#       </Show-ProjectDetails> ----------------------------------------------------------
#     <-RETO Show-ProjectMenu




#     ->FROM Show-ProjectMenu
#       <Get-Project> -----------------------------------------------------------
#       ->FROM Get-Project
#         <Resolve-NoExistingProjects> ------------------------------------------
function Resolve-NoExistingProjects {
    $AskNewProject = Get-UserInput `
        -InitStatements @(,"No Projects found.") `
        -MainQuestion "Would you like to create a new one?" `
        -Options @('y','n')
    switch ($AskNewProject) {
        'y' {
            $Project = Add-Project
        }
        'n' {
            Exit-Project
        }
    }
    return $Project
}
#         </Resolve-NoExistingProjects> -----------------------------------------
#       <-RETO Get-Project

#       ->FROM Get-Project
#         <Resolve-OneExistingProject> ------------------------------------------
function Resolve-OneExistingProject {
    param (
        [System.Object[]]
        $ProjFile
    )
    Write-Host ("Found 1 Project:")
    Write-Host ("`r`n{0}`r`n" -f $ProjFile.BaseName)
    $Project = Import-Project $ProjFile
    return $Project
}
#         </Resolve-OneExistingProject> -----------------------------------------
#       <-RETO Get-Project

#       ->FROM Get-Project
#         <Resolve-ManyExistingProjects> ----------------------------------------
function Resolve-ManyExistingProjects {
    param (
        [System.Object[]]
        $ProjFiles
    )
    $ProjNameList = $ProjFiles | ForEach-Object{$_.BaseName}
    $AskWhichProjFile = Get-UserInput `
        @("Found {0} Projects:`r`n" -f $ProjFiles.Count,
            "`r`n{0}`r`n" -f ([String]::Join("`r`n",$ProjNameList))
        ) `
        "Which one would you like to open?" `
        @(1..($ProjNamesList.Count))
    $ProjFile = $ProjNamesList[([int]$AskWhichProjFile - 1)]
    $Project = Import-Project $ProjFile
    return $Project
}
#         </Resolve-ManyExistingProjects> ---------------------------------------
#       <-RETO Get-Project

function Get-Project {
    param (
        $ProjDir = $PWD
    )
    # get existing projects in $PWD
    Write-Host ("Looking for projects in {0}..." -f $ProjDir)
    $ProjFiles = Get-ChildItem -Path $ProjDir -Filter '*.projxml'
    # if there are none, ask to create a new one
    switch ($ProjFiles.Count)
    {
        0 {$Project = Resolve-NoExistingProjects} # New-Project or Exit
        1 {$Project = Resolve-OneExistingProject $ProjFiles}
        Default {$Project = Resolve-ManyExistingProjects $ProjFiles}
    }
    return $Project
}
#       </Get-Project> ----------------------------------------------------------
#     <-RETO Show-ProjectMenu




function Show-ProjectMenu {
    param (
        $Project
    )
    if ($null -eq $Project) {$Project = Get-Project}
    $ActionList = @(
        "d = show project Details"
        "r = add a new Revision to this project"
        "o = save/close this project and Open a different one"
        "n = save/close this project and create a New one"
        "x = save/close this project and eXit"
    )
    $ActionKeys = @(
        "d"
        "r"
        "o"
        "n"
        "x"
    )
    $UserAction = Get-UserInput `
        @("Available actions:",
          "`r`n{0}`r`n" -f ([String]::Join("`r`n",$ActionList))
        ) `
        "What would you like to do?" `
        $ActionKeys
    switch($Useraction) {
        'd' {Show-ProjectDetails $Project}
        'r' {Add-Revision $Project}
        'o' {Export-Project $Project; Import-Project}
        'n' {Export-Project $Project; New-Project}
        'x' {Export-Project $Project; Exit-Project}
    }
    Show-ProjectMenu $Project
}
#     </Show-ProjectMenu> ---------------------------------------------------------
#   <-RETO Main




function Invoke-Main {
    Find-AutoCadCoreConsole
    Show-ProjectMenu
    return $true
}
#   </Main> -------------------------------------------------------------------
# <-RETO Entry-Point

Invoke-Main

# </Entry-Point> --------------------------------------------------------------
# TERM