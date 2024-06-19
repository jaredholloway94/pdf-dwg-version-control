if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    { Start-Process powershell.exe " -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
Set-Location (Join-Path $HOME ".pdfvc")
    # bypass restricted execution policy, so that we can...
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
# install MergePdf PS module from PSGallery
Install-Module MergePdf
# install pdftk
& ./pdftk_free-2.02-win-setup.exe
# add the $User/.pdfvc path to $Path, so it can be found from anywhere
& ./append_user_path.cmd (Join-Path $HOME ".pdfvc")
