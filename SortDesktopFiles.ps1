<#
    .SYNOPSIS
        Sorts desktop items by extension
    
    .DESCRIPTION
        Each item on desktop is moved to a directory matching the file extension.
    
    .AUTHOR 
        Clint@ClintLang.com
#>

Set-Location -Path ([System.Environment]::GetFolderPath("Desktop"));
$files = [System.IO.DirectoryInfo]::new((Get-Location)).GetFiles("*");

if ($files.Count -lt 1) { EXIT; }

foreach($file in $files) {

    $ext     = $file.Extension.Replace(".", "").Trim();    
    $fn      = $file.Name;
    $flname  = $file.FullName;
    $dest    = Join-Path (Get-Location) $ext;
    $target  = Join-Path $dest $fn;
    $destExists = Test-Path -Path $dest;

    try {

        if (!$destExists) {
            Write-Host "`r`nCreate new directory $ext...." -ForegroundColor Cyan -NoNewline;
            New-Item -Path $dest -ItemType Directory -Force -Confirm:$false | Out-Null;
            Write-Host "Path created" -ForegroundColor Green -BackgroundColor Black;

        }
        
        Write-Host "`r`nMoving file $fn...." -ForegroundColor Cyan -NoNewline;
        Move-Item -Path $flname -Destination $target -Force -Confirm:$false | Out-Null;
        Write-Host "Done" -ForegroundColor Green -BackgroundColor Black;
    }

    catch
    {
        Write-Host " ERROR  " -BackgroundColor Black -ForegroundColor Yellow;
        Write-Host $_.Exception.Message;
    }
}

Write-Host "`r`n`r`n---------------------------------------------------------";
pause;

EXIT;