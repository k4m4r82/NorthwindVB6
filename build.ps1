# This build assumes the following directory structure
# Script assumes all directories exists, initialized by a parent script
#
#  \build_artifact - This folder is created if it is missing and contains output of the build
#  \build_log      - This folder is created if it is missing and contains log of the build
#  \src            - This folder contains the source code or project you want to build

# run : .\psake.ps1 build.ps1

# properties that is used by the script
properties {        			
	$vb6bin = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
	
	$src_dir = Split-Path $psake.build_script_file
    $build_artifact_dir = "$src_dir\build_artifact"
	$build_log_dir = "$src_dir\build_log"
	
	$project_name = "Northwind.vbp"
    $project = "$src_dir\$project_name"		        
    $logfile = "$build_log_dir\$project_name.log"
}

function NormalizeOutDir($outdir) {
    if (-not ($outdir.EndsWith("\"))) {
		$outdir += '\'
	}

    if ($outdir.Contains(" ")) {
		$outdir = $outdir + "\"
	}

    return $outdir
}

function HasFailed($logFile) {
	return ((Select-String "failed" $logfile -Quiet) -or (Select-String "not found" $logfile -Quiet))
}

function HasSucceeded($logFile) {
	return (Select-String "succeeded" $logFile -Quiet)
}

FormatTaskName (("-"*25) + "[{0}]" + ("-"*25))

Task Default -Depends Build

Task Build -Depends CreateLogFile {				
    $outdir = NormalizeOutDir("$build_artifact_dir")
    $failed = $false
    $retries = 0
    $succeeded = $false
	
    Write-Host "Building $name" -ForegroundColor Green
	
	try {
		Exec {& $vb6bin /m $project /out $logfile /outdir $outdir}
	} catch {
		$failed = $true
	}

    while (!($failed -or $succeeded)) {
        Write-Host -NoNewline "."
        Start-Sleep -s 1
        $failed = HasFailed($logfile)
        $succeeded = HasSucceeded($logfile)
        $retries = $retries + 1
        $failed = ($failed -or ($retries -eq 60)) -and !$succeeded
    }

    if ($failed)
    {
        Type $logfile
        throw "Unable to build $name"
    }
    Type $logfile
}

Task CreateLogFile {
	If (!(Test-Path $build_artifact_dir)) 
	{
		New-Item -Path $build_artifact_dir -ItemType Directory
	}
	
	If (!(Test-Path $build_log_dir)) 
	{
		New-Item -Path $build_log_dir -ItemType Directory
	}
	
	Write-host  ("Src dir : {0}" -f $src_dir)
	Write-host  ("Artifact dir : {0}" -f $build_artifact_dir)	
	Write-host  ("Log dir : {0}" -f $logfile)
	
    if (Test-Path $logfile)
    {
        Remove-Item $logfile
    }
	
    $path = [IO.Path]::GetFullPath($logfile)
    New-Item -ItemType file  $path
}