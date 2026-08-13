[CmdletBinding()]
param(
	# A directory whose immediate subdirectories are OneNote notebook backups.
	[Parameter(Mandatory, Position = 0)]
	[ValidateNotNullOrEmpty()]
	[string]$BackupRoot,

	# A new candidate directory. This runner never replaces an existing path.
	[Parameter(Mandatory, Position = 1)]
	[ValidateNotNullOrEmpty()]
	[string]$OutputDirectory,

	# Optional previous audit JSON for a numeric comparison in the new audit.
	[string]$PreviousAudit,

	# Defaults to a sibling of the candidate, never inside it.
	[string]$AuditReport,

	# Lists the selected source notebooks and files without writing anything.
	[switch]$PlanOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$oneNoteExtensions = @('.one', '.onepkg', '.onex')

function Resolve-AbsolutePath([string]$Value) {
	return [System.IO.Path]::GetFullPath($Value)
}

function Get-OneNoteFiles([string]$Directory) {
	return @(Get-ChildItem -LiteralPath $Directory -File | Where-Object {
		$oneNoteExtensions -contains $_.Extension.ToLowerInvariant()
	} | Sort-Object Name)
}

$sourceRoot = Resolve-AbsolutePath $BackupRoot
if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
	throw "Backup root is not a directory: $sourceRoot"
}

$output = Resolve-AbsolutePath $OutputDirectory
if (Test-Path -LiteralPath $output) {
	throw "Candidate output already exists; choose a fresh path: $output"
}

$notebooks = @(
	Get-ChildItem -LiteralPath $sourceRoot -Directory |
		ForEach-Object {
			$files = Get-OneNoteFiles $_.FullName
			if ($files.Count -eq 0) { return }
			[pscustomobject]@{
				Name = $_.Name
				Directory = $_.FullName
				Files = $files
			}
		} |
		Sort-Object Name
)

if ($notebooks.Count -eq 0) {
	throw "No notebook subdirectory containing a .one, .onepkg, or .onex file was found under: $sourceRoot"
}

Write-Host "Selected $($notebooks.Count) notebook backup directories:"
foreach ($notebook in $notebooks) {
	Write-Host "  $($notebook.Name) ($($notebook.Files.Count) backup files)"
	foreach ($file in $notebook.Files) {
		Write-Host "    $($file.Name)"
	}
}

if ($PlanOnly) { return }

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$tsx = Join-Path $repositoryRoot 'node_modules\.bin\tsx.cmd'
if (-not (Test-Path -LiteralPath $tsx -PathType Leaf)) {
	throw "tsx is not installed at $tsx. Run pnpm install using the pinned lockfile first."
}

$exportScript = Join-Path $PSScriptRoot 'export-onenote-backups.ts'
$auditScript = Join-Path $PSScriptRoot 'audit-onenote-output.ts'
$staging = "${output}-staging"
if (Test-Path -LiteralPath $staging) {
	throw "Refusing to overwrite existing staging history: $staging"
}

$exportArguments = @('--disable-warning=ExperimentalWarning', '--tsconfig', 'tsconfig.test.json', $exportScript, $output) + @($notebooks | ForEach-Object Directory)
& $tsx @exportArguments
if ($LASTEXITCODE -ne 0) {
	throw "OneNote backup export failed with exit code $LASTEXITCODE. The candidate and any staging output were retained for inspection."
}

$stagingInside = Join-Path $output '_staging'
if (-not (Test-Path -LiteralPath $stagingInside -PathType Container)) {
	throw "Export completed without its expected staging directory: $stagingInside"
}
Move-Item -LiteralPath $stagingInside -Destination $staging

if ([string]::IsNullOrWhiteSpace($AuditReport)) {
	$AuditReport = "${output}-audit.json"
}
$audit = Resolve-AbsolutePath $AuditReport
if (Test-Path -LiteralPath $audit) {
	throw "Refusing to overwrite existing audit report: $audit"
}
$auditParent = Split-Path -Parent $audit
if (-not (Test-Path -LiteralPath $auditParent -PathType Container)) {
	New-Item -ItemType Directory -Path $auditParent | Out-Null
}

$auditArguments = @('--disable-warning=ExperimentalWarning', '--tsconfig', 'tsconfig.test.json', $auditScript, $output) + @($notebooks | ForEach-Object Directory)
if ($PreviousAudit) {
	$auditArguments += "--previous=$(Resolve-AbsolutePath $PreviousAudit)"
}
$auditJson = & $tsx @auditArguments
$auditExitCode = $LASTEXITCODE
$auditJson | Set-Content -LiteralPath $audit -Encoding utf8
if ($auditExitCode -ne 0) {
	throw "Candidate audit failed with exit code $auditExitCode. See $audit; the candidate and staging were retained."
}

$parsedAudit = Get-Content -LiteralPath $audit -Raw | ConvertFrom-Json
if ($parsedAudit.failureCount -ne 0) {
	throw "Candidate audit reported $($parsedAudit.failureCount) failures. See $audit; the candidate and staging were retained."
}

Write-Host "Candidate export passed audit: $output"
Write-Host "Staging retained outside candidate: $staging"
Write-Host "Audit report: $audit"
