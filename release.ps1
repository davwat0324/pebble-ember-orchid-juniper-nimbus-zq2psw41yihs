<#
  Publishes a new version of Remind.

    .\release.ps1 1.2.0 "What changed, in a sentence or two"

  Sets src\version.py to the new version, commits it, tags v<version> and pushes. The push of the tag
  starts the GitHub Actions workflow (.github\workflows\release.yml), which builds Remind-win64.zip and
  attaches it to a GitHub Release. Installed copies of Remind see the new release the next time they
  start (or from Settings > Check for updates) and offer to update themselves.
#>
param(
    [Parameter(Mandatory)][ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version,
    [string]$Notes = ''
)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (git status --porcelain) { throw 'Commit or stash your changes first (git status is not clean).' }
$vf = 'src\version.py'
$text = Get-Content $vf -Raw
$current = [regex]::Match($text, '__version__ = "([^"]+)"').Groups[1].Value
if ([version]$Version -le [version]$current) { throw "New version $Version must be higher than $current." }

$text = $text -replace '__version__ = "[^"]+"', "__version__ = `"$Version`""
[IO.File]::WriteAllText((Resolve-Path $vf), $text, [Text.UTF8Encoding]::new($false))

git add $vf
git commit -m "Release v$Version" | Out-Null
$msg = if ($Notes) { $Notes } else { "Remind $Version" }
git tag -a "v$Version" -m $msg
git push origin HEAD --follow-tags
Write-Host "Pushed v$Version. GitHub Actions is building the release:"
Write-Host "  https://github.com/$((git remote get-url origin) -replace '.*github.com[:/]','' -replace '\.git$','')/actions"
