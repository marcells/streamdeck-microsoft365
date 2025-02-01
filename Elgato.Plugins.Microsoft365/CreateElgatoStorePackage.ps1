Write-Host "Script root: $PSScriptRoot`n"

$basePath = $PSScriptRoot

if ($PSSCriptRoot.Length -eq 0) {
  $basePath = $PWD.Path;
}

# Load and parse the plugin project file
$pluginProjectFile = "$basePath\Elgato.Plugins.Microsoft365.csproj"
$projectContent = Get-Content $pluginProjectFile | Out-String;
$projectXML = [xml]$projectContent;

$buildConfiguration = "Debug"

# Get the target .net core framework
$targetFrameworkName = $projectXML.Project.PropertyGroup.TargetFramework;

# For now, this PS script will only be run on Windows.
$bindir = "$basePath\bin\Debug\$targetFrameworkName\win-x64"

# Make sure we actually have a directory/build to deploy
If (-not (Test-Path $bindir)) {
  Write-Error "The output directory `"$bindir`" was not found.`n You must first build the `"Elgato.Plugins.Microsoft365`" project before calling this script.";
  exit 1;
}

# Load and parse the plugin's manifest file
$manifestFile = $bindir +"\manifest.json"
$manifestContent = Get-Content $manifestFile | Out-String
$json = ConvertFrom-JSON $manifestcontent

$pluginID = $json.UUID
$baseDestinationDir = "$($env:APPDATA)\Elgato\StreamDeck\Plugins"
$pluginDirectory = "$baseDestinationDir\$pluginID.sdPlugin"

Write-Host "Creating package of '$pluginDirectory' is complete..."

#######################################
# if command is not found, install manually first: npm install -g @elgato/cli
# see docs: https://docs.elgato.com/streamdeck/sdk/introduction/getting-started/#setup-wizard
#######################################
streamdeck pack $pluginDirectory --force --output $baseDestinationDir

Write-Host "Creating package of '$pluginDirectory' is complete..."
Write-Host "Package is found at '$baseDestinationDir'..."

exit 0
