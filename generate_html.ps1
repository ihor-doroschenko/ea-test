# PowerShell script to generate HTML from Enterprise Architect
param (
    [string]$eaFile = "D:\Code\EA\Test.qea",  # Path to your EA file
    [string]$outputPath = "D:\Code\EA\test"  # Path for HTML output
)

# Create an instance of the EA.Repository COM object
try {
    $eaRepo = New-Object -ComObject EA.Repository
} catch {
    Write-Error "Failed to initialize EA.Repository COM object."
    exit 1
}

Write-Output "EA.Repository initialized."

# Open the EA file
try {
    $eaRepo.OpenFile($eaFile)
    Write-Output "Opened EA file: $eaFile"
} catch {
    Write-Error "Failed to open EA file: $eaFile"
    $eaRepo.Exit()
    exit 1
}

# Generate HTML for the first package in the model
try {
    foreach ($model in $eaRepo.Models) {
        Write-Output "Model Name: $($model.Name)"
        $package = $eaRepo.GetPackageByID(1)  # Replace with the actual package ID
        if ($package) {
            Write-Output "Found package: $($package.Name), GUID: $($package.PackageGUID)"
            $eaRepo.GetProjectInterface().RunHTMLReport($package.PackageGUID, $outputPath, "GIF", "<default>", ".html")
            Write-Output "HTML report generated at: $outputPath"
            break
        } else {
            Write-Error "Package with ID 1 not found."
        }
    }
} catch {
    Write-Error "Failed to generate HTML report."
}

# Close the EA file and exit
$eaRepo.CloseFile()
$eaRepo.Exit()
Write-Output "EA.Repository closed."
