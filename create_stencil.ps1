###########################################################################
# Script: create_stencil.ps1
# Author: Adam Newhard
# URL: https://github.com/KiPIDesTAN/visio-stencil-maker
#
# Description:
# The script below iterates through a folder of images and creates a stencil
# from all those images. I use this frequently for architecture diagrams.
# 
# You will need to install Inkscape.
#
# NOTE: Sometimes Visio does not do a good job of converting the SVG to PNG. In fact, some SVG
# to PNG programs have a tough time. This software uses Inkscape's CLI to convert any
# SVGs to PNGs. This allows all logos to come out correctly.
#
###########################################################################

# Variables
# When resizing the svg, this is the maximum dimension for the image.
# The maxDimension is applied to the larger of height and width. This normally
# looks good, but sometimes will look terrible if height to width ratio is
# dramatically different.
$maxDimension = 20
# A list of image extensions that are used for stencil creation.
# SVGs are converted to PNG. They are not included in this list.
$fileTypeList = ".png" 

# Location of the JSON file defining the connection points for the object
$connectionPointFile = ".\shape_connectors.json"
# Define the build directory
$rootDir = "C:\Users\adamnewhard\OneDrive - Microsoft\Documents\Visio\Logos"
# Where to save the vssx file that's created
$saveDir = Join-Path $env:USERPROFILE "Documents\My Shapes" #'C:\Users\adamnewhard\OneDrive - Microsoft\Documents\Visio\Logos'
# Path to inkscape cli, used to convert svgs to png
# & "C:\Program Files\Inkscape\bin\inkscape.com" 'C:\Users\adamnewhard\OneDrive - Microsoft\Documents\Visio\Logos\Fluentbit.svg' --export-type=png --export-filename=Fluentbit.png
$inkscapePath = 'C:\Program Files\Inkscape\bin\inkscape.com'

# Add a slash to the end if it's missing
if (-not $rootDir.EndsWith("\")) { $rootDir += "\" }
if (-not $saveDir.EndsWith("\")) { $saveDir += "\" }

# Stencil save path
$savePath = $saveDir + $(Split-Path $rootDir -Leaf) + ".vssx"   # File to save

# Load the connection point definitions
$connectionPoints = Get-Content $connectionPointFile | ConvertFrom-Json

# Get a sorted list of usable file types and svgs.
# We convert the svgs when a usable file type doesn't exist.
$stencilFiles = Get-ChildItem -Path $rootDir | Where-Object { $_.extension -eq '.svg' } | Sort-Object Name
$allFiles = Get-ChildItem -Path $rootDir | Where-Object { $_.extension -in $fileTypeList } | Sort-Object Name

foreach ($svg in $stencilFiles) {
    $fileExists = $allFiles | Where-Object { $_.BaseName -eq $svg.BaseName }
    if (-not $fileExists) {
        Write-Output "Converting $svg to png."
        $outputFilePath = $(Split-Path -Parent $svg) + "\" + $svg.BaseName + ".png"
        & $inkscapePath $svg.FullName --export-type=png --export-filename=$outputFilePath
    }
}


$stencilFiles = Get-ChildItem -Path $rootDir | Where-Object { $_.extension -in $fileTypeList } | Sort-Object Name

# Define Visio constants
# Indices defined here - https://learn.microsoft.com/en-us/office/vba/api/visio.vissectionindices
$visSectionProp = 243
$visSectionConnectionPts = 7    # Connection points section
$visTagDefault = 0

# Used for the $visSectionProp to set the lable, prompt, and value fields
$visCellPropLabel = 1
$visCellPropPrompt = 2
$visCellPropValue = 0

# Load Visio COM object
$visioApp = New-Object -ComObject Visio.Application
$visioApp.Visible = $true

# If the file already exists, open it. Otherwise, create a new file.
if (Test-Path $savePath) {
    $doc = $visioApp.Documents.Open($savePath)
} else {
 # Create a new Visio document
    $doc = $visioApp.Documents.Add("")
}

$page = $doc.Pages.Item(1) # Uncomment when working with a document

# Get the stencil's master shape collection
$masters = $doc.Masters

# Get a list of all the stencil names. We skip adding stencils tha already exist.
$masterList = $masters | % { $_.Name }

Write-Output "Reviewing stencils for new additions."

foreach ($stencilFile in $stencilFiles) {
    # If the file name is already in the Visio's stencil list, go to the next item.
    if ($stencilFile.BaseName -in $masterList) {
        continue
    }

    Write-Output "Processing $($stencilFile.Name)."

    # Drop the SVG onto the page
    $shape = $page.Import("$($stencilFile.FullName)")

    # Set the text applied for the shape.
    $shapeText = $($stencilFile.BaseName)

    # Scale the newly added image to the proper size.
    $width = $shape.Cells("Width").Result("mm")
    $height = $shape.Cells("Height").Result("mm")

    if ($width -ge $height) {
        $newWidth = $maxDimension
        $scaleFactor = $newWidth / $width
        $newHeight = $height * $scaleFactor
    } else {
        $newHeight = $maxDimension
        $scaleFactor = $newHeight / $height
        $newWidth = $width * $scaleFactor
    }

    $shape.Cells("Width").Result("mm") = $newWidth
    $shape.Cells("Height").Result("mm") = $newHeight

    # Set the text applied for the shape by sourcing the file name and setting each word to uppercase.
    $shapeText = $($stencilFile.BaseName)

    # Add Shape sheet placeholders
    $row = $shape.AddNamedRow($visSectionProp, "Label", $visTagDefault)
    $shape.CellsSRC($visSectionProp, $rowIndex, $visCellPropValue).FormulaU = "`"$($stencilFile.Name)`""
    $shape.CellsSRC($visSectionProp, $rowIndex, $visCellPropLabel).FormulaU = "`"$($shapeText)`""
    $shape.CellsSRC($visSectionProp, $rowIndex, $visCellPropPrompt).FormulaU = "`"$($shapeText)`""

    # Set the text of the shape
    $shape.Text = $shapeText

    # Add connection points
    foreach ($connectionPoint in $connectionPoints) {
        $row = $shape.AddNamedRow($visSectionConnectionPts, $connectionPoint.Name, $visTagDefault)
        $shape.CellsSRC($visSectionConnectionPts, $row, 0).FormulaU = $connectionPoint.X    # X-position
        $shape.CellsSRC($visSectionConnectionPts, $row, 1).FormulaU = $connectionPoint.Y    # Y-position
        $shape.CellsSRC($visSectionConnectionPts, $row, 2).FormulaU = "0"          # Type (0 = inward)
    }

    # Reposition the default text field to the bottom of the icon
    $shape.CellsU("TxtWidth").FormulaU = "TEXTWIDTH(TheText)"   # Find "thetext" value in the shape list
    $shape.CellsU("TxtHeight").FormulaU = "TEXTHEIGHT(TheText,TxtWidth)"
    $shape.CellsU("TxtPinX").FormulaU = "Width*0.5"  # Center horizontally
    $shape.CellsU("TxtPinY").FormulaU = "Height*-0.25"          # Move to bottom
    $shape.CellsU("TxtLocPinX").FormulaU = "TxtWidth*0.5"       # Align text to bottom
    $shape.CellsU("TxtLocPinY").FormulaU = "TxtHeight*0.5"       # Align text to bottom

    # Set font to Arial, size 8pt
    $shape.CellsU("Char.Font").FormulaU = "4"       # Font: Arial
    $shape.CellsU("Char.Size").FormulaU = "8 pt"    # Font size: 8 pt

    # Add the shape to master stencil
    $master = $masters.Add()
    $master.Name = $shapeText

    # NOTE: Copy/paste only works when the shape is of a "group" type
    $shape.Copy()
    $master.Paste()

    # Delete the shape from the doc now that it's in the Master stencil
    $shape.Delete()
}

Write-Host "Saving stencil as $savePath."
# Save the stencil when ready. No need to output the return value.
$null = $doc.SaveAs($savePath)

# Close the document
$doc.Close()

$visioApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($visioApp) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()