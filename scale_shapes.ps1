###########################################################################
# Script: create_stencil.ps1
# Author: Adam Newhard
# URL: https://github.com/KiPIDesTAN/visio-stencil-maker
#
# Description:
# The script below iterates through a folder of images and creates a stencil
# from all those images. I use this frequently for architecture diagrams.
# 
# The script does the following:
# 1. Gets a list of all images in a given folder
# 2. Imports the image as a shape
# 2. Scales the shape down to an appropriate size
# 3. Creates a series of consistent connector points on the shape
# 4. Adjusts the shape's text location and applies text to the shape
# 5. Adds the shape to the master stencil
# 6. Saves the final file as a stencil that can be opened in Visio for use in other files
#
# NOTE: Sometimes Visio does not do a good job of converting the SVG to PNG. In fact, some SVG
# to PNG programs have a tough time. The most accurate solution I've found is to run rsvg-convert,
# a Linux application, to convert the svg to png.
#
# rsvg-convert input.svg -o output.png
#
###########################################################################

# Variables
# When resizing the svg, this is the maximum dimension for the image.
# The maxDimension is applied to the larger of height and width.
$maxDimension = 20
$fileTypeList = ".png", ".svg" # A list of image extensions to look for

# Location of the JSON file defining the connection points for the object
$connectionPointFile = ".\shape_connectors.json"
# Define the build directory
$imageDir = "C:\Users\adamnewhard\OneDrive - Microsoft\Documents\Visio\adam_shapes"
$savePath = $(Get-Location | Select-Object -ExpandProperty Path) + "\mystencil.vssx"

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

# Load the connection point definitions
$connectionPoints = Get-Content $connectionPointFile | ConvertFrom-Json

# Get a sorted list of SVG files
$svgFiles = Get-ChildItem -Path $imageDir | Where-Object { $_.extension -in $fileTypeList } | Sort-Object Name

 # Create a new Visio document
$doc = $visioApp.Documents.Add("")

$page = $doc.Pages.Item(1) # Uncomment when working with a document

# Get the stencil's master shape collection
$masters = $doc.Masters

foreach ($svgFile in $svgFiles) {

    Write-Output "Processing $($svgFile.Name)."

    # Drop the SVG onto the page
    $shape = $page.Import($svgFile.FullName)

    # Set the text applied for the shape.
    $shapeText = $($svgFile.BaseName)

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
    $string = $($svgFile.BaseName)
    $shapeText = ($string.Split(' ') | ForEach-Object { $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower() }) -join " "

    # Add Shape sheet placeholders
    $shape.AddNamedRow($visSectionProp, "Label", $visTagDefault)
    $shape.CellsSRC($visSectionProp, $rowIndex, $visCellPropValue).FormulaU = "`"$($svgFile.Name)`""
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

# Save the stencil when ready
$doc.SaveAs($savePath)

# Close the document
$doc.Close()

