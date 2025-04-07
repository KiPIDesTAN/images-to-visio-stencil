###########################################
#   Title: Images to Visio Stencil
#   Description: This script takes a directory of images and imports them into a Visio Stencil file.
#   URL: https://github.com/KiPIDesTAN/images-to-visio-stencil
#
#   NOTE: There is absolutely NO WARRANTY with this code. It is something I wrote
#   for my architecture work and am sharing in the event someone else can use it.
#   It is inspired by the work at https://github.com/David-Summers/Azure-Design.
#   
###########################################

# Define the directory where images should be acquired from
$buildDir = "<path to images here>"

# Variables
# When resizing the svg, this is the maximum dimension for the image.
# The maxDimension is applied to the larger of height and width.
$maxDimension = 20

# Location of the JSON file defining the connection points for the object.
$connectionPointFile = ".\shape_connectors.json"

# Create the stencil name as the parent directory name with the first letter of the first word capitalized.
$stencilName = $($buildDir -split '\\')[-1] -replace '^(.)', { $_.Value.ToUpper() }

# Define Visio constants
# Indices defined here - https://learn.microsoft.com/en-us/office/vba/api/visio.vissectionindices
$visSectionProp = 243
$visSectionConnectionPts = 7    # Connection points section
$visTagDefault = 0

# Load Visio COM object
$visioApp = New-Object -ComObject Visio.Application
$visioApp.Visible = $true


# Load the connection point definitions
$connectionPoints = Get-Content $connectionPointFile | ConvertFrom-Json

# Get a sorted list of SVG files
$svgFiles = Get-ChildItem -Path $buildDir -Filter *.svg | Sort-Object Name

 # Create a new Visio document
$doc = $visioApp.Documents.Add("")

$page = $doc.Pages.Item(1) # Uncomment when working with a document

# Get the stencil's master shape collection
$masters = $doc.Masters

foreach ($svgFile in $svgFiles) {

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
    $shape.CellsSRC($visSectionProp, $rowIndex, 0).FormulaU = "`"$($shapeText)`""

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

    # # Close the document
    # $doc.Close()
}

# Save the stencil when ready
$doc.SaveAs($buildDir + "\$($stencilName).vssx")

