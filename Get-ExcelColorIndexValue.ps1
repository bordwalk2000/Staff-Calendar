function Get-ExcelColorIndexValue {
    <#
    .SYNOPSIS
        Converts RGB or Hex color values to an Excel color index value.

    .DESCRIPTION
        The Get-ExcelColorIndexValue function calculates the Excel color index value from either RGB or Hex color values.
        This value can be used to set the color of cells in Excel.

    .PARAMETER red
        The red component of the RGB color value (0-255). Mandatory in the RGB parameter set.

    .PARAMETER green
        The green component of the RGB color value (0-255). Mandatory in the RGB parameter set.

    .PARAMETER blue
        The blue component of the RGB color value (0-255). Mandatory in the RGB parameter set.

    .PARAMETER hexColor
        The hex color value (e.g., "#FF0000"). Mandatory in the Hex parameter set.

    .EXAMPLE
        PS C:\> Get-ExcelColorIndexValue -red 255 -green 0 -blue 0
        255

        This example converts the RGB color values (255, 0, 0) to an Excel color index value.

    .EXAMPLE
        PS C:\> Get-ExcelColorIndexValue 231 230 230
        15132391

        This example converts the RGB color values (231, 230, 230) to an Excel color index value.

    .EXAMPLE
        PS C:\> Get-ExcelColorIndexValue -hexColor "#00FF00"
        65280

        This example converts the hex color value ("#00FF00") to an Excel color index value.

    .NOTES
        This function uses the Excel color index formula, which combines RGB values into a single integer.
    #>

    [CmdletBinding(
        DefaultParameterSetName = "RGB"
    )]
    param (
        [Parameter(
            ParameterSetName = "RGB",
            Mandatory,
            Position = 0
        )]
        [int]
        $red,

        [Parameter(
            ParameterSetName = "RGB",
            Mandatory,
            Position = 1
        )]
        [int]
        $green,

        [Parameter(
            ParameterSetName = "RGB",
            Mandatory,
            Position = 2
        )]
        [int]
        $blue,

        [Parameter(
            ParameterSetName = "Hex",
            Mandatory
        )]
        [string]
        $hexColor
    )

    if ($PSCmdlet.ParameterSetName -eq "Hex") {
        # Remove '#' from hexColor string
        $hexColor = $hexColor -replace '#', ''

        # Convert hex to RGB values
        $red = [Convert]::ToInt32($hexColor.Substring(0, 2), 16)
        $green = [Convert]::ToInt32($hexColor.Substring(2, 2), 16)
        $blue = [Convert]::ToInt32($hexColor.Substring(4, 2), 16)
    }

    # Calculate the color index for the RGB values
    return $blue * 65536 + $green * 256 + $red
}

