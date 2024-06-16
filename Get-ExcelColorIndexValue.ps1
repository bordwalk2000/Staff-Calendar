function Get-ExcelColorIndexValue {
    [CmdletBinding(DefaultParameterSetName = "RGB")]
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

