Function New-StaffCalendar {
    [CmdletBinding()]
    param (
        # Calendar year to create
        [Parameter(
            Mandatory
        )]
        [int]
        $year,

        # List of users to be added
        [Parameter(
            Mandatory
        )]
        [string[]]
        $users,

        # Excel file name
        [Parameter()]
        [string]
        $excelFileName = "TPM IT - $year Team Schedule",

        # Worksheet title row
        [Parameter()]
        [string]
        $worksheetTitleRow,

        # Worksheet Zoom Level
        [Parameter(
            HelpMessage = "Set the zoom level you would like for each sheet."
        )]
        [int]
        $worksheetZoomLevel = 100
    )

    # Creates new Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()

    # Get the list of month names and abbreviated month names.
    $monthNameList = (Get-Culture).DateTimeFormat.MonthNames
    $abbreviatedMonthNameList = (Get-Culture).DateTimeFormat.AbbreviatedMonthNames

    # Initialize an array to hold the custom objects
    $months = @()

    # Loop through the month names and create a custom object for each month
    for ($i = 0; $i -lt $monthNameList.Length; $i++) {
        if ($monthNameList[$i] -ne "") {
            $month = [PSCustomObject]@{
                MonthNumber     = $i + 1  # Month number (1-based index)
                MonthName       = $monthNameList[$i]
                AbbreviatedName = $AbbreviatedMonthNameList[$i]
            }
            $months += $month
        }
    }

    foreach ($month in $months) {
        # Add New Sheet
        $worksheet = $workbook.Worksheets.Add(
            [System.Reflection.Missing]::Value, $workbook.Worksheets.Item($workbook.Worksheets.Count)
        )

        # Rename new sheet to month abbreviated name
        $worksheet.Name = $month.AbbreviatedName

        # Set zoom level
        if ($worksheetZoomLevel -ne 100) {
            $excel.ActiveWindow.Zoom = $worksheetZoomLevel
        }

        # Set column widths
        $worksheet.Columns.Item("A").ColumnWidth = 12
        $worksheet.Columns.Item("B").ColumnWidth = 11.5
        $worksheet.Columns.Item("C").ColumnWidth = 11.5
        $worksheet.Columns.Item("D").ColumnWidth = 11.5
        $worksheet.Columns.Item("E").ColumnWidth = 11.5
        $worksheet.Columns.Item("F").ColumnWidth = 11.5
        $worksheet.Columns.Item("G").ColumnWidth = 2

        # Calculate the first and last day of the month
        $firstDayOfMonth = Get-Date -Year $year -Month $month.MonthNumber -Day 1
        $lastDayOfMonth = $firstDayOfMonth.AddMonths(1).AddDays(-1)

        # Initialize an array to hold the workdays
        $workdays = @()

        # Loop through each day of the month
        $currentDay = $firstDayOfMonth
        while ($currentDay -le $lastDayOfMonth) {
            # Check if the current day is a weekday (Monday to Friday)
            if ($currentDay.DayOfWeek -ne 'Saturday' -and $currentDay.DayOfWeek -ne 'Sunday') {
                $workdays += $currentDay
            }
            # Move to the next day
            $currentDay = $currentDay.AddDays(1)
        }

        # Group the workdays by week
        $weeks = @()
        $currentWeek = @()
        $lastWeekNumber = $null

        foreach ($day in $workdays) {
            $weekNumber = [System.Globalization.CultureInfo]::CurrentCulture.Calendar.GetWeekOfYear(
                $day, [System.Globalization.CalendarWeekRule]::FirstDay, [System.DayOfWeek]::Monday
            )
            if ($weekNumber -ne $lastWeekNumber) {
                if ($currentWeek.Count -gt 0) {
                    $weeks += , @($currentWeek)
                    $currentWeek = @()
                }
                $lastWeekNumber = $weekNumber
            }
            $currentWeek += $day
        }

        if ($currentWeek.Count -gt 0) {
            $weeks += , @($currentWeek)
        }

        # Define title cell settings
        $worksheet.Cells.Item(1, 2).Value2 = "$worksheetTitleRow - $($month.MonthName)"
        $worksheet.Cells.Item(1, 2).Font.Size = 22
        $worksheet.Cells.Item(1, 2).Font.Bold = $true

        # Merge and center title cells (B through F)
        $range = $worksheet.Range("B1:F1")
        $range.Merge()
        $range.HorizontalAlignment = -4108  # Center horizontally (xlCenter)
        $range.VerticalAlignment = -4108    # Center vertically (xlCenter)

        # Write the weeks to the Excel worksheet starting at row 4.
        $row = 4
        foreach ($week in $weeks) {
            # Define week name and date range
            $dateCellRange = $worksheet.Range("B$($row-1):F$($row)")

            # Set background color to RGB(231,230,230) or #E7E6E6
            $dateCellRange.Interior.Color = 15132391

            # Add borders to the range with xlContinuous
            $dateCellRange.Borders.LineStyle = 1

            # Insert users rows
            $userRowCount = $row + 1
            foreach ($user in $users) {
                $worksheet.Cells.Item(($userRowCount), (1)) = $user
                # Set borders around all users data cells
                $worksheet.Range(
                    $worksheet.Cells.Item($userRowCount, 1),
                    $worksheet.Cells.Item($userRowCount, 6)
                ).Borders.LineStyle = 1
                $worksheet.Cells.Item(($userRowCount++), (1)).Font.Bold = $true
            }

            # Date data starts at cell B
            $col = 2

            # Insert Day Data
            foreach ($day in $week) {
                # Move start of month cell to the correct location.
                if (
                    # Check if it's the first week of the month.
                    [bool](
                        Compare-Object -ReferenceObject $weeks[0] -DifferenceObject $week -ExcludeDifferent -IncludeEqual
                    ) -and (
                        #Check if it's the first workday of the month.
                        $day -eq $week[0]
                    )
                ) {
                    # If both checks are true then move the starting cell over the day of the week -1.
                    $col = $col + $day.DayOfWeek.value__ - 1
                }
                # Set work day cell
                $weekDayCell = $worksheet.Cells.Item(($row - 1), $col)
                $weekDayCell.Value2 = $day.DayOfWeek.ToString()
                $weekDayCell.Font.Bold = $true
                $weekDayCell.HorizontalAlignment = -4108  # -4108 corresponds to center alignment

                # Set date cell
                $dateCell = $worksheet.Cells.Item($row, $col)
                $dateCell.Value2 = $day.ToString("yyyy-MM-dd")
                $dateCell.Font.Bold = $true
                $dateCell.HorizontalAlignment = -4108  # -4108 corresponds to center alignment

                # Add work hours for each user
                $userRowCount = $row + 1
                foreach ($user in $users) {
                    $worksheet.Cells.Item($userRowCount++, $col).Value2 = "7:30 - 4:30"
                }

                # Incase column count
                $col++
            }
            $row = $row + $users.Count + 3
        }

        # Set Font for the Sheet
        $worksheet.UsedRange.Font.Name = "Calibri"
    }

    # Delete worksheet named Sheet1
    $workbook.Worksheets.Item("Sheet1").Delete()

    # Set Jan cell to tbe the active one
    $workbook.Worksheets.Item("Jan").Activate()

    # Define Excel file properties
    $excelFile = "$PSScriptRoot\$excelFileName.xlsx"

    # Remove existing file
    Remove-Item -Path $excelFile -ErrorAction SilentlyContinue

    # Save the Excel file
    $workbook.SaveAs($excelFile)

    # Excel clean up
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Output "Excel file created: $excelFile"
}