Add-Type -path "C:\Users\hectorc\OneDrive - Keystone Concrete\Documents\WindowsPowerShell\Modules\ImportExcel\7.8.9\EPPlus.dll"


function Format-Cell {
    param($Cell, $Value, $Format = "", $Bold = $false, $BgColor = $null, $FontColor = $null)
    if ($null -ne $Value -and $Value -ne "") { 
        # Only check for NaN/Infinity on numeric types
        $isValidValue = $true
        if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal]) {
            $isValidValue = -not [double]::IsNaN($Value) -and -not [double]::IsInfinity($Value)
        }
        if ($isValidValue) {
            $Cell.Value = $Value 
        }
    }
    if ($Format) { $Cell.Style.Numberformat.Format = $Format }
    if ($Bold) { $Cell.Style.Font.Bold = $true }
    if ($BgColor) {
        $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $Cell.Style.Fill.BackgroundColor.SetColor($BgColor)
    }
    if ($FontColor) { $Cell.Style.Font.Color.SetColor($FontColor) }
}



function Set-CurrencyCell {
    param($Cell, [double]$Value, $Bold = $false)
    Format-Cell $Cell $Value '$#,##0.00' $Bold
}

function Style-TotalRow {
    param($Range)
    $Range.Style.Fill.PatternType = 'Solid'
    $Range.Style.Fill.BackgroundColor.SetColor([Drawing.Color]::FromArgb(220,220,220))
    $Range.Style.Font.Bold = $true
}

# Define cost-family metadata once so headers and rows stay in sync
$costFamilies = @(
    @{ Key = 'HoursCost';       Title = 'Labor %'        },
    @{ Key = 'MaterialCost';    Title = 'Material %'     },
    @{ Key = 'EquipmentCost';   Title = 'Equipment %'    },
    @{ Key = 'SubcontractCost'; Title = 'Subcontract %'  },
    @{ Key = 'OtherCost';       Title = 'Other %'        }
)

function Add-BudgetActualRow {
    param(
        $Sheet,         # Excel worksheet
        [int]$Row,      # Target row
        $JobData,       # Job object from JSON
        [int]$StartCol  # Starting column for first cost type
    )

    $col = $StartCol
    foreach ($fam in $costFamilies) {
        $obj = $JobData.($fam.Key)
        if (-not $obj) { $obj = @{ PercentOfBudget = 0; Actual = 0; Budget = 0; Remaining = 0 } }

        # Percentage first (visible when collapsed) - use proper conditional formatting
        $percentCell = $Sheet.Cells.Item($Row, $col)
        $actual = [double]($obj.Actual ?? 0)
        $budget = [double]($obj.Budget ?? 0)
        $remaining = [double]($obj.Remaining ?? 0)
        
        Set-PercentCellWithConditionalFormatting $percentCell $actual $budget $remaining
        
        # Detail columns (collapsible) - ensure zeros are displayed
        $Sheet.Cells.Item($Row, $col + 1).Value = [double]($obj.Actual ?? 0)
        $Sheet.Cells.Item($Row, $col + 1).Style.Numberformat.Format = '$#,##0.00'
        
        $Sheet.Cells.Item($Row, $col + 2).Value = [double]($obj.Budget ?? 0)
        $Sheet.Cells.Item($Row, $col + 2).Style.Numberformat.Format = '$#,##0.00'
        
        $Sheet.Cells.Item($Row, $col + 3).Value = [double]($obj.Remaining ?? 0)
        $Sheet.Cells.Item($Row, $col + 3).Style.Numberformat.Format = '$#,##0.00'
        
        $col += 5  # Move to next cost type (% + 3 detail columns + 1 gap)
    }
}

function New-PMSummarySheet {
    param($Excel, $JobSummaryData, $PMName, $PMJobs, $Group)
    
    $summarySheet = $Excel.Workbook.Worksheets.Add("Summary")
    $currentRow = 1
    
    # Title
    $formattedDate = [DateTime]::ParseExact($timestamp, "yyyyMMdd", $null).ToString("MM/dd/yyyy")
    Format-Cell $summarySheet.Cells["A1"] "$PMName - $Group Projects - Manhour Summary - $formattedDate" "" $true
    $summarySheet.Cells["A1"].Style.Font.Size = 14
    # Title merge handled later dynamically based on header count
    
    # Build headers dynamically from $costFamilies so order is single-source
    # Each cost family now has: % (visible), Actual, Budget, Remaining (collapsible), Gap
    $costHeaders = @()
    foreach ($fam in $costFamilies) {
        $costHeaders += @($fam.Title, ($fam.Title -replace ' %', ' Cost'), ($fam.Title -replace ' %', ' Budget'), ($fam.Title -replace ' %', ' Remaining'), '')
    }
    # Project section: Project % (visible), Contract Amount, Project Budget, Project Cost, Est. Profit (collapsible)
    $headers = @('Job Number','Job Name') + $costHeaders + @('Project %', 'Contract Amount', 'Project Budget', 'Project Cost', 'Est. Profit')
    $currentRow = 3
    # merge title row across all header columns
    $summarySheet.Cells[1,1,1,$headers.Count].Merge = $true

    Write-Host "[DEBUG] Header count: $($headers.Count)"  # quick test output

    for ($i = 0; $i -lt $headers.Count; $i++) {
        Format-Cell $summarySheet.Cells.Item($currentRow, $i+1) $headers[$i] "" $true ([System.Drawing.Color]::Black) ([System.Drawing.Color]::White)
    }
    $currentRow++
    
    # Create collapsible column groups for cost sections - collapsed by default
    try {
        # Define column groups for each cost type: [start_col, end_col] for detail columns
        $summaryColumnGroups = @(
            @(4, 6),   # Labor: Labor Cost, Labor Budget, Labor Remaining (gap col 7)
            @(9, 11),  # Material: Material Cost, Material Budget, Material Remaining (gap col 12)
            @(14, 16), # Equipment: Equipment Cost, Equipment Budget, Equipment Remaining (gap col 17)
            @(19, 21), # Subcontract: Subcontract Cost, Subcontract Budget, Subcontract Remaining (gap col 22)
            @(24, 26)  # Other: Other Cost, Other Budget, Other Remaining (gap col 27)
        )
        
        foreach ($group in $summaryColumnGroups) {
            for ($col = $group[0]; $col -le $group[1]; $col++) {
                $summarySheet.Column($col).OutlineLevel = 1
                $summarySheet.Column($col).Collapsed = $true
            }
            # Also collapse the gap column
            $summarySheet.Column($group[1] + 1).OutlineLevel = 1
            $summarySheet.Column($group[1] + 1).Collapsed = $true
        }
        
        # Project section: Project % (visible), Contract Amount, Project Budget, Project Cost, Est. Profit (collapsible but expanded by default)
        # Calculate project section start column (after cost families)
        $projectStartCol = 3 + ($costFamilies.Count * 5)  # 2 base cols + (5 cols per cost family)
        for ($col = $projectStartCol + 1; $col -le $projectStartCol + 4; $col++) {  # Skip Project % column
            $summarySheet.Column($col).OutlineLevel = 1
            # Don't set Collapsed = $true for project section - leave expanded by default
        }
    } catch {
        Write-Warning "Unable to create column groups on PM summary sheet: $($_.Exception.Message)"
    }
    
    # PM totals tracking
    $pmTotals = @{ ContractAmount = 0; ProjectBudget = 0; ProjectCost = 0; EstProfit = 0 }
    
    # Initialize cost family totals
    $costTotals = @{}
    foreach ($fam in $costFamilies) {
        $costTotals[$fam.Key] = @{ Actual = 0; Budget = 0; Remaining = 0 }
    }
    
    # Process PM's jobs
    foreach ($jobId in $PMJobs) {
        # Retrieve job object via dynamic property access (JSON keys become PSCustomObject properties)
        $job = $JobSummaryData."$jobId"
        if (-not $job) { continue }
        
        # Job row with hyperlink to individual sheet
        Format-Cell $summarySheet.Cells.Item($currentRow, 1) $jobId
        $worksheetName = if ($jobId.Length -gt 31) { $jobId.Substring(0, 31) } else { $jobId }
        
        # Create hyperlink to the job sheet within the same workbook
        $summarySheet.Cells.Item($currentRow, 1).Hyperlink = New-Object OfficeOpenXml.ExcelHyperLink("'$worksheetName'!A1", $jobId)
        $summarySheet.Cells.Item($currentRow, 1).Style.Font.UnderLine = $true
        $summarySheet.Cells.Item($currentRow, 1).Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        
        Format-Cell $summarySheet.Cells.Item($currentRow, 2) $job.JobName
        
        # Add job data to summary
        
        Add-BudgetActualRow $summarySheet $currentRow $job 3
        
        # Project budget and cost columns
        $projectBudget  = if ($job.ProjectCost.Budget -ne $null) { [double]$job.ProjectCost.Budget } else { 0 }
        $projectActual  = if ($job.ProjectCost.Actual -ne $null) { [double]$job.ProjectCost.Actual } else { 0 }

        $contractAmt    = if ($job.ContractAmount -ne $null) { [double]$job.ContractAmount } else { 0 }
        $estProfitVal   = if ($job.EstProfit -ne $null)   { [double]$job.EstProfit }   else { 0 }

        $projPercentCol = $headers.IndexOf("Project %")       + 1
        $contractCol    = $headers.IndexOf("Contract Amount") + 1
        $projBudgetCol  = $headers.IndexOf("Project Budget")  + 1
        $projCostCol    = $headers.IndexOf("Project Cost")    + 1
        $profitCol      = $headers.IndexOf("Est. Profit")      + 1

        # Project percent column (now first in project section) - use proper conditional formatting
        $projCell = $summarySheet.Cells.Item($currentRow, $projPercentCol)
        $projActual = if ($job.ProjectCost.Actual -ne $null) { [double]$job.ProjectCost.Actual } else { 0 }
        $projBudget = if ($job.ProjectCost.Budget -ne $null) { [double]$job.ProjectCost.Budget } else { 0 }
        $projRemaining = if ($job.ProjectCost.Remaining -ne $null) { [double]$job.ProjectCost.Remaining } else { 0 }
        
        Set-PercentCellWithConditionalFormatting $projCell $projActual $projBudget $projRemaining

        # Project detail columns (collapsible) - ensure zeros are displayed
        $summarySheet.Cells.Item($currentRow,$contractCol).Value = [double]$contractAmt
        $summarySheet.Cells.Item($currentRow,$contractCol).Style.Numberformat.Format = '$#,##0.00'
        
        $summarySheet.Cells.Item($currentRow,$projBudgetCol).Value = [double]$projectBudget
        $summarySheet.Cells.Item($currentRow,$projBudgetCol).Style.Numberformat.Format = '$#,##0.00'
        
        $summarySheet.Cells.Item($currentRow,$projCostCol).Value = [double]$projectActual
        $summarySheet.Cells.Item($currentRow,$projCostCol).Style.Numberformat.Format = '$#,##0.00'
        
        $summarySheet.Cells.Item($currentRow,$profitCol).Value = [double]$estProfitVal
        $summarySheet.Cells.Item($currentRow,$profitCol).Style.Numberformat.Format = '$#,##0.00'
        
        # Totals not used for percentage-only view; can compute later if needed
        
        # accumulate totals
        $pmTotals.ContractAmount += $contractAmt
        $pmTotals.ProjectBudget  += $projectBudget
        $pmTotals.ProjectCost    += $projectActual
        $pmTotals.EstProfit      += $estProfitVal
        
        # Accumulate cost family totals
        foreach ($fam in $costFamilies) {
            $costObj = $job.($fam.Key)
            if ($costObj) {
                $costTotals[$fam.Key].Actual += [double]($costObj.Actual ?? 0)
                $costTotals[$fam.Key].Budget += [double]($costObj.Budget ?? 0)
                $costTotals[$fam.Key].Remaining += [double]($costObj.Remaining ?? 0)
            }
        }
        
        $currentRow++
    }
    # Add TOTAL row (immediately after data rows)
    $totalRow = $currentRow

    # Label
    Format-Cell $summarySheet.Cells.Item($totalRow,2) "TOTAL" "" $true ([System.Drawing.Color]::Black) ([System.Drawing.Color]::Black)
    
    # Add cost family totals to the totals row
    $col = 3
    foreach ($fam in $costFamilies) {
        $costTotal = $costTotals[$fam.Key]
        
        # Calculate percentage
        $percent = if ($costTotal.Budget -gt 0) { 
            [Math]::Round(($costTotal.Actual / $costTotal.Budget) * 100, 0) 
        } else { 0 }
        
        # Percentage column with conditional formatting
        $percentCell = $summarySheet.Cells.Item($totalRow, $col)
        Set-PercentCellWithConditionalFormatting $percentCell $costTotal.Actual $costTotal.Budget $costTotal.Remaining $true
        
        # Detail columns - ensure zeros are displayed
        $summarySheet.Cells.Item($totalRow, $col + 1).Value = [double]$costTotal.Actual
        $summarySheet.Cells.Item($totalRow, $col + 1).Style.Numberformat.Format = "$#,##0.00"
        $summarySheet.Cells.Item($totalRow, $col + 1).Style.Font.Bold = $true
        
        $summarySheet.Cells.Item($totalRow, $col + 2).Value = [double]$costTotal.Budget
        $summarySheet.Cells.Item($totalRow, $col + 2).Style.Numberformat.Format = "$#,##0.00"
        $summarySheet.Cells.Item($totalRow, $col + 2).Style.Font.Bold = $true
        
        $summarySheet.Cells.Item($totalRow, $col + 3).Value = [double]$costTotal.Remaining
        $summarySheet.Cells.Item($totalRow, $col + 3).Style.Numberformat.Format = "$#,##0.00"
        $summarySheet.Cells.Item($totalRow, $col + 3).Style.Font.Bold = $true
        
        $col += 5  # Move to next cost type
    }

    # Total currency values and project percent
    $projPercentCol = $headers.IndexOf("Project %")       + 1
    $contractCol    = $headers.IndexOf("Contract Amount") + 1
    $projBudgetCol  = $headers.IndexOf("Project Budget")  + 1
    $projCostCol    = $headers.IndexOf("Project Cost")    + 1
    $profitCol      = $headers.IndexOf("Est. Profit")      + 1

    # Project % total (now first in project section) - use proper conditional formatting
    $pctCell = $summarySheet.Cells.Item($totalRow, $projPercentCol)
    $projRemainingTotal = $pmTotals.ProjectBudget - $pmTotals.ProjectCost
    
    Set-PercentCellWithConditionalFormatting $pctCell $pmTotals.ProjectCost $pmTotals.ProjectBudget $projRemainingTotal $true

    # Project detail columns (collapsible) - ensure zeros are displayed
    $summarySheet.Cells.Item($totalRow,$contractCol).Value = [double]$pmTotals.ContractAmount
    $summarySheet.Cells.Item($totalRow,$contractCol).Style.Numberformat.Format = "$#,##0.00"
    $summarySheet.Cells.Item($totalRow,$contractCol).Style.Font.Bold = $true
    
    $summarySheet.Cells.Item($totalRow,$projBudgetCol).Value = [double]$pmTotals.ProjectBudget
    $summarySheet.Cells.Item($totalRow,$projBudgetCol).Style.Numberformat.Format = "$#,##0.00"
    $summarySheet.Cells.Item($totalRow,$projBudgetCol).Style.Font.Bold = $true
    
    $summarySheet.Cells.Item($totalRow,$projCostCol).Value = [double]$pmTotals.ProjectCost
    $summarySheet.Cells.Item($totalRow,$projCostCol).Style.Numberformat.Format = "$#,##0.00"
    $summarySheet.Cells.Item($totalRow,$projCostCol).Style.Font.Bold = $true
    
    $summarySheet.Cells.Item($totalRow,$profitCol).Value = [double]$pmTotals.EstProfit
    $summarySheet.Cells.Item($totalRow,$profitCol).Style.Numberformat.Format = "$#,##0.00"
    $summarySheet.Cells.Item($totalRow,$profitCol).Style.Font.Bold = $true

    # Style entire total row - but preserve conditional formatting on percentage cells
    $totalRange = $summarySheet.Cells.Item($totalRow,1,$totalRow,$headers.Count)
    Style-TotalRow $totalRange
    
    # Reapply conditional formatting to percentage cells (cost families + project %)
    $col = 3
    foreach ($fam in $costFamilies) {
        $costTotal = $costTotals[$fam.Key]
        $percentCell = $summarySheet.Cells.Item($totalRow, $col)
        Set-PercentCellWithConditionalFormatting $percentCell $costTotal.Actual $costTotal.Budget $costTotal.Remaining $true
        $col += 5  # Move to next cost type
    }
    
    # Reapply project percentage formatting
    $pctCell = $summarySheet.Cells.Item($totalRow, $projPercentCol)
    Set-PercentCellWithConditionalFormatting $pctCell $pmTotals.ProjectCost $pmTotals.ProjectBudget $projRemainingTotal $true

    # Freeze panes to keep title and headers visible (rows 1-3)
    $summarySheet.View.FreezePanes(4, 1)
    
    # Autofit columns for better spacing
    $summarySheet.Cells.AutoFitColumns()
}

# Helper function to set number formatting based on type
function Set-NumberCell {
    param($Cell, [double]$Value, [string]$Type = "Number", $Bold = $false)
    
    $format = switch ($Type) {
        "Currency" { "$#,##0.00" }
        "Hours" { "#,##0.0" }
        "Percent" { "0\%" }
        default { "#,##0.00" }
    }
    
    # Always set the value, even if it's zero
    $Cell.Value = $Value
    $Cell.Style.Numberformat.Format = $format
    if ($Bold) { $Cell.Style.Font.Bold = $true }
}

# Helper function to set percentage cell with proper conditional formatting
function Set-PercentCellWithConditionalFormatting {
    param(
        $Cell, 
        [double]$Actual = 0, 
        [double]$Budget = 0, 
        [double]$Remaining = 0, 
        $Bold = $false
    )
    
    # Calculate percentage
    $percent = if ($Budget -gt 0) { 
        [Math]::Round(($Actual / $Budget) * 100, 0) 
    } else { 0 }
    
    # Set cell value and format
    $Cell.Value = $percent
    $Cell.Style.Numberformat.Format = "0\%"
    if ($Bold) { $Cell.Style.Font.Bold = $true }
    
    # Apply conditional formatting
    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    
    if ($percent -eq 0) {
        # Check if there's actual cost but no budget (should be bright red)
        if ($Actual -gt 0 -and $Budget -eq 0) {
            $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
            $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        } elseif ($Budget -gt 0 -and $Actual -eq 0) {
            # Budget but no cost (should be bright green)
            $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 198, 239, 206))
            $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
        } else {
            # True zero activity - no budget and no cost (should be yellow)
            $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
        }
    } elseif ($percent -gt 100) {
        # Over budget (light red)
        $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 255, 199, 206))
    } else {
        # Normal range 1-100% (light green)
        $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 198, 239, 206))
    }
}

# Helper function to create merged title/header rows
function New-MergedRow {
    param($Worksheet, [int]$Row, [string]$Text, [int]$StartCol = 1, [int]$EndCol = 12, 
          $Bold = $true, $FontSize = 10, $Alignment = "Left")
    
    $range = $Worksheet.Cells[$Row, $StartCol, $Row, $EndCol]
    $range.Merge = $true
    $range.Value = $Text
    $range.Style.Font.Bold = $Bold
    if ($FontSize -ne 10) { $range.Style.Font.Size = $FontSize }
    $range.Style.HorizontalAlignment = switch ($Alignment) {
        "Center" { [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
        "Right" { [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right }
        default { [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left }
    }
}

# Helper function to create a data row with proper formatting
function New-DataRow {
    param($Worksheet, [int]$Row, $DataObj, [string]$Label1 = "", [string]$Label2 = "", 
          [int]$StartCol = 4, $Bold = $false, $BgColor = $null)
    
    # Labels
    if ($Label1) { Format-Cell $Worksheet.Cells.Item($Row, 1) $Label1 "" $Bold $BgColor }
    if ($Label2) { Format-Cell $Worksheet.Cells.Item($Row, 2) $Label2 "" $Bold $BgColor }
    
    # Apply background to all cells if specified
    if ($BgColor) {
        1..12 | ForEach-Object {
            $Worksheet.Cells.Item($Row, $_).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $Worksheet.Cells.Item($Row, $_).Style.Fill.BackgroundColor.SetColor($BgColor)
        }
    }
    
    # Data columns (4-7: Hours section, 9-12: Cost section)
    if ($DataObj.Hours) {
        # Hours data
        Set-NumberCell $Worksheet.Cells.Item($Row, $StartCol) $DataObj.Hours.Actual "Hours" $Bold
        Set-NumberCell $Worksheet.Cells.Item($Row, $StartCol + 1) $DataObj.Hours.Budget "Hours" $Bold
        Set-NumberCell $Worksheet.Cells.Item($Row, $StartCol + 2) $DataObj.Hours.Remaining "Hours" $Bold
        
        $hoursPercentCell = $Worksheet.Cells.Item($Row, $StartCol + 3)
        $hoursActual = [double]($DataObj.Hours.Actual ?? 0)
        $hoursBudget = [double]($DataObj.Hours.Budget ?? 0)
        $hoursRemaining = [double]($DataObj.Hours.Remaining ?? 0)
        
        Set-PercentCellWithConditionalFormatting $hoursPercentCell $hoursActual $hoursBudget $hoursRemaining $Bold
    }
    
    if ($DataObj.HoursCost) {
        # Cost data
        Set-NumberCell $Worksheet.Cells.Item($Row, $StartCol + 5) $DataObj.HoursCost.Actual "Currency" $Bold
        Set-NumberCell $Worksheet.Cells.Item($Row, $StartCol + 6) $DataObj.HoursCost.Budget "Currency" $Bold
        Set-NumberCell $Worksheet.Cells.Item($Row, $StartCol + 7) $DataObj.HoursCost.Remaining "Currency" $Bold
        
        $costPercentCell = $Worksheet.Cells.Item($Row, $StartCol + 8)
        $costActual = [double]($DataObj.HoursCost.Actual ?? 0)
        $costBudget = [double]($DataObj.HoursCost.Budget ?? 0)
        $costRemaining = [double]($DataObj.HoursCost.Remaining ?? 0)
        
        Set-PercentCellWithConditionalFormatting $costPercentCell $costActual $costBudget $costRemaining $Bold
    }
}

# Helper function to create a data row with all cost types
function New-AllCostDataRow {
    param($Worksheet, [int]$Row, $DataObj, [string]$Label1 = "", [string]$Label2 = "", 
          [int]$StartCol = 4, $Bold = $false, $BgColor = $null, [switch]$CenterAlign)
    
    # Labels
    if ($Label1) { 
        Format-Cell $Worksheet.Cells.Item($Row, 1) $Label1 "" $Bold $BgColor 
        if ($CenterAlign) { $Worksheet.Cells.Item($Row, 1).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
    }
    if ($Label2) { 
        Format-Cell $Worksheet.Cells.Item($Row, 2) $Label2 "" $Bold $BgColor 
        if ($CenterAlign) { $Worksheet.Cells.Item($Row, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
    }
    
    # Apply background to all cells if specified
    if ($BgColor) {
        1..32 | ForEach-Object {
            $Worksheet.Cells.Item($Row, $_).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $Worksheet.Cells.Item($Row, $_).Style.Fill.BackgroundColor.SetColor($BgColor)
        }
    }
    
    # Define all cost types with their column positions
    $costTypes = @(
        @{ Name = "Hours"; StartCol = 4; Type = "Hours" },
        @{ Name = "HoursCost"; StartCol = 9; Type = "Currency" },
        @{ Name = "MaterialCost"; StartCol = 14; Type = "Currency" },
        @{ Name = "SubcontractCost"; StartCol = 19; Type = "Currency" },
        @{ Name = "EquipmentCost"; StartCol = 24; Type = "Currency" },
        @{ Name = "OtherCost"; StartCol = 29; Type = "Currency" }
    )
    
    foreach ($costType in $costTypes) {
        $costObj = $DataObj.($costType.Name)
        $col = $costType.StartCol
        
        # Handle missing cost objects by creating default zero values
        if (-not $costObj) {
            $costObj = @{ Actual = 0; Budget = 0; Remaining = 0; PercentOfBudget = 0 }
        }
        
        # Percentage column first (most important for collapsed view)
        $percentCell = $Worksheet.Cells.Item($Row, $col)
        $actual = [double]($costObj.Actual ?? 0)
        $budget = [double]($costObj.Budget ?? 0)
        $remaining = [double]($costObj.Remaining ?? 0)
        
        Set-PercentCellWithConditionalFormatting $percentCell $actual $budget $remaining $Bold
        
        # Actual, Budget, Remaining columns - always show values, even zeros
        Set-NumberCell $Worksheet.Cells.Item($Row, $col + 1) ([double]($costObj.Actual ?? 0)) $costType.Type $Bold
        Set-NumberCell $Worksheet.Cells.Item($Row, $col + 2) ([double]($costObj.Budget ?? 0)) $costType.Type $Bold
        Set-NumberCell $Worksheet.Cells.Item($Row, $col + 3) ([double]($costObj.Remaining ?? 0)) $costType.Type $Bold
        
        # Apply center alignment if requested
        if ($CenterAlign) {
            $Worksheet.Cells.Item($Row, $col).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            $Worksheet.Cells.Item($Row, $col + 1).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            $Worksheet.Cells.Item($Row, $col + 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            $Worksheet.Cells.Item($Row, $col + 3).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }
    }
}

function New-ReportSheet {
    param($JobId, $JobData, $Worksheet, [switch]$ReturnData, [switch]$IncludeSummaryLink)
    
    $grey = [System.Drawing.Color]::FromArgb(255, 217, 217, 217)
    $black = [System.Drawing.Color]::Black
    $white = [System.Drawing.Color]::White
    
    # Calculate initial row based on whether we have the summary link
    $initialRow = if ($IncludeSummaryLink) { 5 } else { 4 }
    $currentRow = $initialRow
    
    # Return to Summary link
    if ($IncludeSummaryLink) {
        $Worksheet.Cells["A1"].Value = "Return to Summary"
        $Worksheet.Cells["A1"].Hyperlink = New-Object OfficeOpenXml.ExcelHyperLink("'Summary'!A1", "Return to Summary")
        $Worksheet.Cells["A1"].Style.Font.UnderLine = $true
        $Worksheet.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        New-MergedRow $Worksheet 1 "Return to Summary" 1 3 $true 10 "Left"
    }
    
    # Calculate summary values first
    $contractAmount = if ($JobData.ContractAmount -ne $null) { $JobData.ContractAmount } else { 0 }
    $projectCostToUse = if ($JobData.ProjectCost.Actual -gt $JobData.ProjectCost.Budget) { $JobData.ProjectCost.Actual } else { $JobData.ProjectCost.Budget }
    $estProfit = if ($contractAmount -gt 0 -and $projectCostToUse -ne $null) { 
        [Math]::Round([double]$contractAmount - [double]$projectCostToUse, 2) 
    } else { 0 }
    
    $hourlyActualRate = if ($JobData.Hours.Actual -gt 0) { [Math]::Round($JobData.HoursCost.Actual / $JobData.Hours.Actual, 2) } else { 0 }
    $hourlyBudgetRate = if ($JobData.Hours.Budget -gt 0) { [Math]::Round($JobData.HoursCost.Budget / $JobData.Hours.Budget, 2) } else { 0 }
    $hourlyPercent = if ($JobData.Hours.Budget -gt 0 -and $JobData.HoursCost.Budget -gt 0 -and $JobData.Hours.Actual -gt 0) {
        $actualRate = $JobData.HoursCost.Actual / $JobData.Hours.Actual
        $budgetRate = $JobData.HoursCost.Budget / $JobData.Hours.Budget
        if ($budgetRate -gt 0) { [Math]::Round(($actualRate / $budgetRate) * 100, 0) } else { 0 }
    } else { 0 }
    
    # Title and PM rows - added after auto-fit, adjusted for frozen columns
    $formattedDate = [DateTime]::ParseExact($timestamp, "yyyyMMdd", $null).ToString("MM/dd/yyyy")
    $titleRow = if ($IncludeSummaryLink) { 2 } else { 1 }
    
    # Title split into two lines - ID/Date and Project Name separately
    New-MergedRow $Worksheet $titleRow "$JobId - $formattedDate" 1 3 $true 14
    New-MergedRow $Worksheet ($titleRow + 1) "$($JobData.JobName)" 1 3 $true 12
    New-MergedRow $Worksheet ($titleRow + 2) "Project Manager: $($JobData.PM)" 1 3 $true
    $currentRow++  # Add space row after PM name
    
    # Summary section in left columns (A-B) - rearranged order
    Format-Cell $Worksheet.Cells.Item($currentRow, 1) "Contract Amount" "" $true
    Set-NumberCell $Worksheet.Cells.Item($currentRow, 2) $contractAmount "Currency" $true
    $currentRow++
    
    Format-Cell $Worksheet.Cells.Item($currentRow, 1) "Project Budget" "" $true
    Set-NumberCell $Worksheet.Cells.Item($currentRow, 2) $JobData.ProjectCost.Budget "Currency" $true
    $currentRow++
    
    Format-Cell $Worksheet.Cells.Item($currentRow, 1) "Project Cost" "" $true
    Set-NumberCell $Worksheet.Cells.Item($currentRow, 2) $JobData.ProjectCost.Actual "Currency" $true
    $currentRow++
    
    Format-Cell $Worksheet.Cells.Item($currentRow, 1) "Est. Profit" "" $true
    Set-NumberCell $Worksheet.Cells.Item($currentRow, 2) $estProfit "Currency" $true
    $currentRow++
    
    Format-Cell $Worksheet.Cells.Item($currentRow, 1) "Project %" "" $true
    $projPercentCell = $Worksheet.Cells.Item($currentRow, 2)
    $projActual = [double]($JobData.ProjectCost.Actual ?? 0)
    $projBudget = [double]($JobData.ProjectCost.Budget ?? 0)
    $projRemaining = [double]($JobData.ProjectCost.Remaining ?? 0)
    
    Set-PercentCellWithConditionalFormatting $projPercentCell $projActual $projBudget $projRemaining $true
    $currentRow += 2
    
    # Add Grand Total row for all cost types in the frozen header section
    New-AllCostDataRow $Worksheet $currentRow $JobData "Grand Total" "" 4 $true $grey -CenterAlign
    $currentRow++  # Remove extra space after grand total (was += 2)
    
    # Main table headers - rearranged with % first in each group for better collapsing
    $mainHeaders = @(
        "", "Row Labels", "",
        "Hours %", "Hours", "Hours Budget", "Hours Remaining", "",
        "Labor %", "Labor Cost", "Labor Budget", "Labor Remaining", "",
        "Material %", "Material Cost", "Material Budget", "Material Remaining", "",
        "Subcontract %", "Subcontract Cost", "Subcontract Budget", "Subcontract Remaining", "",
        "Equipment %", "Equipment Cost", "Equipment Budget", "Equipment Remaining", "",
        "Other %", "Other Cost", "Other Budget", "Other Remaining"
    )
    
    for ($i = 0; $i -lt $mainHeaders.Count; $i++) {
        Format-Cell $Worksheet.Cells.Item($currentRow, $i+1) $mainHeaders[$i] "" $true $black $white
    }
    $currentRow++
    
    # Create collapsible column groups for each cost section - collapsed by default
    try {
        # Define column groups: [start_col, end_col] for each cost section's detail columns
        $columnGroups = @(
            @(5, 8),   # Hours: Hours, Hours Budget, Hours Remaining, Gap
            @(10, 13), # Labor: Labor Cost, Labor Budget, Labor Remaining, Gap  
            @(15, 18), # Material: Material Cost, Material Budget, Material Remaining, Gap
            @(20, 23), # Subcontract: Subcontract Cost, Subcontract Budget, Subcontract Remaining, Gap
            @(25, 28), # Equipment: Equipment Cost, Equipment Budget, Equipment Remaining, Gap
            @(30, 32)  # Other: Other Cost, Other Budget, Other Remaining (no gap)
        )
        
        foreach ($group in $columnGroups) {
            for ($col = $group[0]; $col -le $group[1]; $col++) {
                $Worksheet.Column($col).OutlineLevel = 1
                $Worksheet.Column($col).Collapsed = $true
            }
        }
    } catch {
        Write-Warning "Unable to create column groups on worksheet for job ${JobId}: $($_.Exception.Message)"
    }
    
    # Freeze panes to include summary section - freeze first 3 columns (A, B, C)
    $freezeRow = $currentRow  # This will freeze everything above the detail data (summary section + headers)
    try {
        $Worksheet.View.FreezePanes($freezeRow, 4)  # Freeze at row $freezeRow, column 4 (D) - this freezes columns A, B, C and all summary rows
    } catch {
        Write-Warning "Unable to freeze panes on worksheet for job ${JobId}: $($_.Exception.Message)"
    }
    
    # Phases and categories - filter out phases 000 and 999
    if ($JobData.Phases) {
        foreach ($phaseKey in $JobData.Phases.PSObject.Properties.Name) {
            # Skip phases 000 and 999
            if ($phaseKey -eq "000" -or $phaseKey -eq "999" -or $phaseKey -eq "001" -or $phaseKey -eq "777") {
                continue
            }
            
            $phase = $JobData.Phases.$phaseKey
            
            # Phase row with grey background - use new helper for all cost types
            New-AllCostDataRow $Worksheet $currentRow $phase $phaseKey $phase.PhaseName 4 $true $grey
            $currentRow++
            $phaseStartRow = $currentRow
            
            # Categories
            if ($phase.Categories) {
                foreach ($catKey in $phase.Categories.PSObject.Properties.Name) {
                    $category = $phase.Categories.$catKey
                    New-AllCostDataRow $Worksheet $currentRow $category "   $catKey" "   $($category.CategoryName)"
                    $currentRow++
                }
                
                # Create collapsible group for this phase
                $phaseEndRow = $currentRow
                if ($phaseEndRow -ge $phaseStartRow) {
                    for ($groupRow = $phaseStartRow; $groupRow -le $phaseEndRow; $groupRow++) {
                        $Worksheet.Row($groupRow).OutlineLevel = 1
                        $Worksheet.Row($groupRow).Collapsed = $true
                    }
                }
            }
            $currentRow++
        }
    }
    
    # Auto-fit columns
    1..32 | ForEach-Object { $Worksheet.Column($_).AutoFit() }
    
    if ($ReturnData) {
        return @{
            JobId = $JobId; JobName = $JobData.JobName; PM = $JobData.PM
            Hours = $JobData.Hours; HoursCost = $JobData.HoursCost; ProjectCost = $JobData.ProjectCost
            ContractAmount = $contractAmount; EstProfit = $estProfit
        }
    }
}

# ------------------------------------------------------------
# Quick test helper
# ------------------------------------------------------------
function Test-PMSummary {
    <#
        .SYNOPSIS
            Generates a throw-away Excel file for a single PM using the latest JSON report so you can quickly confirm the summary sheet columns populate correctly.

        .PARAMETER PMName
            The project manager name to filter jobs by.

        .PARAMETER JsonPath
            Optional path to a specific JSON file. If omitted, the function looks under the reports folder for the most recent file.

        .EXAMPLE
            Test-PMSummary -PMName "John Doe"
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$PMName,
        [string]$JsonPath
    )

    # Locate JSON if not provided
    if (-not $JsonPath) {
        $latest = Get-ChildItem -Path (Join-Path $PSScriptRoot ".." "reports") -Recurse -Filter "*_JobData_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $latest) { Write-Error "Could not find any JobData JSON files under reports/"; return }
        $JsonPath = $latest.FullName
    }

    if (-not (Test-Path $JsonPath)) { Write-Error "JSON path '$JsonPath' not found."; return }
    Write-Host "Using JSON: $JsonPath"

    $jobData = Get-Content $JsonPath -Raw | ConvertFrom-Json
    if (-not $jobData) { Write-Error "Failed to parse JSON."; return }

    # Collect jobs for the requested PM
    $pmJobs = @()
    foreach ($prop in $jobData.PSObject.Properties) {
        $jobNum = $prop.Name
        $job    = $prop.Value
        if ($null -ne $job.PM -and ($job.PM -eq $PMName)) {
            $pmJobs += $jobNum
        }
    }

    if ($pmJobs.Count -eq 0) { Write-Host "No jobs found for PM '$PMName'."; return }

    # Prepare Excel package (requires ImportExcel/EPPlus already loaded by script header)
    $package = New-Object OfficeOpenXml.ExcelPackage

    # Provide timestamp variable expected by New-PMSummarySheet
    $global:timestamp = (Get-Date).ToString("yyyyMMdd")

    # Assume group is derived from JSON file path (folder name two levels up)
    $group = Split-Path (Split-Path $JsonPath -Parent) -Leaf

    New-PMSummarySheet -Excel $package -JobSummaryData $jobData -PMName $PMName -PMJobs $pmJobs -Group $group

    # Generate individual job sheets referenced by hyperlinks
    foreach ($jobId in $pmJobs) {
        $wsName = if ($jobId.Length -gt 31) { $jobId.Substring(0,31) } else { $jobId }
        # Skip if sheet already exists (avoid duplicates)
        if (-not $package.Workbook.Worksheets[$wsName]) {
            $ws = $package.Workbook.Worksheets.Add($wsName)
            $jobObj = $jobData."$jobId"
            if ($jobObj) { New-ReportSheet -JobId $jobId -JobData $jobObj -Worksheet $ws -IncludeSummaryLink }
        }
    }

    $pmNameSafe = $PMName.Replace(' ','_').Replace('.','').Replace(',','')
    $outFile = Join-Path (Split-Path $JsonPath -Parent) "CostSummary_${pmNameSafe}_${timestamp}.xlsx"
    $package.SaveAs([System.IO.FileInfo]$outFile)
    Write-Host "Test summary written to $outFile"
}


Test-PMSummary -PMName "Justin Garza" -JsonPath "C:\WorkV4\ProjectReports\reports\Houston Placement\20250704\Houston Placement_JobData_20250704.json"