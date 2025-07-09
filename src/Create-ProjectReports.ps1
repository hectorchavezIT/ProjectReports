Add-Type -path "C:\Users\hectorc\OneDrive - Keystone Concrete\Documents\WindowsPowerShell\Modules\ImportExcel\7.8.9\EPPlus.dll"

# Consolidated cell formatting function
function Set-CellData {
    param(
        $Cell, 
        $Value = "", 
        [string]$Format = "", 
        [bool]$Bold = $false, 
        $BgColor = $null, 
        $FontColor = $null,
        [string]$Type = "Text"
    )
    
    # Set value with validation for numeric types - always show zeros for numeric types
    if ($Value -ne $null) {
        $isValidValue = $true
        if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal]) {
            $isValidValue = -not [double]::IsNaN($Value) -and -not [double]::IsInfinity($Value)
        }
        if ($isValidValue) { 
            $Cell.Value = $Value 
        }
    } elseif ($Type -in @("Currency", "Hours", "Percent")) {
        # For numeric types, always show zero if no value provided
        $Cell.Value = 0
    }
    
    # Apply format based on type if no explicit format given
    if (-not $Format) {
        $Format = switch ($Type) {
            "Currency" { '$#,##0.00' }
            "Hours" { '#,##0.0' }
            "Percent" { '0\%' }
            default { '' }
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

# Simplified percentage cell with conditional formatting
function Set-PercentCell {
    param($Cell, [double]$Actual = 0, [double]$Budget = 0, [double]$Remaining = 0, [bool]$Bold = $false)
    
    $percent = if ($Budget -gt 0) { [Math]::Round(($Actual / $Budget) * 100, 0) } else { 0 }
    Set-CellData $Cell $percent -Type "Percent" -Bold $Bold
    
    # Conditional formatting
    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $bgColor = if ($percent -eq 0) {
        if ($Actual -gt 0 -and $Budget -eq 0) { [System.Drawing.Color]::Red }
        elseif ($Budget -gt 0 -and $Actual -eq 0) { [System.Drawing.Color]::FromArgb(255, 198, 239, 206) }
        else { [System.Drawing.Color]::LightGray }
    } elseif ($percent -gt 100) { [System.Drawing.Color]::FromArgb(255, 255, 199, 206) }
    else { [System.Drawing.Color]::FromArgb(255, 198, 239, 206) }
    
    $Cell.Style.Fill.BackgroundColor.SetColor($bgColor)
    if ($bgColor -eq [System.Drawing.Color]::Red) {
        $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    }
}

# Utility to safely convert to double
function Get-SafeDouble { param($Value) [double]($Value ?? 0) }



# Cost family metadata
$costFamilies = @(
    @{ Key = 'HoursCost'; Title = 'Labor %' },
    @{ Key = 'MaterialCost'; Title = 'Material %' },
    @{ Key = 'EquipmentCost'; Title = 'Equipment %' },
    @{ Key = 'SubcontractCost'; Title = 'Subcontract %' },
    @{ Key = 'OtherCost'; Title = 'Other %' },
    @{ Key = 'AdministrativeCost'; Title = 'Administrative %' }
)

# Consolidated function to add cost data row
function Add-CostDataRow {
    param($Sheet, [int]$Row, $JobData, [int]$StartCol, [bool]$Bold = $false, $BgColor = $null)
    
    $col = $StartCol
    foreach ($fam in $costFamilies) {
        $obj = $JobData.($fam.Key) ?? @{ Actual = 0; Budget = 0; Remaining = 0 }
        $actual = Get-SafeDouble $obj.Actual
        $budget = Get-SafeDouble $obj.Budget
        $remaining = Get-SafeDouble $obj.Remaining
        
        # Percentage column (visible when collapsed)
        Set-PercentCell $Sheet.Cells.Item($Row, $col) $actual $budget $remaining $Bold
        
        # Detail columns (collapsible)
        Set-CellData $Sheet.Cells.Item($Row, $col + 1) $actual -Type "Currency" -Bold $Bold -BgColor $BgColor
        Set-CellData $Sheet.Cells.Item($Row, $col + 2) $budget -Type "Currency" -Bold $Bold -BgColor $BgColor
        Set-CellData $Sheet.Cells.Item($Row, $col + 3) $remaining -Type "Currency" -Bold $Bold -BgColor $BgColor
        
        $col += 5  # Move to next cost type
    }
}

# Simplified column grouping
function Set-ColumnGroups {
    param($Sheet, $Groups, [bool]$Collapsed = $true)
    
    try {
        foreach ($group in $Groups) {
            # Ensure we have a valid array with exactly 2 elements
            if ($group -is [array] -and $group.Count -eq 2) {
                $startCol = [int]$group[0]
                $endCol = [int]$group[1]
                
                # Validate column range
                if ($startCol -ge 1 -and $endCol -ge $startCol -and $endCol -le 1000) {
                    for ($col = $startCol; $col -le $endCol; $col++) {
                        try {
                            $Sheet.Column($col).OutlineLevel = 1
                            if ($Collapsed) { $Sheet.Column($col).Collapsed = $true }
                        } catch {
                            Write-Warning "Unable to group column $col : $($_.Exception.Message)"
                        }
                    }
                } else {
                    Write-Warning "Invalid column range: $startCol to $endCol (from group: $($group -join ','))"
                }
            } else {
                Write-Warning "Invalid group format: $($group -join ',') (expected 2-element array)"
            }
        }
    } catch {
        Write-Warning "Unable to create column groups: $($_.Exception.Message)"
    }
}

# Simplified row grouping
function Set-RowGroups {
    param($Sheet, [int]$StartRow, [int]$EndRow, [bool]$Collapsed = $true)
    
    try {
        ($StartRow..$EndRow) | ForEach-Object {
            $Sheet.Row($_).OutlineLevel = 1
            if ($Collapsed) { $Sheet.Row($_).Collapsed = $true }
        }
    } catch {
        Write-Warning "Unable to create row groups: $($_.Exception.Message)"
    }
}

function New-PMSummarySheet {
    param($Excel, $JobSummaryData, $PMName, $PMJobs, $Group)
    
    $summarySheet = $Excel.Workbook.Worksheets.Add("Summary")
    $formattedDate = [DateTime]::ParseExact($timestamp, "yyyyMMdd", $null).ToString("MM/dd/yyyy")
    
    # Title
    Set-CellData $summarySheet.Cells["A1"] "Cost Summary - $formattedDate" -Bold $true
    $summarySheet.Cells["A1"].Style.Font.Size = 14
    
    # PM Name on separate row
    Set-CellData $summarySheet.Cells["A2"] $PMName -Bold $true
    $summarySheet.Cells["A2"].Style.Font.Size = 12
    
    # Build headers
    $costHeaders = $costFamilies | ForEach-Object { 
        @($_.Title, ($_.Title -replace ' %', ' Cost'), ($_.Title -replace ' %', ' Budget'), ($_.Title -replace ' %', ' Remaining'), '')
    }
    $headers = @('Job Number','Job Name') + $costHeaders + @('Project %', 'Contract Amount', 'Project Budget', 'Project Cost', 'Est. Profit')
    
    # Merge title rows and add headers
    $summarySheet.Cells[1,1,1,$headers.Count].Merge = $true
    $summarySheet.Cells[2,1,2,$headers.Count].Merge = $true
    for ($i = 0; $i -lt $headers.Count; $i++) {
        Set-CellData $summarySheet.Cells.Item(3, $i+1) $headers[$i] -Bold $true -BgColor ([System.Drawing.Color]::Black) -FontColor ([System.Drawing.Color]::White)
    }
    
    # Calculate project section columns first
    [int]$projectStartCol = 3 + ($costFamilies.Count * 5)  # Should be column 28 (Project %)
    [int]$contractCol = $projectStartCol + 1              # Column 29 - Contract Amount  
    [int]$profitCol = $projectStartCol + 4                # Column 32 - Est. Profit
    
    # Create column groups for cost sections (include gap columns so they collapse together)
    $summaryColumnGroups = @(
        @(4,7),   # Labor: Cost, Budget, Remaining + Gap
        @(9,12),  # Material: Cost, Budget, Remaining + Gap  
        @(14,17), # Equipment: Cost, Budget, Remaining + Gap
        @(19,22), # Subcontract: Cost, Budget, Remaining + Gap
        @(24,27)  # Other: Cost, Budget, Remaining + Gap
    )
    Set-ColumnGroups $summarySheet $summaryColumnGroups $true  # Cost sections collapsed
    
    # Project section columns (expanded by default) - handle separately to set different collapse state
    if ($profitCol -le $headers.Count) {
        # Project detail group: Contract Amount, Project Budget, Project Cost, Est. Profit
        # Group all project detail columns except Project % (which stays visible)
        try {
            for ($col = $contractCol; $col -le $profitCol; $col++) {
                $summarySheet.Column($col).OutlineLevel = 1
                # Expanded by default (don't set Collapsed = true)
            }
        } catch {
            Write-Warning "Failed to create project section grouping: $($_.Exception.Message)"
        }
    }
    
    # Initialize totals
    $pmTotals = @{ ContractAmount = 0; ProjectBudget = 0; ProjectCost = 0; EstProfit = 0 }
         $costTotals = @{}
     $costFamilies | ForEach-Object { $costTotals[$_.Key] = @{ Actual = 0; Budget = 0; Remaining = 0 } }
     
     [int]$currentRow = 4
    
    # Process jobs
    foreach ($jobId in $PMJobs) {
        $job = $JobSummaryData."$jobId"
        if (-not $job) { continue }
        
        # Job info with hyperlink
        Set-CellData $summarySheet.Cells.Item($currentRow, 1) $jobId
        $worksheetName = if ($jobId.Length -gt 31) { $jobId.Substring(0, 31) } else { $jobId }
        $summarySheet.Cells.Item($currentRow, 1).Hyperlink = New-Object OfficeOpenXml.ExcelHyperLink("'$worksheetName'!A1", $jobId)
        $summarySheet.Cells.Item($currentRow, 1).Style.Font.UnderLine = $true
        $summarySheet.Cells.Item($currentRow, 1).Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        Set-CellData $summarySheet.Cells.Item($currentRow, 2) $job.JobName
        
        # Add cost data
        Add-CostDataRow $summarySheet $currentRow $job 3
        
        # Project data
        $projectBudget = Get-SafeDouble $job.ProjectCost.Budget
        $projectActual = Get-SafeDouble $job.ProjectCost.Actual
        $contractAmt = Get-SafeDouble $job.ContractAmount
        $estProfitVal = Get-SafeDouble $job.EstProfit
        
        # Use pre-calculated column positions instead of IndexOf
        [int]$projPercentCol = $projectStartCol     # Project %
        [int]$contractCol = $projectStartCol + 1    # Contract Amount
        [int]$projBudgetCol = $projectStartCol + 2  # Project Budget
        [int]$projCostCol = $projectStartCol + 3    # Project Cost  
        [int]$profitCol = $projectStartCol + 4      # Est. Profit
        
        # Project percentage
        Set-PercentCell $summarySheet.Cells.Item($currentRow, $projPercentCol) $projectActual $projectBudget ($projectBudget - $projectActual)
        
        # Project details
        Set-CellData $summarySheet.Cells.Item($currentRow, $contractCol) $contractAmt -Type "Currency"
        Set-CellData $summarySheet.Cells.Item($currentRow, $projBudgetCol) $projectBudget -Type "Currency"
        Set-CellData $summarySheet.Cells.Item($currentRow, $projCostCol) $projectActual -Type "Currency"
        Set-CellData $summarySheet.Cells.Item($currentRow, $profitCol) $estProfitVal -Type "Currency"
        
        # Accumulate totals
        $pmTotals.ContractAmount += $contractAmt
        $pmTotals.ProjectBudget += $projectBudget
        $pmTotals.ProjectCost += $projectActual
        $pmTotals.EstProfit += $estProfitVal
        
        $costFamilies | ForEach-Object {
            $costObj = $job.($_.Key)
            if ($costObj) {
                $costTotals[$_.Key].Actual += Get-SafeDouble $costObj.Actual
                $costTotals[$_.Key].Budget += Get-SafeDouble $costObj.Budget
                $costTotals[$_.Key].Remaining += Get-SafeDouble $costObj.Remaining
            }
        }
        
        $currentRow++
    }
    
    # Totals row
    $grey = [System.Drawing.Color]::FromArgb(255, 217, 217, 217)
    Set-CellData $summarySheet.Cells.Item($currentRow, 1) "TOTAL" -Bold $true -BgColor $grey
    Set-CellData $summarySheet.Cells.Item($currentRow, 2) "" -Bold $true -BgColor $grey
    
    # Cost family totals
    Add-CostDataRow $summarySheet $currentRow $costTotals 3 $true $grey
    
    # Project totals
    Set-PercentCell $summarySheet.Cells.Item($currentRow, $projPercentCol) $pmTotals.ProjectCost $pmTotals.ProjectBudget ($pmTotals.ProjectBudget - $pmTotals.ProjectCost) $true
    Set-CellData $summarySheet.Cells.Item($currentRow, $contractCol) $pmTotals.ContractAmount -Type "Currency" -Bold $true -BgColor $grey
    Set-CellData $summarySheet.Cells.Item($currentRow, $projBudgetCol) $pmTotals.ProjectBudget -Type "Currency" -Bold $true -BgColor $grey
    Set-CellData $summarySheet.Cells.Item($currentRow, $projCostCol) $pmTotals.ProjectCost -Type "Currency" -Bold $true -BgColor $grey
    Set-CellData $summarySheet.Cells.Item($currentRow, $profitCol) $pmTotals.EstProfit -Type "Currency" -Bold $true -BgColor $grey
    
    # Freeze panes and autofit
    $summarySheet.View.FreezePanes(4, 3)
    $summarySheet.Cells.AutoFitColumns()
}

# Consolidated function for all cost data rows in detailed sheets
function Add-AllCostDataRow {
    param($Worksheet, [int]$Row, $DataObj, [string]$Label1 = "", [string]$Label2 = "", 
          [bool]$Bold = $false, $BgColor = $null, [bool]$CenterAlign = $false)
    
    # Labels
    if ($Label1) { 
        Set-CellData $Worksheet.Cells.Item($Row, 1) $Label1 -Bold $Bold -BgColor $BgColor
        if ($CenterAlign) { 
            $Worksheet.Cells.Item($Row, 1).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center 
        } elseif ($Label1.Trim() -match '^\d+$') {
            # Right-align phase and category numbers (numeric identifiers)
            $Worksheet.Cells.Item($Row, 1).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
        }
    }
    if ($Label2) { 
        Set-CellData $Worksheet.Cells.Item($Row, 2) $Label2 -Bold $Bold -BgColor $BgColor
        if ($CenterAlign) { $Worksheet.Cells.Item($Row, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
    }
    
    # Cost types with their column positions
    $costTypes = @(
        @{ Name = "Hours"; StartCol = 3; Type = "Hours" },
        @{ Name = "HoursCost"; StartCol = 8; Type = "Currency" },
        @{ Name = "MaterialCost"; StartCol = 13; Type = "Currency" },
        @{ Name = "SubcontractCost"; StartCol = 18; Type = "Currency" },
        @{ Name = "EquipmentCost"; StartCol = 23; Type = "Currency" },
        @{ Name = "OtherCost"; StartCol = 28; Type = "Currency" }
    )
    
    foreach ($costType in $costTypes) {
        $costObj = $DataObj.($costType.Name) ?? @{ Actual = 0; Budget = 0; Remaining = 0 }
        $col = $costType.StartCol
        
        $actual = Get-SafeDouble $costObj.Actual
        $budget = Get-SafeDouble $costObj.Budget
        $remaining = Get-SafeDouble $costObj.Remaining
        
        # Percentage column first
        Set-PercentCell $Worksheet.Cells.Item($Row, $col) $actual $budget $remaining $Bold
        
        # Detail columns
        Set-CellData $Worksheet.Cells.Item($Row, $col + 1) $actual -Type $costType.Type -Bold $Bold -BgColor $BgColor
        Set-CellData $Worksheet.Cells.Item($Row, $col + 2) $budget -Type $costType.Type -Bold $Bold -BgColor $BgColor
        Set-CellData $Worksheet.Cells.Item($Row, $col + 3) $remaining -Type $costType.Type -Bold $Bold -BgColor $BgColor
        
        if ($CenterAlign) {
            ($col..($col + 3)) | ForEach-Object {
                $Worksheet.Cells.Item($Row, $_).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            }
        }
    }
}

function New-ReportSheet {
    param($JobId, $JobData, $Worksheet, [switch]$ReturnData, [switch]$IncludeSummaryLink)
    
    $grey = [System.Drawing.Color]::FromArgb(255, 217, 217, 217)
    [int]$currentRow = if ($IncludeSummaryLink) { 5 } else { 4 }
    
    # Return to Summary link
    if ($IncludeSummaryLink) {
        $Worksheet.Cells["A1"].Value = "Return to Summary"
        $Worksheet.Cells["A1"].Hyperlink = New-Object OfficeOpenXml.ExcelHyperLink("'Summary'!A1", "Return to Summary")
        $Worksheet.Cells["A1"].Style.Font.UnderLine = $true
        $Worksheet.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
    }
    
    # Calculate summary values
    $contractAmount = Get-SafeDouble $JobData.ContractAmount
    $projectCostToUse = [Math]::Max((Get-SafeDouble $JobData.ProjectCost.Actual), (Get-SafeDouble $JobData.ProjectCost.Budget))
    $estProfit = if ($contractAmount -gt 0) { [Math]::Round($contractAmount - $projectCostToUse, 2) } else { 0 }
    
    $hourlyActualRate = if ($JobData.Hours.Actual -gt 0) { [Math]::Round($JobData.HoursCost.Actual / $JobData.Hours.Actual, 2) } else { 0 }
    $hourlyBudgetRate = if ($JobData.Hours.Budget -gt 0) { [Math]::Round($JobData.HoursCost.Budget / $JobData.Hours.Budget, 2) } else { 0 }
    
    # Title section
    $formattedDate = [DateTime]::ParseExact($timestamp, "yyyyMMdd", $null).ToString("MM/dd/yyyy")
    [int]$titleRow = if ($IncludeSummaryLink) { 2 } else { 1 }
    
    # Create merged title rows
    $Worksheet.Cells.Item($titleRow, 1, $titleRow, 3).Merge = $true
    Set-CellData $Worksheet.Cells.Item($titleRow, 1) "$JobId - $formattedDate" -Bold $true
    $Worksheet.Cells.Item($titleRow, 1).Style.Font.Size = 14
    
    $Worksheet.Cells.Item(($titleRow + 1), 1, ($titleRow + 1), 3).Merge = $true
    Set-CellData $Worksheet.Cells.Item(($titleRow + 1), 1) $JobData.JobName -Bold $true
    $Worksheet.Cells.Item(($titleRow + 1), 1).Style.Font.Size = 12
    
    $Worksheet.Cells.Item(($titleRow + 2), 1, ($titleRow + 2), 3).Merge = $true
    Set-CellData $Worksheet.Cells.Item(($titleRow + 2), 1) "Project Manager: $($JobData.PM)" -Bold $true
    
    $currentRow++
    
        # Project summary section with collapsible details
    
    # Project % row (visible, shown first)
    Set-CellData $Worksheet.Cells.Item($currentRow, 1) "Project %" -Bold $true
    Set-PercentCell $Worksheet.Cells.Item($currentRow, 2) (Get-SafeDouble $JobData.ProjectCost.Actual) (Get-SafeDouble $JobData.ProjectCost.Budget) (Get-SafeDouble $JobData.ProjectCost.Remaining) $true
    $Worksheet.Cells.Item($currentRow, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $currentRow++
    
    $projDetailStartRow = $currentRow
    
    # Detail rows (will be collapsed)
    @(
        @("Contract Amount", $contractAmount),
        @("Project Budget", (Get-SafeDouble $JobData.ProjectCost.Budget)),
        @("Project Cost", (Get-SafeDouble $JobData.ProjectCost.Actual)),
        @("Est. Profit", $estProfit)
    ) | ForEach-Object {
        Set-CellData $Worksheet.Cells.Item($currentRow, 1) $_[0] -Bold $true
        Set-CellData $Worksheet.Cells.Item($currentRow, 2) $_[1] -Type "Currency" -Bold $true
        $Worksheet.Cells.Item($currentRow, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $currentRow++
    }
    
    # Collapse project detail rows
    Set-RowGroups $Worksheet $projDetailStartRow ($currentRow)
    $currentRow += 1  # Gap after project %
    
        # Hourly rate section with collapsible details
    
    # Hourly % row (visible, shown first)
    Set-CellData $Worksheet.Cells.Item($currentRow, 1) "Hourly %" -Bold $true
    Set-PercentCell $Worksheet.Cells.Item($currentRow, 2) $hourlyActualRate $hourlyBudgetRate ($hourlyBudgetRate - $hourlyActualRate) $true
    $Worksheet.Cells.Item($currentRow, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $currentRow++
    
    $hourlyDetailStartRow = $currentRow
    
    # Hourly detail rows (will be collapsed)
    Set-CellData $Worksheet.Cells.Item($currentRow, 1) "Hourly Cost" -Bold $true
    Set-CellData $Worksheet.Cells.Item($currentRow, 2) $hourlyActualRate -Type "Currency" -Bold $true
    $Worksheet.Cells.Item($currentRow, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $currentRow++
    
    Set-CellData $Worksheet.Cells.Item($currentRow, 1) "Hourly Budget" -Bold $true
    Set-CellData $Worksheet.Cells.Item($currentRow, 2) $hourlyBudgetRate -Type "Currency" -Bold $true
    $Worksheet.Cells.Item($currentRow, 2).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $currentRow++

    # Collapse hourly detail rows
    Set-RowGroups $Worksheet $hourlyDetailStartRow ($currentRow)
    $currentRow += 1
    
    # Grand Total row
    Add-AllCostDataRow $Worksheet $currentRow $JobData "" "" $true $grey $true
    # Ensure column B (next to Grand Total) also has gray background
    # Set-CellData $Worksheet.Cells.Item($currentRow, 2) "" -Bold $true -BgColor $grey
    $currentRow++
    
    # Main table headers
    $mainHeaders = @(
        "", "Row Labels",
        "Hours %", "Hours", "Hours Budget", "Hours Remaining", "",
        "Labor %", "Labor Cost", "Labor Budget", "Labor Remaining", "",
        "Material %", "Material Cost", "Material Budget", "Material Remaining", "",
        "Subcontract %", "Subcontract Cost", "Subcontract Budget", "Subcontract Remaining", "",
        "Equipment %", "Equipment Cost", "Equipment Budget", "Equipment Remaining", "",
        "Other %", "Other Cost", "Other Budget", "Other Remaining"
    )
    
    for ($i = 0; $i -lt $mainHeaders.Count; $i++) {
        Set-CellData $Worksheet.Cells.Item($currentRow, $i+1) $mainHeaders[$i] -Bold $true -BgColor ([System.Drawing.Color]::Black) -FontColor ([System.Drawing.Color]::White)
    }
    $currentRow++
    
    # Create collapsible column groups (include gap columns so they collapse together)
    $columnGroups = @(
        @(4,7),   # Hours: Hours, Hours Budget, Hours Remaining + Gap
        @(9,12),  # Labor: Labor Cost, Labor Budget, Labor Remaining + Gap
        @(14,17), # Material: Material Cost, Material Budget, Material Remaining + Gap
        @(19,22), # Subcontract: Subcontract Cost, Subcontract Budget, Subcontract Remaining + Gap
        @(24,27), # Equipment: Equipment Cost, Equipment Budget, Equipment Remaining + Gap
        @(29,31)  # Other: Other Cost, Other Budget, Other Remaining (no gap after last)
    )
    
    # Filter out any groups that exceed the actual column count
    $validGroups = $columnGroups | Where-Object { $_[1] -le $mainHeaders.Count }
    Set-ColumnGroups $Worksheet $validGroups
    
    # Freeze panes
    try { $Worksheet.View.FreezePanes($currentRow, 3) } 
    catch { Write-Warning "Unable to freeze panes: $($_.Exception.Message)" }
    
    # Process phases and categories
    if ($JobData.Phases) {
        foreach ($phaseKey in $JobData.Phases.PSObject.Properties.Name) {
            if ($phaseKey -eq "000") { continue }  # Skip phase 000
            
            $phase = $JobData.Phases.$phaseKey
            Add-AllCostDataRow $Worksheet $currentRow $phase $phaseKey $phase.PhaseName $true $grey
            $currentRow++
            $phaseStartRow = $currentRow
            
            # Categories
            if ($phase.Categories) {
                foreach ($catKey in $phase.Categories.PSObject.Properties.Name) {
                    $category = $phase.Categories.$catKey
                    Add-AllCostDataRow $Worksheet $currentRow $category "   $catKey" "   $($category.CategoryName)"
                    $currentRow++
                }
                
                # Collapse category rows under this phase
                if ($currentRow -gt $phaseStartRow) {
                    Set-RowGroups $Worksheet $phaseStartRow ($currentRow )
                }
            }
            $currentRow++
        }
    }
    
    # Auto-fit columns
    $Worksheet.Cells.AutoFitColumns()
    
    if ($ReturnData) {
        return @{
            JobId = $JobId; JobName = $JobData.JobName; PM = $JobData.PM
            Hours = $JobData.Hours; HoursCost = $JobData.HoursCost; ProjectCost = $JobData.ProjectCost
            ContractAmount = $contractAmount; EstProfit = $estProfit
        }
    }
}

# Helper functions for cleaner code
function Get-SafeFileName { param($Name) $Name -replace '[^\w\s-]', '' -replace '\s+', '_' }
function Get-UniqueWorksheetName { 
    param($Package, $Name)
    $wsName = if ($Name.Length -gt 31) { $Name.Substring(0, 31) } else { $Name }
    $counter = 1
    while ($Package.Workbook.Worksheets[$wsName]) {
        $suffix = "_$counter"
        $maxLength = 31 - $suffix.Length
        $wsName = $Name.Substring(0, [Math]::Min($Name.Length, $maxLength)) + $suffix
        $counter++
    }
    return $wsName
}

function New-ExcelFile {
    param($JobId, $JobData, $OutputPath, [switch]$IncludeSummaryLink)
    
    $package = New-Object OfficeOpenXml.ExcelPackage
    try {
        $ws = $package.Workbook.Worksheets.Add($JobId)
        New-ReportSheet -JobId $JobId -JobData $JobData -Worksheet $ws -IncludeSummaryLink:$IncludeSummaryLink
        $package.SaveAs([System.IO.FileInfo]$OutputPath)
        return $true
    } catch {
        Write-Warning "Failed to create $OutputPath : $($_.Exception.Message)"
        return $false
    } finally {
        $package.Dispose()
    }
}

function New-GrandSummaryFile {
    param($JobData, $PMGroups, $Group, $OutputPath)
    
    $package = New-Object OfficeOpenXml.ExcelPackage
    try {
        # Create Grand Summary sheet
        $summarySheet = $package.Workbook.Worksheets.Add("Summary")
        $formattedDate = [DateTime]::ParseExact($timestamp, "yyyyMMdd", $null).ToString("MM/dd/yyyy")
        
        # Calculate total job count for title
        $totalJobCount = ($PMGroups.GetEnumerator() | ForEach-Object { $_.Value.Count } | Measure-Object -Sum).Sum
        
        # Title
        Set-CellData $summarySheet.Cells["A1"] "All PMs Cost Summary ($totalJobCount Jobs)" -Bold $true
        $summarySheet.Cells["A1"].Style.Font.Size = 14
        
        # Group and Date on separate row
        Set-CellData $summarySheet.Cells["A2"] "$Group - $formattedDate" -Bold $true
        $summarySheet.Cells["A2"].Style.Font.Size = 12
        
        # Build headers
        $costHeaders = $costFamilies | ForEach-Object { 
            @($_.Title, ($_.Title -replace ' %', ' Cost'), ($_.Title -replace ' %', ' Budget'), ($_.Title -replace ' %', ' Remaining'), '')
        }
        $headers = @('Job Number','Job Name') + $costHeaders + @('Project %', 'Contract Amount', 'Project Budget', 'Project Cost', 'Est. Profit')
        
        # Merge title rows and add headers
        $summarySheet.Cells[1,1,1,$headers.Count].Merge = $true
        $summarySheet.Cells[2,1,2,$headers.Count].Merge = $true
        for ($i = 0; $i -lt $headers.Count; $i++) {
            Set-CellData $summarySheet.Cells.Item(3, $i+1) $headers[$i] -Bold $true -BgColor ([System.Drawing.Color]::Black) -FontColor ([System.Drawing.Color]::White)
        }
        
        # Calculate project section columns
        [int]$projectStartCol = 3 + ($costFamilies.Count * 5)  # Project %
        [int]$contractCol = $projectStartCol + 1              # Contract Amount  
        [int]$profitCol = $projectStartCol + 4                # Est. Profit
        
        # Create column groups for cost sections
        $summaryColumnGroups = @(
            @(4,7),   # Labor: Cost, Budget, Remaining + Gap
            @(9,12),  # Material: Cost, Budget, Remaining + Gap  
            @(14,17), # Equipment: Cost, Budget, Remaining + Gap
            @(19,22), # Subcontract: Cost, Budget, Remaining + Gap
            @(24,27)  # Other: Cost, Budget, Remaining + Gap
        )
        Set-ColumnGroups $summarySheet $summaryColumnGroups $true
        
        # Project section columns (expanded by default)
        try {
            for ($col = $contractCol; $col -le $profitCol; $col++) {
                $summarySheet.Column($col).OutlineLevel = 1
            }
        } catch {
            Write-Warning "Failed to create project section grouping: $($_.Exception.Message)"
        }
        
        # Initialize grand totals
        $grandTotals = @{ ContractAmount = 0; ProjectBudget = 0; ProjectCost = 0; EstProfit = 0; JobCount = 0 }
        $grandCostTotals = @{}
        $costFamilies | ForEach-Object { $grandCostTotals[$_.Key] = @{ Actual = 0; Budget = 0; Remaining = 0 } }
        
        [int]$currentRow = 4
        
        # Process each PM with their jobs
        $PMGroups.GetEnumerator() | Sort-Object Key | ForEach-Object {
            $pmName, $pmJobs = $_.Key, $_.Value
            
            # Initialize PM totals
            $pmTotals = @{ ContractAmount = 0; ProjectBudget = 0; ProjectCost = 0; EstProfit = 0 }
            $pmCostTotals = @{}
            $costFamilies | ForEach-Object { $pmCostTotals[$_.Key] = @{ Actual = 0; Budget = 0; Remaining = 0 } }
            
            # First pass: Calculate PM totals
            foreach ($jobId in $pmJobs) {
                $job = $JobData."$jobId"
                if (-not $job) { continue }
                
                # Project totals
                $contractAmt = Get-SafeDouble $job.ContractAmount
                $projectBudget = Get-SafeDouble $job.ProjectCost.Budget
                $projectActual = Get-SafeDouble $job.ProjectCost.Actual
                $estProfitVal = Get-SafeDouble $job.EstProfit
                
                # Accumulate PM totals
                $pmTotals.ContractAmount += $contractAmt
                $pmTotals.ProjectBudget += $projectBudget
                $pmTotals.ProjectCost += $projectActual
                $pmTotals.EstProfit += $estProfitVal
                
                $costFamilies | ForEach-Object {
                    $costObj = $job.($_.Key)
                    if ($costObj) {
                        $pmCostTotals[$_.Key].Actual += Get-SafeDouble $costObj.Actual
                        $pmCostTotals[$_.Key].Budget += Get-SafeDouble $costObj.Budget
                        $pmCostTotals[$_.Key].Remaining += Get-SafeDouble $costObj.Remaining
                    }
                }
            }
            
            # PM Total row (displayed first)
            $lightGrey = [System.Drawing.Color]::FromArgb(255, 230, 230, 230)
            Set-CellData $summarySheet.Cells.Item($currentRow, 1) "TOTAL" -Bold $true -BgColor $lightGrey
            Set-CellData $summarySheet.Cells.Item($currentRow, 2) $pmName -Bold $true -BgColor $lightGrey
            
            # PM cost totals
            Add-CostDataRow $summarySheet $currentRow $pmCostTotals 3 $true $lightGrey
            
            # PM project totals
            [int]$projPercentCol = $projectStartCol
            [int]$contractColPos = $projectStartCol + 1
            [int]$projBudgetCol = $projectStartCol + 2
            [int]$projCostCol = $projectStartCol + 3
            [int]$profitColPos = $projectStartCol + 4
            
            Set-PercentCell $summarySheet.Cells.Item($currentRow, $projPercentCol) $pmTotals.ProjectCost $pmTotals.ProjectBudget ($pmTotals.ProjectBudget - $pmTotals.ProjectCost) $true
            Set-CellData $summarySheet.Cells.Item($currentRow, $contractColPos) $pmTotals.ContractAmount -Type "Currency" -Bold $true -BgColor $lightGrey
            Set-CellData $summarySheet.Cells.Item($currentRow, $projBudgetCol) $pmTotals.ProjectBudget -Type "Currency" -Bold $true -BgColor $lightGrey
            Set-CellData $summarySheet.Cells.Item($currentRow, $projCostCol) $pmTotals.ProjectCost -Type "Currency" -Bold $true -BgColor $lightGrey
            Set-CellData $summarySheet.Cells.Item($currentRow, $profitColPos) $pmTotals.EstProfit -Type "Currency" -Bold $true -BgColor $lightGrey
            $currentRow++
            
            # Track job rows for grouping (jobs come after TOTAL)
            $jobRowStart = $currentRow
            
            # Second pass: Display job details (collapsible)
            foreach ($jobId in $pmJobs) {
                $job = $JobData."$jobId"
                if (-not $job) { continue }
                
                # Job info with hyperlink
                Set-CellData $summarySheet.Cells.Item($currentRow, 1) $jobId
                $worksheetName = Get-UniqueWorksheetName $package $jobId
                $summarySheet.Cells.Item($currentRow, 1).Hyperlink = New-Object OfficeOpenXml.ExcelHyperLink("'$worksheetName'!A1", $jobId)
                $summarySheet.Cells.Item($currentRow, 1).Style.Font.UnderLine = $true
                $summarySheet.Cells.Item($currentRow, 1).Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                Set-CellData $summarySheet.Cells.Item($currentRow, 2) $job.JobName
                
                # Add cost data for this job
                Add-CostDataRow $summarySheet $currentRow $job 3
                
                # Project data
                $projectBudget = Get-SafeDouble $job.ProjectCost.Budget
                $projectActual = Get-SafeDouble $job.ProjectCost.Actual
                $contractAmt = Get-SafeDouble $job.ContractAmount
                $estProfitVal = Get-SafeDouble $job.EstProfit
                
                Set-PercentCell $summarySheet.Cells.Item($currentRow, $projPercentCol) $projectActual $projectBudget ($projectBudget - $projectActual)
                Set-CellData $summarySheet.Cells.Item($currentRow, $contractColPos) $contractAmt -Type "Currency"
                Set-CellData $summarySheet.Cells.Item($currentRow, $projBudgetCol) $projectBudget -Type "Currency"
                Set-CellData $summarySheet.Cells.Item($currentRow, $projCostCol) $projectActual -Type "Currency"
                Set-CellData $summarySheet.Cells.Item($currentRow, $profitColPos) $estProfitVal -Type "Currency"
                
                $currentRow++
            }
            
            # Group the job rows to make them collapsible under the TOTAL row
            if ($currentRow -gt $jobRowStart) {
                Set-RowGroups $summarySheet $jobRowStart ($currentRow) $true
            }
            
            # Accumulate grand totals
            $grandTotals.ContractAmount += $pmTotals.ContractAmount
            $grandTotals.ProjectBudget += $pmTotals.ProjectBudget
            $grandTotals.ProjectCost += $pmTotals.ProjectCost
            $grandTotals.EstProfit += $pmTotals.EstProfit
            $grandTotals.JobCount += $pmJobs.Count
            
            $costFamilies | ForEach-Object {
                $grandCostTotals[$_.Key].Actual += $pmCostTotals[$_.Key].Actual
                $grandCostTotals[$_.Key].Budget += $pmCostTotals[$_.Key].Budget
                $grandCostTotals[$_.Key].Remaining += $pmCostTotals[$_.Key].Remaining
            }
            
            $currentRow++  # Single space between PMs
        }
        
        # Grand Totals row
        $grey = [System.Drawing.Color]::FromArgb(255, 217, 217, 217)
        Set-CellData $summarySheet.Cells.Item($currentRow, 1) "GRAND TOTAL" -Bold $true -BgColor $grey
        Set-CellData $summarySheet.Cells.Item($currentRow, 2) "" -Bold $true -BgColor $grey
        
        # Cost family grand totals
        Add-CostDataRow $summarySheet $currentRow $grandCostTotals 3 $true $grey
        
        # Project grand totals
        Set-PercentCell $summarySheet.Cells.Item($currentRow, $projPercentCol) $grandTotals.ProjectCost $grandTotals.ProjectBudget ($grandTotals.ProjectBudget - $grandTotals.ProjectCost) $true
        Set-CellData $summarySheet.Cells.Item($currentRow, $contractColPos) $grandTotals.ContractAmount -Type "Currency" -Bold $true -BgColor $grey
        Set-CellData $summarySheet.Cells.Item($currentRow, $projBudgetCol) $grandTotals.ProjectBudget -Type "Currency" -Bold $true -BgColor $grey
        Set-CellData $summarySheet.Cells.Item($currentRow, $projCostCol) $grandTotals.ProjectCost -Type "Currency" -Bold $true -BgColor $grey
        Set-CellData $summarySheet.Cells.Item($currentRow, $profitColPos) $grandTotals.EstProfit -Type "Currency" -Bold $true -BgColor $grey
        
        # Freeze panes and autofit
        $summarySheet.View.FreezePanes(4, 3)
        $summarySheet.Cells.AutoFitColumns()
        
        # Add all job sheets
        $JobData.PSObject.Properties | Where-Object { $_.Value.PM -and $_.Value.PM.Trim() } | Sort-Object Name | ForEach-Object {
            $jobId = $_.Name
            $jobObj = $_.Value
            $wsName = Get-UniqueWorksheetName $package $jobId
            $ws = $package.Workbook.Worksheets.Add($wsName)
            New-ReportSheet -JobId $jobId -JobData $jobObj -Worksheet $ws -IncludeSummaryLink
        }
        
        # Save grand summary file
        $package.SaveAs([System.IO.FileInfo]$OutputPath)
        
    } finally { 
        $package.Dispose() 
    }
}

function Create-ProjectReports {
    param([Parameter(Mandatory=$true)][string]$JsonPath, [string]$OutputBasePath)

    # Load and validate data
    if (-not (Test-Path $JsonPath)) { Write-Error "JSON path not found: $JsonPath"; return }
    $jobData = Get-Content $JsonPath -Raw | ConvertFrom-Json
    if (-not $jobData) { Write-Error "Failed to parse JSON: $JsonPath"; return }

    # Setup
    $global:timestamp = (Get-Date).ToString("yyyyMMdd")
    if (-not $OutputBasePath) { $OutputBasePath = Split-Path $JsonPath -Parent }
    $group = Split-Path (Split-Path (Split-Path $JsonPath -Parent) -Parent) -Leaf
    
    # Group jobs by PM
    $pmGroups = @{}
    $jobData.PSObject.Properties | Where-Object { $_.Value.PM -and $_.Value.PM.Trim() } | ForEach-Object {
        $pmName = $_.Value.PM.Trim()
        if (-not $pmGroups.ContainsKey($pmName)) { $pmGroups[$pmName] = @() }
        $pmGroups[$pmName] += $_.Name
    }
    
    if (-not $pmGroups -or $pmGroups.Count -eq 0) { 
        Write-Error "No jobs with PM assignments found"; return 
    }
    
    Write-Host "Processing $($pmGroups.Count) PMs..."
    # Debug output
    $pmGroups.GetEnumerator() | ForEach-Object {
        Write-Host "  Debug: PM='$($_.Key)' has $($_.Value.Count) jobs: $($_.Value -join ', ')"
    }
    
    # Process each PM
    $pmGroups.GetEnumerator() | ForEach-Object {
        $pmName, $pmJobs = $_.Key, $_.Value
        
        # Validate PM name and skip if invalid
        if (-not $pmName -or -not $pmName.Trim()) {
            Write-Warning "Skipping PM with empty name"
            return
        }
        
        $pmNameSafe = Get-SafeFileName $pmName
        if (-not $pmNameSafe) {
            Write-Warning "Skipping PM '$pmName' - could not create safe filename"
            return
        }
        
        $pmFolder = Join-Path $OutputBasePath $pmNameSafe
        Write-Host "PM: $pmName ($($pmJobs.Count) jobs)"
        
        # Create PM folder
        New-Item -Path $pmFolder -ItemType Directory -Force | Out-Null
        
        # Create PM summary file
        $package = New-Object OfficeOpenXml.ExcelPackage
        try {
            # PM Summary sheet
            New-PMSummarySheet -Excel $package -JobSummaryData $jobData -PMName $pmName -PMJobs $pmJobs -Group $group
            
            # Add job sheets to PM summary
            $pmJobs | ForEach-Object {
                $jobObj = $jobData."$_"
                if ($jobObj) {
                    $wsName = Get-UniqueWorksheetName $package $_
                    $ws = $package.Workbook.Worksheets.Add($wsName)
                    New-ReportSheet -JobId $_ -JobData $jobObj -Worksheet $ws -IncludeSummaryLink
                }
            }
            
            # Save PM summary
            $summaryFile = Join-Path $pmFolder "CostSummary_${pmNameSafe}_${timestamp}.xlsx"
            $package.SaveAs([System.IO.FileInfo]$summaryFile)
            
        } finally { $package.Dispose() }
        
        # Create individual job files
        $jobCount = ($pmJobs | ForEach-Object {
            $jobObj = $jobData."$_"
            if ($jobObj) {
                $jobFile = Join-Path $pmFolder "$(Get-SafeFileName $_)_${timestamp}.xlsx"
                if (New-ExcelFile $_ $jobObj $jobFile) { 1 } else { 0 }
            }
        } | Measure-Object -Sum).Sum
        
        Write-Host "  Created: PM summary + $jobCount individual files"
    }
    
    # Create Grand Summary file
    $grandSummaryFile = Join-Path $OutputBasePath "GrandSummary_${timestamp}.xlsx"
    New-GrandSummaryFile -JobData $jobData -PMGroups $pmGroups -Group $group -OutputPath $grandSummaryFile
    Write-Host "  Created: Grand Summary file"

    Write-Host "Completed! Output: $OutputBasePath"
}

function Create-AllProjectReports {
    param([Parameter(Mandatory=$true)][string]$ReportsPath, [string]$JsonPattern = "*_JobData_*.json")
    
    Get-ChildItem -Path $ReportsPath -Recurse -Filter $JsonPattern -ErrorAction Stop | 
    Sort-Object LastWriteTime -Descending | 
    ForEach-Object { 
        Write-Host "`nProcessing: $($_.Name)"
        Create-ProjectReports -JsonPath $_.FullName 
    }
}


# Auto-run when executed directly (not dot-sourced)
    Create-AllProjectReports -ReportsPath "C:\WorkV4\ProjectReports\reports"