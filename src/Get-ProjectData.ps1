

function New-BudgetActualObject {
    param($Budget, $Actual)
    return @{
        Budget = [Math]::Round($Budget, 2)
        Actual = [Math]::Round($Actual, 2)
        Remaining = [Math]::Round($Budget - $Actual, 2)
        PercentOfBudget = if ($Budget -gt 0) { [Math]::Round(($Actual / $Budget) * 100, 2) } else { 0 }
    }
}

function Test-JobHasData {
    param($Job)
    
    # Check main cost categories for any non-zero budget or actual values
    $costCategories = @('Hours', 'HoursCost', 'MaterialCost', 'SubcontractCost', 'EquipmentCost', 'OtherCost', 'AdministrativeCost', 'ProjectCost')
    
    foreach ($category in $costCategories) {
        $obj = $Job.$category
        if ($obj -and (($obj.Budget -gt 0) -or ($obj.Actual -gt 0))) {
            return $true
        }
    }
    
    # Check contract amount
    if ($Job.ContractAmount -gt 0) {
        return $true
    }
    
    # If we get here, the job has no meaningful data
    return $false
}

function Test-CategoryHasData {
    param($Category)
    
    # Check only hours and labor cost for any non-zero budget or actual values
    $costCategories = @('Hours', 'HoursCost')
    
    foreach ($costName in $costCategories) {
        $obj = $Category.$costName
        if ($obj -and (($obj.Budget -gt 0) -or ($obj.Actual -gt 0))) {
            return $true
        }
    }
    
    return $false
}

function Test-PhaseHasData {
    param($Phase)
    
    # Check only hours and labor cost for any non-zero budget or actual values
    $costCategories = @('Hours', 'HoursCost')
    
    foreach ($costName in $costCategories) {
        $obj = $Phase.$costName
        if ($obj -and (($obj.Budget -gt 0) -or ($obj.Actual -gt 0))) {
            return $true
        }
    }
    
    return $false
}

# Helper function to safely convert database values to double, handling DBNull
function ConvertTo-SafeDouble {
    param([object]$Value)
    if ($Value -is [DBNull] -or $Value -eq $null) {
        return 0.0
    }
    try {
        return [double]$Value
    }
    catch {
        return 0.0
    }
}

function Export-JobDataToJson {
    param (
        [string]$Group = "",
        [DateTime]$CutoffDate = [DateTime]::MinValue
    )

    # Load and validate config
    $ConfigPath = Join-Path (Join-Path $PSScriptRoot "..") "config" | Join-Path -ChildPath "config.json"
    if (-not (Test-Path $ConfigPath)) { Write-Error "Configuration file not found at $ConfigPath"; return }
    $Config = Get-Content $ConfigPath | ConvertFrom-Json

    # Process all groups if none specified
    if ([string]::IsNullOrEmpty($Group)) {
        Write-Host "Processing all groups..."
        $Config.PSObject.Properties.Name | ForEach-Object { 
            Write-Host "Processing group: $_"
            Export-JobDataToJson -Group $_ -CutoffDate $CutoffDate 
        }
        return
    }

    # Validate group exists
    if (-not $Config.$Group) {
        Write-Error "Group '$Group' not found. Valid groups: $($Config.PSObject.Properties.Name -join ', ')"
        return
    }

    $groupConfig = $Config.$Group
    if (-not $groupConfig.classes -or $groupConfig.classes.Count -eq 0) {
        Write-Host "No job classes configured for '$Group'. Skipping."
        return
    }

    # Setup
    $CutoffDate = if ($CutoffDate -eq [DateTime]::MinValue) { Get-Date } else { $CutoffDate }
    $dateFormatted = $CutoffDate.ToString("yyyyMMdd")
    $outputDir = Join-Path "reports" $Group $dateFormatted
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    $OutputFile = Join-Path $outputDir "${Group}_JobData_$dateFormatted.json"

    # Build query filters
    $classFilter = "'$($groupConfig.classes -join "','")'"
    $companyFilter = if ($groupConfig.companies -and $groupConfig.companies.Count -gt 0) {
        " AND j.company IN ($($groupConfig.companies -join ','))"
    } else { "" }
    $contractCompanyFilter = if ($groupConfig.companies -and $groupConfig.companies.Count -gt 0) {
        " AND jc.Company IN ($($groupConfig.companies -join ','))"
    } else { "" }

    Write-Host "Using cutoff date: $dateFormatted for $Group"
    if ($companyFilter) { Write-Host "Filtering by companies: $($groupConfig.companies -join ', ')" }

    # Execute query
    $conn = New-Object System.Data.SqlClient.SqlConnection("Server=housql01;Database=DataExtract;Integrated Security=SSPI")
    try {
        $conn.Open()
        
        $query = @"
;WITH FilteredDetail AS (
    SELECT jobnum, phasenum, catnum, type, hours, cost
    FROM   jcdetail
    WHERE  date <= '$dateFormatted'
      AND  type IN (1,2,3,4,33,34,35,36,5,6,7,8,9,37,38,39,40,41)
      AND  company IN ($($groupConfig.companies -join ','))
),
DetailAgg AS (
    SELECT jobnum, phasenum, catnum,
           SUM(CASE WHEN type IN (33,34,35,36) THEN hours ELSE 0 END) AS BudgetHours,
           SUM(CASE WHEN type IN (1,2,3,4)     THEN hours ELSE 0 END) AS ActualHours,
           SUM(CASE WHEN type IN (33,34,35,36) THEN cost  ELSE 0 END) AS LaborBudget,
           SUM(CASE WHEN type IN (1,2,3,4)     THEN cost  ELSE 0 END) AS LaborCost,
           SUM(CASE WHEN type = 37                THEN cost  ELSE 0 END) AS MaterialBudget,
           SUM(CASE WHEN type = 5                 THEN cost  ELSE 0 END) AS MaterialCost
          ,SUM(CASE WHEN type = 38                THEN cost  ELSE 0 END) AS SubcontractBudget
          ,SUM(CASE WHEN type = 6                 THEN cost  ELSE 0 END) AS SubcontractCost
          ,SUM(CASE WHEN type = 39                THEN cost  ELSE 0 END) AS EquipmentBudget
          ,SUM(CASE WHEN type = 7                 THEN cost  ELSE 0 END) AS EquipmentCost
          ,SUM(CASE WHEN type = 40                THEN cost  ELSE 0 END) AS OtherBudget
          ,SUM(CASE WHEN type = 8                 THEN cost  ELSE 0 END) AS OtherCost
          ,SUM(CASE WHEN type = 41                THEN cost  ELSE 0 END) AS AdministrativeBudget
          ,SUM(CASE WHEN type = 9                 THEN cost  ELSE 0 END) AS AdministrativeCost
    FROM   FilteredDetail
    GROUP BY jobnum, phasenum, catnum
),
JobTotals AS (
    SELECT jobnum,
           SUM(BudgetHours)   AS TotalBudgetHours,
           SUM(ActualHours)   AS TotalActualHours,
           SUM(LaborBudget)   AS TotalLaborBudget,
           SUM(LaborCost)     AS TotalLaborCost,
           SUM(MaterialBudget) AS TotalMaterialBudget,
           SUM(MaterialCost)   AS TotalMaterialCost
          ,SUM(SubcontractBudget) AS TotalSubcontractBudget
          ,SUM(SubcontractCost)   AS TotalSubcontractCost
          ,SUM(EquipmentBudget) AS TotalEquipmentBudget
          ,SUM(EquipmentCost)   AS TotalEquipmentCost
          ,SUM(OtherBudget) AS TotalOtherBudget
          ,SUM(OtherCost)   AS TotalOtherCost
          ,SUM(AdministrativeBudget) AS TotalAdministrativeBudget
          ,SUM(AdministrativeCost)   AS TotalAdministrativeCost
          -- Calculate total project budget and actual from detail components
          ,(SUM(LaborBudget) + SUM(MaterialBudget) + SUM(SubcontractBudget) + SUM(EquipmentBudget) + SUM(OtherBudget) + SUM(AdministrativeBudget)) AS TotalProjectBudget
          ,(SUM(LaborCost) + SUM(MaterialCost) + SUM(SubcontractCost) + SUM(EquipmentCost) + SUM(OtherCost) + SUM(AdministrativeCost)) AS TotalProjectActual
    FROM   DetailAgg
    GROUP BY jobnum
)
SELECT j.jobnum,
       j.name AS JobName,
       COALESCE(j.pm,'')            AS pm,
       COALESCE(j.AllPayPMNumber,'') AS AllPayPMNumber,
       jt.TotalBudgetHours,
       jt.TotalActualHours,
       jt.TotalLaborBudget,
       jt.TotalLaborCost,
       jt.TotalMaterialBudget,
       jt.TotalMaterialCost,
       jt.TotalSubcontractBudget,
       jt.TotalSubcontractCost,
       jt.TotalEquipmentBudget,
       jt.TotalEquipmentCost,
       jt.TotalOtherBudget,
       jt.TotalOtherCost,
       jt.TotalAdministrativeBudget,
       jt.TotalAdministrativeCost,
       jt.TotalProjectActual        AS TotalProjectCost,
       jt.TotalProjectBudget        AS TotalProjectBudget,
       (SELECT SUM(jc.COST)
        FROM   jcdetail jc
               INNER JOIN jcchangeorderstep jco
                 ON jc.jobnum = jco.jobnum
                AND jc.ponum  = jco.ordernum
        WHERE  jc.TYPE   = 19
          AND  jco.type  = 20
          $contractCompanyFilter
          AND  jc.jobnum = j.jobnum) AS ContractAmount,
       da.phasenum,
       COALESCE(p.name,'Unknown') AS PhaseName,
       da.catnum,
       COALESCE(c.name,'Unknown') AS CatName,
       da.BudgetHours,
       da.ActualHours,
       da.LaborBudget,
       da.LaborCost,
       da.MaterialBudget,
       da.MaterialCost
      ,da.SubcontractBudget
      ,da.SubcontractCost
      ,da.EquipmentBudget
      ,da.EquipmentCost
      ,da.OtherBudget
      ,da.OtherCost
      ,da.AdministrativeBudget
      ,da.AdministrativeCost
FROM   jcjob j
       INNER JOIN DetailAgg da ON j.jobnum = da.jobnum
       INNER JOIN JobTotals jt ON j.jobnum = jt.jobnum
       -- Match phase names by converting phasenum to INT on both sides so zero-padded values (e.g. "001")
       -- in jcphase still join correctly with numeric values coming from jcdetail.
       LEFT  JOIN (
            SELECT DISTINCT jobnum,
                   CAST(phasenum AS INT)              AS phasenum,
                   name
            FROM   jcphase
       ) p  ON  j.jobnum                = p.jobnum
            AND CAST(da.phasenum AS INT) = p.phasenum

       -- Match category names the same way (convert both phase and category numbers to INT)
       LEFT  JOIN (
            SELECT DISTINCT jobnum,
                   CAST(phasenum AS INT)              AS phasenum,
                   CAST(catnum   AS INT)              AS catnum,
                   name
            FROM   jccat
       ) c  ON  j.jobnum                = c.jobnum
            AND CAST(da.phasenum AS INT) = c.phasenum
            AND CAST(da.catnum   AS INT) = c.catnum
WHERE  j.closed = 0
  AND  j.status = 1
  AND  LEFT(j.class, 4) IN ($classFilter)
  $companyFilter
ORDER BY j.jobnum, da.phasenum, da.catnum
"@

        # Use an explicit SqlCommand to allow custom timeout
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $query
        $cmd.CommandTimeout = 300  # seconds (adjust as needed)
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $data = New-Object System.Data.DataSet
        $adapter.Fill($data) | Out-Null
        
        # Build job data structure
        $jobData = [ordered]@{}
        

        
        foreach ($row in $data.Tables[0].Rows) {
            $jobNum = $row["jobnum"]
            $phaseNum = $row["phasenum"]
            $catNum = $row["catnum"]
            
            if ($phaseNum -is [DBNull] -or $catNum -is [DBNull]) { continue }
            
            # Initialize job if not exists
            if (-not $jobData[$jobNum]) {
                # Pre-compute derived values
                $budgetHours = ConvertTo-SafeDouble $row["TotalBudgetHours"]
                $actualHours = ConvertTo-SafeDouble $row["TotalActualHours"]
                $laborBudget = ConvertTo-SafeDouble $row["TotalLaborBudget"]
                $laborCost   = ConvertTo-SafeDouble $row["TotalLaborCost"]
                $materialBudget = ConvertTo-SafeDouble $row["TotalMaterialBudget"]
                $materialCost   = ConvertTo-SafeDouble $row["TotalMaterialCost"]
                $subcontractBudget = ConvertTo-SafeDouble $row["TotalSubcontractBudget"]
                $subcontractCost   = ConvertTo-SafeDouble $row["TotalSubcontractCost"]
                $equipmentBudget = ConvertTo-SafeDouble $row["TotalEquipmentBudget"]
                $equipmentCost   = ConvertTo-SafeDouble $row["TotalEquipmentCost"]
                $otherBudget     = ConvertTo-SafeDouble $row["TotalOtherBudget"]
                $otherCost       = ConvertTo-SafeDouble $row["TotalOtherCost"]
                $administrativeBudget = ConvertTo-SafeDouble $row["TotalAdministrativeBudget"]
                $administrativeCost   = ConvertTo-SafeDouble $row["TotalAdministrativeCost"]

                $budgetRate  = if ($budgetHours -gt 0) { [Math]::Round($laborBudget / $budgetHours, 2) } else { 0 }
                $actualRate  = if ($actualHours -gt 0) { [Math]::Round($laborCost / $actualHours, 2) } else { 0 }

                $projBudget  = ConvertTo-SafeDouble $row["TotalProjectBudget"]
                $projActual  = ConvertTo-SafeDouble $row["TotalProjectCost"]
                $projectCostToUse = if ($projActual -gt $projBudget) { $projActual } else { $projBudget }

                $contractAmt = if ($row["ContractAmount"] -is [DBNull] -or $row["ContractAmount"] -eq $null) { 0 } else { [Math]::Round((ConvertTo-SafeDouble $row["ContractAmount"]), 2) }

                $jobData[$jobNum] = @{
                    JobNumber = $jobNum
                    JobName = $row["JobName"]
                    PM = $row["pm"]
                    AllPayPMNumber = $row["AllPayPMNumber"]
                    Hours = New-BudgetActualObject $budgetHours $actualHours
                    HoursCost = New-BudgetActualObject $laborBudget $laborCost
                    MaterialCost = New-BudgetActualObject $materialBudget $materialCost
                    SubcontractCost = New-BudgetActualObject $subcontractBudget $subcontractCost
                    EquipmentCost = New-BudgetActualObject $equipmentBudget $equipmentCost
                    OtherCost = New-BudgetActualObject $otherBudget $otherCost
                    AdministrativeCost = New-BudgetActualObject $administrativeBudget $administrativeCost
                    ProjectCost = New-BudgetActualObject $projBudget $projActual
                    ContractAmount = $contractAmt
                    HourlyRate = New-BudgetActualObject $budgetRate $actualRate
                    EstProfit = if ($contractAmt -gt 0) { [Math]::Round($contractAmt - $projectCostToUse, 2) } else { 0 }
                    Phases = [ordered]@{}
                }
            }
            
            # Initialize phase if not exists
            if (-not $jobData[$jobNum].Phases[$phaseNum]) {
                $jobData[$jobNum].Phases[$phaseNum] = @{
                    PhaseName = $row["PhaseName"]
                    Categories = [ordered]@{}
                    Hours = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                    HoursCost = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                    MaterialCost = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                    SubcontractCost = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                    EquipmentCost = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                    OtherCost = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                    AdministrativeCost = @{ Budget = 0; Actual = 0; Remaining = 0; PercentOfBudget = 0 }
                }
            }
            
            # Add category and update phase totals
            $budgetHours = ConvertTo-SafeDouble $row["BudgetHours"]
            $actualHours = ConvertTo-SafeDouble $row["ActualHours"]
            $laborBudget = ConvertTo-SafeDouble $row["LaborBudget"]
            $laborCost = ConvertTo-SafeDouble $row["LaborCost"]
            $materialBudget = ConvertTo-SafeDouble $row["MaterialBudget"]
            $materialCost = ConvertTo-SafeDouble $row["MaterialCost"]
            $subcontractBudget = ConvertTo-SafeDouble $row["SubcontractBudget"]
            $subcontractCost = ConvertTo-SafeDouble $row["SubcontractCost"]
            $equipmentBudget = ConvertTo-SafeDouble $row["EquipmentBudget"]
            $equipmentCost = ConvertTo-SafeDouble $row["EquipmentCost"]
            $otherBudget = ConvertTo-SafeDouble $row["OtherBudget"]
            $otherCost = ConvertTo-SafeDouble $row["OtherCost"]
            $administrativeBudget = ConvertTo-SafeDouble $row["AdministrativeBudget"]
            $administrativeCost = ConvertTo-SafeDouble $row["AdministrativeCost"]
            
            $jobData[$jobNum].Phases[$phaseNum].Categories[$catNum] = @{
                CategoryName = $row["CatName"]
                Hours = New-BudgetActualObject $budgetHours $actualHours
                HoursCost = New-BudgetActualObject $laborBudget $laborCost
                MaterialCost = New-BudgetActualObject $materialBudget $materialCost
                SubcontractCost = New-BudgetActualObject $subcontractBudget $subcontractCost
                EquipmentCost = New-BudgetActualObject $equipmentBudget $equipmentCost
                OtherCost = New-BudgetActualObject $otherBudget $otherCost
                AdministrativeCost = New-BudgetActualObject $administrativeBudget $administrativeCost
            }
            
            # Update phase totals
            $phase = $jobData[$jobNum].Phases[$phaseNum]
            $phase.Hours.Budget += $budgetHours
            $phase.Hours.Actual += $actualHours
            $phase.HoursCost.Budget += $laborBudget
            $phase.HoursCost.Actual += $laborCost
            $phase.MaterialCost.Budget += $materialBudget
            $phase.MaterialCost.Actual += $materialCost
            $phase.SubcontractCost.Budget += $subcontractBudget
            $phase.SubcontractCost.Actual += $subcontractCost
            $phase.EquipmentCost.Budget += $equipmentBudget
            $phase.EquipmentCost.Actual += $equipmentCost
            $phase.OtherCost.Budget += $otherBudget
            $phase.OtherCost.Actual += $otherCost
            $phase.AdministrativeCost.Budget += $administrativeBudget
            $phase.AdministrativeCost.Actual += $administrativeCost
        }
        
        # Finalize phase calculations
        foreach ($job in $jobData.Values) {
            foreach ($phase in $job.Phases.Values) {
                $phase.Hours = New-BudgetActualObject $phase.Hours.Budget $phase.Hours.Actual
                $phase.HoursCost = New-BudgetActualObject $phase.HoursCost.Budget $phase.HoursCost.Actual
                $phase.MaterialCost = New-BudgetActualObject $phase.MaterialCost.Budget $phase.MaterialCost.Actual
                $phase.SubcontractCost = New-BudgetActualObject $phase.SubcontractCost.Budget $phase.SubcontractCost.Actual
                $phase.EquipmentCost = New-BudgetActualObject $phase.EquipmentCost.Budget $phase.EquipmentCost.Actual
                $phase.OtherCost = New-BudgetActualObject $phase.OtherCost.Budget $phase.OtherCost.Actual
                $phase.AdministrativeCost = New-BudgetActualObject $phase.AdministrativeCost.Budget $phase.AdministrativeCost.Actual
            }
        }

        # Filter out phases and categories with no data
        $totalCategoriesRemoved = 0
        $totalPhasesRemoved = 0
        
        foreach ($job in $jobData.Values) {
            $phasesToRemove = @()
            
            foreach ($phaseKey in $job.Phases.Keys) {
                $phase = $job.Phases[$phaseKey]
                
                # Filter out categories with no data
                $categoriesToRemove = @()
                foreach ($categoryKey in $phase.Categories.Keys) {
                    $category = $phase.Categories[$categoryKey]
                    if (-not (Test-CategoryHasData $category)) {
                        $categoriesToRemove += $categoryKey
                    }
                }
                
                # Remove empty categories
                foreach ($categoryKey in $categoriesToRemove) {
                    $phase.Categories.Remove($categoryKey)
                    $totalCategoriesRemoved++
                }
                
                # Check if phase has data after category filtering
                if ($phase.Categories.Count -eq 0 -or -not (Test-PhaseHasData $phase)) {
                    $phasesToRemove += $phaseKey
                }
            }
            
            # Remove empty phases
            foreach ($phaseKey in $phasesToRemove) {
                $job.Phases.Remove($phaseKey)
                $totalPhasesRemoved++
            }
        }
        
        Write-Host "Filtered out $totalCategoriesRemoved categories and $totalPhasesRemoved phases with no data"

        # Filter out jobs with no meaningful data
        $originalCount = $jobData.Count
        $filteredJobData = [ordered]@{}
        $emptyJobs = @()
        
        foreach ($jobKey in $jobData.Keys) {
            if (Test-JobHasData $jobData[$jobKey]) {
                $filteredJobData[$jobKey] = $jobData[$jobKey]
            } else {
                $emptyJobs += $jobKey
            }
        }
        
        $filteredCount = $filteredJobData.Count
        $removedCount = $originalCount - $filteredCount
        
        Write-Host "Total jobs found: $originalCount"
        Write-Host "Jobs with data: $filteredCount"
        Write-Host "Empty jobs filtered out: $removedCount"
        
        if ($emptyJobs.Count -gt 0) {
            Write-Host "Filtered out jobs: $($emptyJobs -join ', ')"
        }

        # Export filtered data to JSON
        if ($filteredJobData.Count -gt 0) {
            $filteredJobData | ConvertTo-Json -Depth 15 | Out-File $OutputFile -Encoding utf8
        Write-Host "Data exported to $OutputFile"
        } else {
            Write-Warning "No jobs with meaningful data found. No file created."
        }
        
        return $filteredJobData
    }
    catch { Write-Error "Error: $_" }
    finally { if ($conn.State -eq 'Open') { $conn.Close() } }
}

