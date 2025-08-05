

# Extract configuration values from environment variables
$tenantId = $env:KeystoneTenantID
$clientId = $env:ReportsClientID
$clientSecret = $env:ReportsClientSecret

# Validate that environment variables are set
if (-not $tenantId) { throw "Environment variable 'KeystoneTenantID' is not set" }
if (-not $clientId) { throw "Environment variable 'ReportsClientID' is not set" }
if (-not $clientSecret) { throw "Environment variable 'ReportsClientSecret' is not set" }



function Get-AccessToken {
    try {
        # Get access token
        $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        $tokenBody = @{
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }
        
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
        return $tokenResponse.access_token
    }
    catch {
        Write-Error "Failed to get access token: $_"
        return $null
    }
}

function Test-UserExists {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Email,
        
        [Parameter(Mandatory=$true)]
        [string]$AccessToken
    )
    
    try {
        # Set up headers for Graph API requests
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "ConsistencyLevel" = "eventual"
        }
        $escapedEmail = [System.Web.HttpUtility]::UrlEncode($Email)
        $userUrl = "https://graph.microsoft.com/v1.0/users?`$filter=(mail eq '$escapedEmail' or userPrincipalName eq '$escapedEmail' or proxyAddresses/any(x: x eq 'SMTP:$escapedEmail'))&`$count=true"
        
        # Make the API call
        $response = Invoke-RestMethod -Uri $userUrl -Headers $headers -Method Get
        
        # Check if any users match the query
        return ($response.value.Count -gt 0)
    }
    catch {
        Write-Warning "Error checking if user exists: $_"
        return $false
    }
}

function Send-EmailWithAttachment {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SenderEmail,

        [Parameter(Mandatory=$true)]
        [string[]]$To,
        
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [string]$Body,
        
        [Parameter(Mandatory=$false)]
        [string[]]$AttachmentPaths,
        
        [Parameter(Mandatory=$false)]
        [string]$Cc,
        
        [Parameter(Mandatory=$true)]
        [string]$AccessToken
    )
    
    $maxRetries = 3
    $retryDelaySeconds = 30
    $attempt = 0

    while ($attempt -lt $maxRetries) {
        $attempt++
        try {
            # Set up headers for Graph API requests
            $headers = @{
                "Authorization" = "Bearer $AccessToken"
                "Content-Type"  = "application/json"
            }
            
                         # Prepare email message
             $emailMessage = @{
                 message = @{
                     subject      = $Subject
                     body         = @{
                         contentType = "HTML"
                         content     = $Body
                     }
                     toRecipients = @()
                 }
             }
            
            # Add all recipients to the toRecipients array
            foreach ($recipient in $To) {
                $emailMessage.message.toRecipients += @(
                    @{
                        emailAddress = @{
                            address = $recipient
                        }
                    }
                )
            }
            
            # Add CC recipients if provided
            if ($Cc) {
                $emailMessage.message.ccRecipients = @(
                    @{
                        emailAddress = @{
                            address = $Cc
                        }
                    }
                )
            }
            
            # Process attachments if provided
            if ($AttachmentPaths -and $AttachmentPaths.Count -gt 0) {
                $attachments = @()
                
                foreach ($filePath in $AttachmentPaths) {
                    if (Test-Path -Path $filePath) {
                        $fileName = Split-Path -Path $filePath -Leaf
                        $fileContent = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($filePath))
                        
                        $attachments += @{
                            "@odata.type" = "#microsoft.graph.fileAttachment"
                            name          = $fileName
                            contentType   = "application/octet-stream"
                            contentBytes  = $fileContent
                        }
                    }
                    else {
                        Write-Warning "Attachment file not found: $filePath"
                    }
                }
                
                if ($attachments.Count -gt 0) {
                    $emailMessage.message.attachments = $attachments
                }
            }
            
            # Send the email
            $sendMailUrl = "https://graph.microsoft.com/v1.0/users/$SenderEmail/sendMail"
            Invoke-RestMethod -Uri $sendMailUrl -Headers $headers -Method Post -Body ($emailMessage | ConvertTo-Json -Depth 4)
            
            Write-Output "Email sent successfully to $($To -join ', ')"
            return $true
        }
        catch {
            $errorMessage = $_.ToString()
            # Remove escape characters from error message
            $errorMessage = $errorMessage -replace '\\u0027', "'" -replace '\\u0022', '"' -replace '\\n', "`n" -replace '\\r', "`r" -replace '\\t', "`t"
            
            if ($errorMessage -like "*MailboxInfoStale*") {
                if ($attempt -lt $maxRetries) {
                    Write-Warning "Attempt $attempt of $maxRetries failed for $($To -join ', '): MailboxInfoStale. Retrying in $retryDelaySeconds seconds..."
                    Start-Sleep -Seconds $retryDelaySeconds
                } else {
                    Write-Error "Failed to send email to $($To -join ', ') after $maxRetries attempts. Error: $errorMessage"
                    return $false
                }
            } else {
                # For other errors, fail immediately
                Write-Error "Failed to send email: $errorMessage"
                return $false
            }
        }
    }
}

function Send-Reports {
    param (
        [string]$Group = "",
        [DateTime]$CutoffDate = [DateTime]::Today
    )

    $dateForFiles = $CutoffDate.ToString("yyyyMMdd")
    $dateForHuman = $CutoffDate.ToString("MM/dd/yyyy")

    $configPath = "config\config.json"
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

    if ($Group -eq "") {
        foreach ($group in $config.PSObject.Properties.Name) {
            Send-Reports -Group $group -CutoffDate $CutoffDate
        }
        return
    }

    $senderEmail = $config.$Group.SenderEmail
    $summaryRecipients = $config.$Group.SummaryRecipients
    $errorRecipients = $config.$Group.ErrorRecipients

    $accessToken = Get-AccessToken

    if (-not $accessToken) {
        Write-Error "Failed to get access token"
        return
    }

    #Send Summary Report
    $reports_root_path = "reports\${Group}\${dateForFiles}"
    $summary_workbook_path = "${reports_root_path}\${Group}_Manhour_Summary_${dateForFiles}.xlsx"
    write-host "Summary workbook path: $summary_workbook_path"

    # Read the HTML template for summary reports
    $summaryTemplatePath = Join-Path $PSScriptRoot "templates\SummaryJobReport.html"
    $summaryHtmlTemplate = Get-Content -Path $summaryTemplatePath -Raw -Encoding UTF8
    
    # Replace placeholders in the template
    $summaryBody = $summaryHtmlTemplate -replace "{{GROUP}}", $Group -replace "{{GENERATION_DATE}}", (Get-Date).ToString("MM/dd/yyyy")

    
    Send-EmailWithAttachment -SenderEmail $senderEmail -To $summaryRecipients -Subject "${Group} Manhour Summary Report - $dateForHuman" -Body $summaryBody -AttachmentPaths $summary_workbook_path -Cc $Cc -AccessToken $accessToken
    

         #Send individual reports to each user
     $folders = Get-ChildItem -Path $reports_root_path -Directory
     $user_does_not_exist = @()
     $error_files = @()
     $invalid_pm_jobs = @{}  # Hash table to store PM -> jobs mapping
 
     foreach ($folder in $folders) {
         $folder_name = $folder.Name
 
         # replcae space with underscore
         $email_recipient = $folder_name.Replace(" ", ".") + "@keystoneconcrete.com"
         $email_recipient = $email_recipient.Replace("'", "")

         write-host "Email recipient: $email_recipient"
 
         # Get files from the folder regardless of user existence
         $all_files_in_folder = Get-ChildItem -Path "${reports_root_path}\${folder_name}" -File
 
            if (Test-UserExists -Email $email_recipient -AccessToken $accessToken) {
              # Read the HTML template for individual reports
              $individualTemplatePath = Join-Path $PSScriptRoot "templates\IndividualJobReports.html"
              $individualHtmlTemplate = Get-Content -Path $individualTemplatePath -Raw -Encoding UTF8
              
              # Get unique job names for active projects count
              $unique_jobs = @()
              foreach ($file in $all_files_in_folder) {
                  $job_name = $file.Name.Split('_')[0]
                  if ($job_name -and -not $unique_jobs.Contains($job_name)) {
                      $unique_jobs += $job_name
                  }
              }
              
              # Replace placeholders in the template
              $body = $individualHtmlTemplate -replace "{{TOTAL_REPORTS}}", $all_files_in_folder.Count
              $body = $body -replace "{{ACTIVE_PROJECTS}}", $unique_jobs.Count
              $body = $body -replace "{{REPORT_DATE}}", $dateForHuman
              $body = $body -replace "{{DATE_FOR_HUMAN}}", $dateForHuman
              $body = $body -replace "{{GENERATION_DATE}}", (Get-Date).ToString("MM/dd/yyyy")
              
              $subject = "Job Manhour Reports - $dateForHuman"
              Send-EmailWithAttachment -SenderEmail $senderEmail -To $email_recipient -Subject $subject -Body $body -AttachmentPaths $all_files_in_folder -AccessToken $accessToken
 
         } else {
             $user_does_not_exist += $email_recipient
             $error_files += $all_files_in_folder
             
             # Get job names from the files in this folder
             $job_names = @()
             foreach ($file in $all_files_in_folder) {
                 # Extract job name from filename (assuming format like "JOB123_20250725.xlsx")
                 $job_name = $file.Name.Split('_')[0]
                 if ($job_name -and -not $job_names.Contains($job_name)) {
                     $job_names += $job_name
                 }
             }
             $invalid_pm_jobs[$email_recipient] = $job_names
         }
     }

    # Send error email to error recipients
    if ($user_does_not_exist.Count -gt 0) {
        # Read the HTML template
        $templatePath = Join-Path $PSScriptRoot "templates\InvalidPMNamesError.html"
        $htmlTemplate = Get-Content -Path $templatePath -Raw -Encoding UTF8
        
                 # Generate the list of invalid emails as HTML list items with job information
         $invalidEmailsHtml = ""
         foreach ($email in $user_does_not_exist) {
             # Extract display name from email (remove domain and replace dots with spaces)
             $displayName = $email.Split('@')[0] -replace '\.', ' '
             $job_names = $invalid_pm_jobs[$email]
             $job_count = $job_names.Count
             
                           $invalidEmailsHtml += "            <li class='pm-item'>`n"
              $invalidEmailsHtml += "                <div class='pm-name'><strong>$displayName</strong> ($email)</div>`n"
              
              if ($job_count -gt 0) {
                  $jobs_comma_separated = $job_names -join ', '
                  $invalidEmailsHtml += "                <div class='job-list'>$jobs_comma_separated</div>`n"
              }
              
              $invalidEmailsHtml += "            </li>`n"
         }
        
        # Replace the placeholder with the actual invalid emails
        $error_body = $htmlTemplate -replace "{{INVALID_EMAILS_PLACEHOLDER}}", $invalidEmailsHtml.TrimEnd()
        $error_body = $error_body -replace "{{GENERATION_DATE}}", (Get-Date).ToString("MM/dd/yyyy")


        $error_subject = "Invalid PM Names in Job Reports - $dateForHuman"

        Send-EmailWithAttachment -SenderEmail $senderEmail -To $errorRecipients -Subject $error_subject -Body $error_body -AttachmentPaths $error_files -AccessToken $accessToken
    }



}

