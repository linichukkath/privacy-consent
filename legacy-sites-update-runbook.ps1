param(
    [Parameter(Mandatory=$false)]
    [object]$WebhookData,

    [Parameter(Mandatory=$false)]
    [string]$siteUrl = 'https://allianzms.sharepoint.com/teams/DE1890-2624280',

    [Parameter(Mandatory=$false)]
    [string]$listItemId = 6664
)

# Import SharePoint module FIRST to avoid DLL version conflicts
# Azure Automation may pre-load it; skip if already loaded to avoid duplicate exports
if (-not (Get-Module Microsoft.Online.SharePoint.PowerShell)) {
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
}

# Import custom module (which loads additional DLLs)
# Client secrets are read from Azure Automation Credential assets or Variables.
# In Azure Automation the module is installed by name; locally fall back to the manifest beside this script.
$localModuleManifest = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'legacy-sites-update-runbook.psd1' } else { $null }
if (Get-Module -ListAvailable -Name 'legacy-sites-update-runbook') {
    Import-Module 'legacy-sites-update-runbook' -Force -ErrorAction Stop
}
elseif ($localModuleManifest -and (Test-Path $localModuleManifest)) {
    Import-Module $localModuleManifest -Force -ErrorAction Stop
}
else {
    throw "Module 'legacy-sites-update-runbook' not found in PSModulePath or beside the runbook script."
}

Function Resolve-TenantAuthConfigBySiteUrl {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )

    try {
        $siteHost = ([System.Uri]$SiteUrl).Host.ToLowerInvariant()

        # Keep auth settings local to this runbook (no dependency on auth.ps1 resolver).
        $tenantMap = @{
            Test = @{
                Tenant   = "allianzmstest"
                TenantId = "0793811a-e290-4ab9-bc33-296559e28bea"
                ClientId = "84b5c5bc-7654-4dbe-b3a2-a12caa3932f5"
                SecretName = "allianzmstest-client-secret"
                CertificateName = "pfx-test-tenant"
            }
            Production = @{
                Tenant   = "allianzms"
                TenantId = "6e06e42d-6925-47c6-b9e7-9581c7ca302a"
                ClientId = "d4dfa3f9-88ac-4d83-af72-cc72ff128699"
                SecretName = "client-secret"
                CertificateName = "pfx"
            }
        }

        if ($siteHost -match '^allianzmstest[a-z0-9-]*\.sharepoint\.com$') {
            Write-LogINFO "Test tenant detected from URL host: $siteHost"
            return [PSCustomObject]@{
                Environment = "Test"
                Tenant      = $tenantMap.Test.Tenant
                TenantId    = $tenantMap.Test.TenantId
                ClientId    = $tenantMap.Test.ClientId
                SecretName = $tenantMap.Test.SecretName
                CertificateName = $tenantMap.Test.CertificateName
            }
        }

        if ($siteHost -match '^allianzms(?!test)[a-z0-9-]*\.sharepoint\.com$') {
            Write-LogINFO "Production tenant detected from URL host: $siteHost"
            return [PSCustomObject]@{
                Environment = "Production"
                Tenant      = $tenantMap.Production.Tenant
                TenantId    = $tenantMap.Production.TenantId
                ClientId    = $tenantMap.Production.ClientId
                SecretName = $tenantMap.Production.SecretName
                CertificateName = $tenantMap.Production.CertificateName
            }
        }

        throw "Unable to map site URL host '$siteHost' to a supported tenant. Expected allianzmstest*.sharepoint.com or allianzms*.sharepoint.com"
    }
    catch {
        Write-LogERR "Tenant resolution failed for URL '$SiteUrl': $($_.Exception.Message)"
        throw
    }
}

$isUpdateItemsBatch = $false
$batchLibraryGuid = $null
$batchLibraryTitle = $null
$batchStartId = 0
$batchEndId = 0

# Handle webhook data vs manual execution
if ($WebhookData) {
    Write-LogINFO "Webhook RequestBody: $($WebhookData.RequestBody)"

    if ($WebhookData.RequestBody) {
        try {
            $requestData = $WebhookData.RequestBody | ConvertFrom-Json
            $siteUrl = $requestData.siteUrl
            $listItemId = $requestData.listItemId

            if ($requestData.operation -eq 'UpdateItemsBatch') {
                $isUpdateItemsBatch = $true
                $batchLibraryGuid = [string]$requestData.libraryGUID
                $batchLibraryTitle = [string]$requestData.libraryTitle
                $batchStartId = [int]$requestData.startId
                $batchEndId = [int]$requestData.endId
                Write-LogINFO "UpdateItems webhook worker received: library '$batchLibraryTitle', IDs $batchStartId-$batchEndId"
            }

            # Handle case where siteUrl is a JSON object string e.g. {"Url":"https://...","Description":"..."}
            try {
                $siteUrlObj = $siteUrl | ConvertFrom-Json -ErrorAction Stop
                if ($siteUrlObj.Url) {
                    Write-LogINFO "siteUrl is a JSON object - extracting Url property"
                    $siteUrl = $siteUrlObj.Url
                }
            }
            catch {
                # siteUrl is a plain string, use as-is
            }
        }
        catch {
            Write-LogERR "Failed to parse webhook data: $_"
            throw "Invalid webhook data format. Expected JSON with 'siteUrl' property."
        }
    }
    else {
        Write-LogERR "WebhookData.RequestBody is null or empty"
        throw "Webhook data is missing RequestBody"
    }
}
elseif ($siteUrl) {
    # Manual execution with siteUrl parameter
}
else {
    Write-LogERR "No webhook data or siteUrl parameter provided"
    throw "Runbook requires either WebhookData or siteUrl parameter"
}

# Validate siteUrl
if ([string]::IsNullOrEmpty($siteUrl)) {
    Write-LogERR "Site URL is null or empty"
    throw "Site URL is required"
}

# Get redirected URL if needed
$siteUrl = GetRedirectedSiteUrl($siteUrl)

# Resolve tenant/auth configuration from URL
$tenantAuthConfig = Resolve-TenantAuthConfigBySiteUrl -SiteUrl $siteUrl
$tenant = $tenantAuthConfig.Tenant
$tenantId = $tenantAuthConfig.TenantId
$clientId = $tenantAuthConfig.ClientId
Write-LogINFO "Environment: $($tenantAuthConfig.Environment)"

# Retrieve client secret from an Azure Automation Credential asset or Variable.
# Certificate authentication remains preferred when the certificate is available.
$secretName = 'client-secret'
$certificateName = $tenantAuthConfig.CertificateName
$clientSecret = $null

if ([string]::IsNullOrWhiteSpace($clientSecret) -and (Get-Command Get-AutomationVariable -ErrorAction SilentlyContinue)) {
    try {
        $clientSecret = Get-AutomationVariable -Name $secretName -ErrorAction Stop
        Write-LogINFO "Client secret retrieved from Automation Variable '$secretName'."
    }
    catch {
        Write-LogWARN "Automation Variable '$secretName' was not available: $($_.Exception.Message)"
    }
}

Write-LogINFO "=== Using App-Only Authentication ==="
Write-LogINFO "Client ID: $clientId"
Write-LogINFO "Tenant: $tenant"
Write-LogINFO "Tenant ID: $tenantId"

$certificate = $null
# Get-AutomationCertificate only exists inside Azure Automation; its absence means a local run.
$isLocalRun = -not (Get-Command Get-AutomationCertificate -ErrorAction SilentlyContinue)

if ($isLocalRun) {
    # Step 1 (local): load the .pfx from the certs folder referenced by configs\localhost-config[-test].json
    Write-LogINFO "Local run detected. Loading certificate from local config..."
    $localConfigFile = Get-ConfigFileByUrl -SiteUrl $siteUrl
    $localConfig = Get-Content $localConfigFile -Raw | ConvertFrom-Json

    if ($localConfig.ClientId -and $localConfig.ClientId -ne $clientId) {
        throw "ClientId in '$localConfigFile' ($($localConfig.ClientId)) does not match the tenant app ($clientId). Fix the local config before running."
    }

    if ([string]::IsNullOrWhiteSpace($localConfig.CertificatePath)) {
        Write-LogWarn "No CertificatePath in '$localConfigFile'."
    }
    else {
        $certPath = if ([System.IO.Path]::IsPathRooted($localConfig.CertificatePath)) {
            $localConfig.CertificatePath
        } else {
            Join-Path $global:CurrDir $localConfig.CertificatePath
        }
        if (-not (Test-Path $certPath)) {
            throw "Local certificate not found: $certPath"
        }
        try {
            $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                $certPath,
                [string]$localConfig.CertificatePassword,
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
            )
        }
        catch {
            throw "Failed to load local certificate '$certPath': $($_.Exception.Message)"
        }
        if (-not $certificate.HasPrivateKey) {
            throw "Local certificate '$certPath' does not contain a private key."
        }
        Write-LogINFO "Certificate loaded from $certPath"
        Write-LogINFO "Thumbprint: $($certificate.Thumbprint)"
        Write-LogINFO "Subject: $($certificate.Subject)"
    }

    if ([string]::IsNullOrWhiteSpace($clientSecret) -and -not [string]::IsNullOrWhiteSpace($localConfig.ClientSecret)) {
        $clientSecret = $localConfig.ClientSecret
    }

    if ($null -eq $certificate -and [string]::IsNullOrWhiteSpace($clientSecret)) {
        throw "No authentication material is available locally. Set CertificatePath/CertificatePassword (or ClientSecret) in '$localConfigFile'."
    }
}
else {
    # Step 1: Get certificate from Azure Automation
    Write-LogINFO "Retrieving certificate from Azure Automation..."
    try {
        $certificate = Get-AutomationCertificate -Name $certificateName -ErrorAction Stop
        Write-LogINFO "Certificate retrieved successfully"
        Write-LogINFO "Thumbprint: $($certificate.Thumbprint)"
        Write-LogINFO "Subject: $($certificate.Subject)"
        Write-LogINFO "HasPrivateKey: $($certificate.HasPrivateKey)"
    }
    catch {
        Write-LogWarn "Certificate not available from Azure Automation: $_"
        Write-LogWarn "Will fall back to client secret authentication"
    }

    if ($null -eq $certificate -and [string]::IsNullOrWhiteSpace($clientSecret)) {
        throw "No authentication material is available. Configure certificate '$certificateName' or client secret Credential/Variable '$secretName' in the Automation Account."
    }
}

# Initialize migration tracker (no-op if listItemId is null - manual/test runs)
Initialize-MigrationTracker `
    -ListItemId $listItemId `
    -ClientId $clientId `
    -ClientSecret $clientSecret `
    -TenantId $tenantId `
    -Certificate $certificate

# Step 2: Create SharePoint CSOM context
Write-LogINFO "Creating SharePoint CSOM context for: $siteUrl"
try {
    $context = GetSPOContext `
        -siteUrl $siteUrl `
        -clientId $clientId `
        -clientSecret $clientSecret `
        -tenantId $tenantId `
        -certificate $certificate

    if ($null -eq $context) {
        throw "Failed to create SharePoint context"
    }

    # Verify connection by loading Web
    Write-LogINFO "Verifying authentication by loading Web properties..."
    $context.Load($context.Web)
    ExecuteQueryWithRetry -Ctx $context
    Write-LogINFO "[OK] Authentication succeeded - Web title: $($context.Web.Title)"
}
catch {
    Write-LogERR "Failed to create SharePoint context: $_"
    throw "Authentication failed for site: $siteUrl"
}

if ($isUpdateItemsBatch) {
    if ([string]::IsNullOrWhiteSpace($batchLibraryGuid) -or $batchStartId -le 0 -or $batchEndId -lt $batchStartId) {
        throw "Invalid UpdateItemsBatch webhook payload. libraryGUID, startId, and endId are required."
    }

    $batchCompanyId = Get-CompanyId -context $context -siteUrl $siteUrl
    $batchIsDms = Get-IsDMS -context $context
    $batchLanguage = Get-Language -context $context
    $batchUpdateConfigPath = if ($batchIsDms) { "$global:CurrDir\configs\legacy-to-e5-update.json" } else { "$global:CurrDir\configs\legacy-to-e3-update.json" }
    $batchUpdateConfig = Get-Content $batchUpdateConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $batchContentTypes = Get-Content "$global:CurrDir\configs\content-types.json" -Raw -Encoding UTF8 | ConvertFrom-Json
    $batchTerms = getDocumentClassTerms -context $context
    $batchTimeZoneOffset = getTimezoneOffset -ctx $context

    Write-LogINFO "Processing UpdateItems webhook batch for '$batchLibraryTitle'."
    $batchResult = updateItems `
        -siteUrl $siteUrl `
        -libraryGUID $batchLibraryGuid `
        -libraryTitle $batchLibraryTitle `
        -contentTypesConfig $batchContentTypes `
        -documentClassTerms $batchTerms `
        -updateConfig $batchUpdateConfig `
        -language $batchLanguage `
        -timeZoneOffset $batchTimeZoneOffset `
        -context $context `
        -StartId $batchStartId `
        -EndId $batchEndId

    # Report counts on the OUTPUT stream so the dispatching master job can read them back through
    # ARM and log real totals. Emitted before any throw, so a batch that fails still reports how
    # much it managed to update. Write-Log* uses the Warning stream, so Output carries only this.
    $batchUpdatedCount = @($batchResult.ItemChanges).Count
    $batchFailedCount = @($batchResult.Exceptions).Count
    $batchExceptionsCount = @($batchResult.Exceptions).Count

    # Each child writes its own reports for its own slice. The per-item detail lives here and
    # nowhere else - shipping millions of rows back to the master through a job output stream is
    # not possible, so the master collects URLs instead of records. The batch number keeps
    # concurrent children from overwriting one another in SiteAssets.
    $batchReportSuffix = "-batch$($requestData.batchNumber)"
    $batchItemReportUrl = $null
    $batchCtReportUrl = $null
    $batchExceptionsReportUrl = $null
    try {
        if ($batchUpdatedCount -gt 0) {
            $batchItemReportUrl = NewItemMigrationReport `
                -itemChanges @($batchResult.ItemChanges) `
                -mainSiteUrl $siteUrl `
                -context $context `
                -fileNameSuffix $batchReportSuffix
            if ($batchItemReportUrl) { Write-LogINFO "Item migration report for this batch: $batchItemReportUrl" }

            if ($batchExceptionsCount -gt 0) {
                $batchExceptionsReportUrl = NewExceptionsReport `
                    -exceptions @($batchResult.Exceptions) `
                    -mainSiteUrl $siteUrl `
                    -context $context `
                    -fileNameSuffix $batchReportSuffix
                if ($batchExceptionsReportUrl) { Write-LogINFO "Item exception report for this batch: $batchExceptionsReportUrl" }
            }

            if ($batchResult.BeforeCtCounts -and $batchResult.BeforeCtCounts.Count -gt 0) {
                $batchCtReportUrl = NewContentTypeSummaryReport `
                    -beforeCtCounts $batchResult.BeforeCtCounts `
                    -itemChanges @($batchResult.ItemChanges) `
                    -mainSiteUrl $siteUrl `
                    -context $context `
                    -fileNameSuffix $batchReportSuffix
                if ($batchCtReportUrl) { Write-LogINFO "Content type summary report for this batch: $batchCtReportUrl" }
            }
        }
    }
    catch {
        # Reporting must never lose a successful migration - the items are already updated.
        Write-LogWarn "Could not generate batch reports: $($_.Exception.Message)"
    }

    $batchSummaryJson = [PSCustomObject]@{
        batchNumber   = $requestData.batchNumber
        libraryTitle  = $batchLibraryTitle
        startId       = $batchStartId
        endId         = $batchEndId
        updated       = $batchUpdatedCount
        failed        = $batchFailedCount
        itemReportUrl = $batchItemReportUrl
        ctReportUrl   = $batchCtReportUrl
    } | ConvertTo-Json -Depth 3 -Compress

    # Write the result where the parent can fetch it by an exact, computable name. This is the
    # primary channel: reading a child's Output stream back through ARM proved unreliable, and a
    # parent that cannot read results reports "0 items updated" for work that actually happened.
    try {
        $batchResultFileName = Get-UpdateItemsBatchResultFileName `
            -LibraryGUID $batchLibraryGuid -BatchNumber ([int]$requestData.batchNumber)
        $batchResultLocalPath = Join-Path (Get-RunbookTempDir) $batchResultFileName
        $batchSummaryJson | Out-File -FilePath $batchResultLocalPath -Encoding UTF8

        $batchResultToken = Get-AzureADAccessToken `
            -tenantId $context.AuthState.TenantId `
            -clientId $context.AuthState.ClientId `
            -clientSecret $context.AuthState.ClientSecret `
            -certificate $context.AuthState.Certificate `
            -resourceUrl $siteUrl
        $batchResultHeaders = @{
            Authorization = "Bearer $batchResultToken"
            Accept        = "application/json;odata=verbose"
            ContentType   = "application/json;odata=verbose"
        }
        $batchResultUploaded = uploadFileToSiteAssets `
            -siteUrl $siteUrl `
            -headers $batchResultHeaders `
            -sourceFilePath $batchResultLocalPath `
            -targetFileName $batchResultFileName `
            -context $context
        if ($batchResultUploaded) {
            Write-LogINFO "Batch result published: $batchResultFileName"
        }
        else {
            Write-LogWarn "Could not publish batch result file '$batchResultFileName'; the master job will fall back to the job output stream."
        }
    }
    catch {
        Write-LogWarn "Could not publish batch result file: $($_.Exception.Message)"
    }

    # Secondary channel, kept as a fallback.
    Write-Output "DMS-BATCH-RESULT $batchSummaryJson"
    Write-LogINFO "Batch summary: $batchUpdatedCount item(s) updated, $batchFailedCount failed (IDs $batchStartId-$batchEndId)."

    if ($batchFailedCount -gt 0) {
        Write-LogERR "UpdateItems webhook batch completed with $batchFailedCount item failure(s)."
        throw "UpdateItems webhook batch failed for library '$batchLibraryTitle' ($batchStartId-$batchEndId)."
    }
    Write-LogINFO "UpdateItems webhook batch completed successfully."
    return
}

# Step 3: Connect to tenant for SPO Management Shell operations
Write-LogINFO "Connecting to tenant for SPO Management Shell operations..."
$tenantConnected = $false
try {
    LogInToTenant `
        -tenant $tenant `
        -clientId $clientId `
        -tenantId $tenantId `
        -certificate $certificate `
        -clientSecret $clientSecret

    $tenantConnected = $true
    Write-LogINFO "[OK] Successfully connected to tenant - SPO cmdlets available"

    # Verify connection by getting site information
    try {
        $spoSite = Get-SPOSite -Identity $siteUrl -ErrorAction Stop
        Write-LogINFO "Site information retrieved:"
        Write-LogINFO "  Title: $($spoSite.Title)"
        Write-LogINFO "  URL: $($spoSite.Url)"
        Write-LogINFO "  Storage Quota: $($spoSite.StorageQuota) MB"
        Write-LogINFO "  Owner: $($spoSite.Owner)"
        Write-LogINFO "  Lock State: $($spoSite.LockState)"
    }
    catch {
        Write-LogWarn "Tenant connection test failed: $_"
        $tenantConnected = $false
    }
}
catch {
    Write-LogWarn "Failed to connect to tenant: $_"
    Write-LogWarn "SPO Management Shell cmdlets will not be available"
    Write-LogWarn "Script will continue using CSOM operations only"
}

# Get parameters dynamically from property bag
$companyId = Get-CompanyId -context $context -siteUrl $siteUrl
$isDMS = Get-IsDMS -context $context
$language = Get-Language -context $context

$global:RunbookJobId = if ($PSPrivateMetadata.JobId) { $PSPrivateMetadata.JobId.ToString() } else { "local-$(Get-Date -Format 'yyyyMMddHHmmss')" }

function Invoke-RunbookLogUpload {
    param($Context)
    try {
        $logFile = Save-RunbookLog -JobId $global:RunbookJobId
        if (-not $logFile) { return }
        $trackerSiteUrl = Get-MigrationTrackerSiteUrl
        $logAccessToken = Get-AzureADAccessToken `
            -tenantId $Context.AuthState.TenantId `
            -clientId $Context.AuthState.ClientId `
            -clientSecret $Context.AuthState.ClientSecret `
            -certificate $Context.AuthState.Certificate `
            -resourceUrl $trackerSiteUrl
        $logHeaders = @{
            Authorization = "Bearer $logAccessToken"
            Accept        = "application/json;odata=verbose"
            ContentType   = "application/json;odata=verbose"
        }
        $uploaded = uploadFileToSiteAssets `
            -siteUrl $trackerSiteUrl `
            -headers $logHeaders `
            -sourceFilePath $logFile.FilePath `
            -targetFileName $logFile.FileName `
            -context $Context
        if ($uploaded) {
            Add-MigrationTrackerError -Message "Full job log: $trackerSiteUrl/SiteAssets/$($logFile.FileName)"
        }
    }
    catch {
        Write-Warning "Log upload failed: $($_.Exception.Message)"
    }
}

Write-LogINFO '=== The legacy update script has started ==='
Write-LogINFO "Site URL: $siteUrl"
Write-LogINFO "Company ID: $companyId"
Write-LogINFO "Is DMS: $isDMS"
Write-LogINFO "Language: $language"

try {
    Set-MigrationTrackerStatus -Status "QC running"

    # Callback for QC passed -> Migrating transition (invoked from startLegacyUpdate)
    $global:OnMigrationPhaseChange = {
        param([string]$Phase)
        if ($Phase -eq "QCPassed") {
            Set-MigrationTrackerQCResult -Passed $true -Findings ""
            Set-MigrationTrackerStatus -Status "QC passed"
            Set-MigrationTrackerStatus -Status "Migrating"
        }
    }

    $result = startLegacyUpdate `
        -url $siteUrl `
        -context $context `
        -isSourceDMSUpgrade $false `
        -companyId $companyId `
        -libraryUrl $null `
        -tenant $tenant `
        -isDMSUPS $isDMS `
        -languageUPS $language `
        -tenantConnected $tenantConnected

    if ($result) {
        Invoke-RunbookLogUpload -Context $context
        Complete-MigrationTracker -Success $true
    }
    else {
        Add-MigrationTrackerError -Message "Migration or integration test failed for $siteUrl"
        Invoke-RunbookLogUpload -Context $context
        Complete-MigrationTracker -Success $false
    }
}
catch {
    $errorMsg = $_.Exception.Message

    # PowerShell wraps thrown [System.Exception] in RuntimeException, so Data lives on InnerException
    $exceptionData = if ($_.Exception.Data["ErrorType"]) { $_.Exception.Data } `
                     elseif ($_.Exception.InnerException) { $_.Exception.InnerException.Data } `
                     else { $null }

    if ($exceptionData -and $exceptionData["ErrorType"] -eq "QCFailed") {
        # QC failure - checked-out files and/or compliance-tagged files found
        $reportUrl = $exceptionData["ConsolidatedReportUrl"]
        $sitesWithCheckedOut = $exceptionData["SitesWithCheckedOutFiles"]
        $sitesWithComplianceTags = $exceptionData["SitesWithComplianceTaggedFiles"]

        $findingsParts = @()
        if ($sitesWithCheckedOut -gt 0) {
            $findingsParts += "Checked-out files detected in $sitesWithCheckedOut site(s)"
        }
        if ($sitesWithComplianceTags -gt 0) {
            $findingsParts += "Compliance-tagged files detected in $sitesWithComplianceTags site(s)"
        }
        $findings = ($findingsParts -join "; ") + ". Report: $reportUrl"

        $fixInstructions = "The migration could not proceed due to QC failures.`n`n"
        if ($sitesWithCheckedOut -gt 0) {
            $fixInstructions += "Checked-out files:`n1. Review the QC report: $reportUrl`n2. Ask file owners to check in or discard their checkouts.`n`n"
        }
        if ($sitesWithComplianceTags -gt 0) {
            $fixInstructions += "Compliance-tagged files:`n1. Review the QC report: $reportUrl`n2. Contact the compliance/records team to assess whether retention labels must be removed before migration.`n`n"
        }
        $fixInstructions += "Once all issues are resolved, re-submit the migration request."

        Set-MigrationTrackerQCResult -Passed $false -Findings $findings
        Set-MigrationTrackerFixInstructions -Instructions $fixInstructions
        Add-MigrationTrackerError -Message $errorMsg
        Invoke-RunbookLogUpload -Context $context
        Complete-MigrationTracker -Success $false -FailureStatus "QC failed"
    }
    else {
        # General migration failure
        Add-MigrationTrackerError -Message $errorMsg
        Invoke-RunbookLogUpload -Context $context
        Complete-MigrationTracker -Success $false
    }

    throw   # Re-throw so Azure Automation marks the job as failed
}
finally {
    $global:OnMigrationPhaseChange = $null
    if ($tenantConnected) {
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
    }
}