# Input bindings are passed in via param block.
param($Timer)

function Write-Blob {
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [datetime]$dateTime,
        [Parameter(Mandatory = $true, Position = 1)]
        [psobject]$logdata,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$contentType
    )

    if ($dateTime.kind.tostring() -ne 'Utc'){
        $dateTime = $dateTime.ToUniversalTime()
    }
    # Add DateTime to hashtable
    #$logdata.add("DateTime", $dateTime)
    $logdata | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $dateTime

    #Build the JSON file
    $logMessage = ConvertTo-Json $logdata -Depth 20
    Write-Verbose -Message $logMessage

    $fileName = $contentType + "-" + $dateTime.ToString("yyyyMMddHHmmss") + ".json"
    $blobName = $dateTime.ToString("yyyy/MM/dd") + "/" + $fileName
    $tempFilePath = "$env:TEMP\$fileName"
    $logMessage | Out-File $tempFilePath

    $containerName = "o365logs"
    if((Get-AzStorageContainer -Context $Context).Name -notcontains $containerName) {
        New-AzStorageContainer -Name $containerName -Context $Context
    }

    Set-AzStorageBlobContent -Container $containerName -Blob $blobName -Context $Context -File $tempFilePath -Force
    Remove-Item -Path $tempFilePath -Force
}

function Get-AuthToken{
    [cmdletbinding()]
        Param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$ClientID,
            [parameter(Mandatory = $true, Position = 1)]
            [string]$ClientSecret,
            [Parameter(Mandatory = $true, Position = 2)]
            [string]$tenantdomain,
            [Parameter(Mandatory = $true, Position = 3)]
            [string]$TenantGUID
        )
    # Create app of type Web app / API in Azure AD, generate a Client Secret, and update the client id and client secret here
    $loginURL = "https://login.microsoftonline.com/"
    # Get the tenant GUID from Properties | Directory ID under the Azure Active Directory section
    $resource = "https://manage.office.com"
    # auth
    $body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
    $oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    return $headerParams 
}

function Get-O365Data{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$startTime,
        [parameter(Mandatory = $true, Position = 1)]
        [string]$endTime,
        [Parameter(Mandatory = $true, Position = 2)]
        [psobject]$headerParams,
        [parameter(Mandatory = $true, Position = 3)]
        [string]$tenantGuid
    )
    #List Available Content
    $contentTypes = $env:contentTypes.split(",")
    Write-Host "ContentTypes: $contentTypes"
    #Loop for each content Type like Audit.General
    foreach($contentType in $contentTypes){
        Write-Host "Invoke REST API for $contentType"
        $listAvailableContentUri = "https://manage.office.com/api/v1.0/$tenantGUID/activity/feed/subscriptions/content?contentType=$contentType&PublisherIdentifier=$env:publisher&startTime=$startTime&endTime=$endTime"
        do {
            #List Available Content
            $contentResult = Invoke-RestMethod -Method GET -Headers $headerParams -Uri $listAvailableContentUri
            Write-Host "ContentResult: $contentResult"
            #Loop for each Content
            foreach($obj in $contentResult){
                #Retrieve Content
                $data = Invoke-RestMethod -Method GET -Headers $headerParams -Uri ($obj.contentUri)
                $data.Count
                Write-Blob (Get-Date) $data $contentType
            }
            
            #Handles Pagination
            $nextPageResult = Invoke-WebRequest -Method GET -Headers $headerParams -Uri $listAvailableContentUri
            If(($nextPageResult.Headers.NextPageUrl) -ne $null){
                $nextPage = $true
                $listAvailableContentUri = $nextPageResult.Headers.NextPageUrl
            }
            Else{$nextPage = $false}
        } until ($nextPage -eq $false)
    }
}
# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

#add last run time to blob file to ensure no missed packages
$endTime = $currentUTCtime | Get-Date -Format yyyy-MM-ddTHH:mm:ss
$azstoragestring = $Env:WEBSITE_CONTENTAZUREFILECONNECTIONSTRING
$Context = New-AzStorageContext -ConnectionString $azstoragestring
if((Get-AzStorageContainer -Context $Context).Name -contains "lastlog"){
    #Set Container
    $Blob = Get-AzStorageBlob -Context $Context -Container (Get-AzStorageContainer -Name "lastlog" -Context $Context).Name -Blob "lastlog.log"
    $lastlogTime = $blob.ICloudBlob.DownloadText()
    $startTime = $lastlogTime | Get-Date -Format yyyy-MM-ddTHH:mm:ss
    $endTime | Out-File "$env:TEMP\lastlog.log"
    Set-AzStorageBlobContent -file "$env:TEMP\lastlog.log" -Container (Get-AzStorageContainer -Name "lastlog" -Context $Context).Name -Context $Context -Force
}
else {
    #create container
    $azStorageContainer = New-AzStorageContainer -Name "lastlog" -Context $Context
    $endTime | Out-File "$env:TEMP\lastlog.log"
    Set-AzStorageBlobContent -file "$env:TEMP\lastlog.log" -Container $azStorageContainer.name -Context $Context -Force
    $startTime = $currentUTCtime.AddSeconds(-300) | Get-Date -Format yyyy-MM-ddTHH:mm:ss
}
$startTime
$endTime
$lastlogTime


$headerParams = Get-AuthToken $env:clientID $env:clientSecret $env:domain $env:tenantGuid
Get-O365Data $startTime $endTime $headerParams $env:tenantGuid


# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
