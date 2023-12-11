# Michael Schmidt (C) 2023
# The script is intended to serve as an example.
# Export a calendar from Exchange public folder and convert it into an iCalendar file (PowerShell).
# The script use snippets from the script Search-Appointments.ps1 from https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts/blob/master/Legacy/Search-Appointments.ps1.
# You need the EWS Managed API from here: https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts/blob/master/Legacy/Microsoft.Exchange.WebServices.dll.
# Version 0.1

# temp file for ics creation
$tempFile = 'C:\ProgramData\IntranetCalendar\Temp\ics.tmp'

# The file with events in iCalendar format (webserver location).
$icsFile = '\\wpserver\wordpress$\intranet\calendar\intranet.ics'

# Events are selected from ... to ...
# days before today
$daysEventBefore = 30
# days after today 
$daysEventTo = 1825

# Logging true or false
$loggingEnabled = $true
$logFile = 'C:\ProgramData\IntranetCalendar\Logs\updateIntranetCalendar.log'

# # Path to the Public folder with the calendar, from which events should be retrieved
$publicFolderPath = "\PublicFolder\Intranet" 
# Account with permission to the public folder
$useEncryptedPassword = $true
$credentialName = "service"
# encrypted string -> convert secure string to encrypted string, e. g.
# 1. Run a powershell session as service account
# 2. $secureString = ConvertTo-SecureString -AsPlainText -Force -String "password"
# 3. ConvertFrom-SecureString -SecureString $secureString -> credentialPassword
$credentialPassword = '01000000d08c9ddf0115d1118c7a00c04fc297eb0100000011469eb8224427459479'
$credentialDomain = 'DOMAIN'
# Exchange On-premises EWS URL
$ewsUrl = "https://exchange.domain.com/EWS/Exchange.asmx"

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource='
    namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
        public TrustAll() {
        }
        public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert,
        System.Net.WebRequest req, int problem) {
        return true;
        }
    }
    }
'
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
## end code from http://poshcode.org/624

# I used the iCalendar validator for the files to be created.
# The following error kept occurring: "Lines not delimited by CRLF sequence near line" - 1 Reference: RFC 5545 3.1. Content Lines
# I found a solution here: https://github.com/PowerShell/PowerShell/issues/16511.
$PSDefaultParameterValues['Out-File:LineTerminator'] = 'CRLF'

# functions
Function ErrorReported($Context)
{
    # Check for any error, and return the result ($true means a new error has been detected)

    # We check for errors using $Error variable, as try...catch isn't reliable when remoting
    if ([String]::IsNullOrEmpty($Error[0])) { return $false }

    # We have an error, have we already reported it?
    if ($Error[0] -eq $script:LastError) { return $false }

    # New error, so log it and return $true
    $script:LastError = $Error[0]
    if ($Context)
    {
        LogError "($Context): $($Error[0])" Red
    }
    else
    {
        LogError "$($Error[0])" Red
    }
    return $true
}
Function ReportError($Context)
{
    # Reports error without returning the result
    ErrorReported $Context | Out-Null
}

# EventType: ERROR, DEBUG and INFO
function LogToFile([String]$EventType,[string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
	$logTime = Get-Date -Format dd.MM.yyyy-hh:mm:ss
    $logInfo = "$logTime $EventType `"$Details`""
    if ($FastFileLogging)
    {
        # Writing the log file using a FileStream (that we keep open) is significantly faster than using out-file (which opens, writes, then closes the file each time it is called)
        $fastFileLogError = $Error[0]
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            Write-Verbose "Opening/creating log file: $LogFile"
            $script:logFileStream = New-Object IO.FileStream($LogFile, ([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
                $script:logFileStream = $null
            }
        }
        if ($script:logFileStream)
        {
            if (!$script:logFileStreamWriter)
            {
                $script:logFileStreamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            }
            $script:logFileStreamWriter.WriteLine($logInfo)
            $script:logFileStreamWriter.Flush()
            if ( $Error[0] -ne $fastFileLogError )
            {
                $FastFileLogging = $false
                Write-Host "Fast file logging disabled due to error: $Error[0]" -ForegroundColor Red
            }
            else
            {
                return
            }
        }
    }
	$logInfo | Out-File -Encoding utf8 $LogFile -Append
}

Function LogError([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
    LogToFile "ERROR" $Details
}

Function LogInfo([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
    LogToFile "INFO" $Details
}

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
    if ( !$VerboseLogFile -and !$DebugLogFile -and ($VerbosePreference -eq "SilentlyContinue") ) { return }
    LogToFile "DEBUG" $Details
}

Function LogDebug([string]$Details)
{
    Write-Debug $Details
    if (!$DebugLogFile -and ($DebugPreference -eq "SilentlyContinue") ) { return }
    LogToFile "DEBUG" $Details
}

Function LoadEWSManagedAPI
{
	# Find and load the managed API
    $ewsApiLocation = @()
    $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
    # ReportError "LoadEWSManagedAPI"

    if (!$ewsApiLoaded)
    {
        # Failed to load the EWS API, so try to install it from Nuget
        $ewsapi = Find-Package "Exchange.WebServices.Managed.Api"
        if ($ewsapi.Entities.Name.Equals("Microsoft"))
        {
	        # We have found EWS API package, so install as current user (confirm with user first)
	        Write-Host "EWS Managed API is not installed, but is available from Nuget.  Install now for current user (required for this script to continue)? (Y/n)" -ForegroundColor Yellow
	        $response = Read-Host
	        if ( $response.ToLower().Equals("y") )
	        {
		        Install-Package $ewsapi -Scope CurrentUser -Force
                $ewsApiLoaded = $(LoadLibraries -searchProgramFiles $true -dllNames @("Microsoft.Exchange.WebServices.dll") -dllLocations ([ref]$ewsApiLocation))
                # ReportError "LoadEWSManagedAPI"
	        }
        }
    }

    if ($ewsApiLoaded)
    {
        if ($ewsApiLocation[0])
        {
            LogInfo "Using EWS Managed API found at: $($ewsApiLocation[0])" Gray
            $script:EWSManagedApiPath = $ewsApiLocation[0]
        }
        else
        {
            Write-Host "Failed to read EWS API location: $ewsApiLocation"
            Exit
        }
    }

    return $ewsApiLoaded
}

Function LoadLibraries()
{
    param (
        [bool]$searchProgramFiles,
        $dllNames,
        [ref]$dllLocations = @()
    )
    # Attempt to find and load the specified libraries

    foreach ($dllName in $dllNames)
    {
        # First check if the dll is in current directory
        LogDebug "Searching for DLL: $dllName"
        $dll = $null
        try
        {
            $dll = Get-ChildItem $dllName -ErrorAction SilentlyContinue
        }
        catch {}

        if ($searchProgramFiles)
        {
            if ($dll -eq $null)
            {
	            $dll = Get-ChildItem -Recurse "C:\Program Files (x86)" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            if (!$dll)
	            {
		            $dll = Get-ChildItem -Recurse "C:\Program Files" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq $dllName ) }
	            }
            }
        }
        $script:LastError = $Error[0] # We do this to suppress any errors encountered during the search above

        if ($dll -eq $null)
        {
            LogError "Unable to load locate $dll" Red
            return $false
        }
        else
        {
            try
            {
		        LogInfo ([string]::Format("Loading {2} v{0} found at: {1}", $dll.VersionInfo.FileVersion, $dll.VersionInfo.FileName, $dllName))
		        Add-Type -Path $dll.VersionInfo.FileName
                if ($dllLocations)
                {
                    $dllLocations.value += $dll.VersionInfo.FileName
                    ReportError
                }
            }
            catch
            {
                ReportError "LoadLibraries"
                return $false
            }
        }
    }
    return $true
}

Function CreateEWSService {
	# set Exchange version  
	$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
	# create Exchange service object  
	$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
	$service.URL = New-Object Uri($ewsUrl)
	# set credentials to EWS service
	if ($useEncryptedPassword) {
		try {
			$credentialPassword = $credentialPassword | ConvertTo-SecureString
			$credentials = New-Object System.Management.Automation.PSCredential($credentialName, $credentialPassword)
			$service.Credentials = $credentials.GetNetworkCredential()
		} catch {
			$exception = $_
			LogError "Error while get password from encrypted string $($exception)." Red
			Exit 1
		}
	} else {
		$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($credentialName, $credentialPassword, $credentialDomain)
	}
	return $service
}

Function FolderIdFromPath {
    param (
            $FolderPath = "$( throw 'Folder Path is a mandatory Parameter' )"
          )
    process{
        ## Find and Bind to Folder based on Path  
        #Define the path to search should be seperated with \  
        try {
			$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)   
			$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$folderid)  
			#Split the Search path into an array  
			$fldArray = $FolderPath.Split("\") 
			 #Loop through the Split Array and do a Search for each level of folder 
			for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
				#Perform search based on the displayname of each folder level 
				$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
				$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
				$findFolderResults = $exchangeService.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
				if ($findFolderResults.TotalCount -gt 0){ 
					foreach($folder in $findFolderResults.Folders){ 
						$tfTargetFolder = $folder                
					} 
				} 
				else{ 
					"Error Folder Not Found"  
					return $null
				}     
			}  
			if($tfTargetFolder -ne $null){
				return $tfTargetFolder.Id.UniqueId.ToString()
			}
		} catch {
			$exception = $_
			LogError "Error while connecting to the Exchange Server: $($exception)." Red
			Exit 1
		}
    }
}

Function ConvertTo-Ics
{
    [CmdletBinding(PositionalBinding=$false,
                  ConfirmImpact='Low')]
    Param
    (
		# Subject of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Title",'Name')]
        [string]
        $subject,

        # Start Time of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("StartTime", 'Starts')]
        [datetime]
        $DTStart,

        # End Time of event
        [Parameter(ParameterSetName='EndTime',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("EndTime",'Ends')]
        [datetime]
        $DTEnd,

        # Location of event
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Room","Track")]
        [string]
        $Location,

        # is an all day event
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("All day event")]
        [bool]
        $IsAllDay
    )
    Process
    {
        function Add-LineFold ([String]$Text) {
            # Simpel implementation to comply with https://icalendar.org/iCalendar-RFC-5545/3-1-content-lines.html
            $x = 60
            while($x -lt $text.Length) {
                $text = $text.Insert($x, "`r`n")
                $x = $x + 60
            }
            $text
        }

        $Start = "{0:yyyyMMddTHHmmss}" -f [Datetime]::ParseExact($DTStart, 'MM/dd/yyyy HH:mm:ss', $null)
        $End = "{0:yyyyMMddTHHmmss}" -f [Datetime]::ParseExact($DTEnd, 'MM/dd/yyyy HH:mm:ss', $null)
        $subject = Add-LineFold -Text "SUMMARY:$subject"
        $Location = Add-LineFold -Text "LOCATION:$Location"
        $Description = Add-LineFold -Text "DESCRIPTION:$Description"
		$UID = (Get-Random).ToString()
		
		"BEGIN:VEVENT" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"UID:$UID" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"DTSTAMP:20230119T070000Z" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"DTSTART:$Start" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"DTEND:$End" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"SEQUENCE:1" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"$subject" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"$Location" | Out-File -Encoding utf8 -FilePath $tempFile -Append
		"END:VEVENT" | Out-File -Encoding utf8 -FilePath $tempFile -Append
    }
}

# start
LogInfo "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting." Green

# checks ... to be continued
# ics file
if ( [String]::IsNullOrEmpty($icsFile) ) { 
	LogError "Path to ics file not set, cannot continue" Red
	Exit 1
}
$icsFileDir = Split-Path -Path $icsFile -Parent
if (-Not (Test-Path -Path $icsFileDir -PathType Container)) {
	LogError "Directory $($icsFileDir) for ics file not exists, cannot continue." Red
	Exit 1
}

# First try to load the EWS Managed API
if (!(LoadEWSManagedAPI))
{
	LogError "Failed to locate EWS Managed API, cannot continue." Red
	Exit 1
}

# connect to Exchange Server
$exchangeService = CreateEWSService
if (!($exchangeService))
{
	LogError "Failed to connect, cannot continue." Red
	Exit 1
}

# calulate the dates
$today = Get-Date
$startsAfterDate = ($today.AddDays(- $daysEventBefore))
$endsBeforeDate = ($today.AddDays($daysEventTo))

# create the icalendar file
# load the calendar foder object
$calenderFolderID = FolderIdFromPath -FolderPath $publicFolderPath
$SubFolderId =  New-Object Microsoft.Exchange.WebServices.Data.FolderId($calenderFolderID)
$SubFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService, $SubFolderId)
if($SubFolder -ne $null) {
	# create the ics file
	LogInfo "Creating ics temp file ($($tempFile)) from Exchange calendar..." Green
	"BEGIN:VCALENDAR" | Out-File -Encoding utf8 -FilePath $tempFile
	"PRODID:-//PowerShell/COOP//NONSGML v1.0//EN" | Out-File -Encoding utf8 -FilePath $tempFile -Append
	"VERSION:2.0" | Out-File -Encoding utf8 -FilePath $tempFile -Append
    #Define ItemView to retrive just 1000 Items    
    $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
    $fiItems = $null    
    do {    
        $appointments = $exchangeService.FindItems($SubFolder.Id, $ivItemView)    
	LogInfo "Calendar has $($appointments.Items.Count) items." Green
        Foreach($appointment in $appointments.Items){
			# filter by Date
			$appointmentStartDate = $appointment.Start
			$appointmentEndDate = $appointment.End
			if ($appointmentStartDate -ge $startsAfterDate -And $appointmentStartDate -le $endsBeforeDate) {
				ConvertTo-Ics -Subject $appointment.Subject -DTStart $appointmentStartDate -DTEnd $appointmentEndDate -Location $appointment.Location
			}
        }    
        $ivItemView.Offset += $appointments.Items.Count    
    } while($appointments.MoreAvailable -eq $true)
	
	# end the ics file
	"END:VCALENDAR"  | Out-File -Encoding utf8 -FilePath $tempFile -Append
	LogInfo "Creating ics temp file ($($tempFile)) has finished." Green
} else {
	LogError "Failed to get calendar folder object." Red
	Exit 1
}

# copy the ics file to the webserver location
LogInfo "Copy the ics temp file ($($tempFile)) to $($icsFile)." Green
try {
  Copy-Item $tempFile -Destination $icsFile -force
} catch {
  $exception = $_
  LogError "Error while copying the ics file: $($exception)." Red
  Exit 1
}
Exit 0