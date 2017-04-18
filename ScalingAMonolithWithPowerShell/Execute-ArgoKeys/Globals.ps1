#region Show/Hide PowerShell Window
Add-Type -Name Window -Namespace Console -MemberDefinition @"
		[DllImport("Kernel32.dll")]
		public static extern IntPtr GetConsoleWindow();

		[DllImport("user32.dll")]
		public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
"@
function Show-Console {
	$consolePtr = [Console.Window]::GetConsoleWindow()
	[Console.Window]::ShowWindow($consolePtr, 5)
}

function Hide-Console {
	$consolePtr = [Console.Window]::GetConsoleWindow()
	[Console.Window]::ShowWindow($consolePtr, 0)
}
#endregion

switch ($environment) {
	'Test' {
		$argoServer = 'ArgoTest'
		$argoServerDir = "\\$argoServer\ADS\TEST"
		$argoLocalDir = 'C:\ADS\TEST'
	}
	'Training' {
		$argoServer = 'ArgoTrain'
		$argoServerDir = "\\$argoServer\ADS\TRAIN"
		$argoLocalDir = 'C:\ADS\TRAIN'
	}
	default {
		$argoServer = 'ArgoProd'
		$argoServerDir = "\\$argoServer\ADS\PROD"
		$argoLocalDir = 'C:\ADS\CLIENT'
	}
}
#region Helper Functions
function ConvertTo-DataTable {
<#
	.SYNOPSIS
		A brief description of the ConvertTo-DataTable function.

	.DESCRIPTION
		A detailed description of the ConvertTo-DataTable function.

	.PARAMETER Source
		An array that needs converted to a DataTable object

	.PARAMETER Match
		A description of the Match parameter.

	.PARAMETER NotMatch
		A description of the NotMatch parameter.

	.PARAMETER InsertHeaderRow
		Add a top row to the DataTable, useful if you want to force to make a selection.

	.EXAMPLE
		$DataTable = ConvertTo-DataTable $Source
#>
	[CmdLetBinding(DefaultParameterSetName = "None")]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.Array]$Source,
		[Parameter(Position = 1, ParameterSetName = 'Like')]
		[String]$Match = ".+",
		[Parameter(Position = 2, ParameterSetName = 'NotLike')]
		[String]$NotMatch = ".+",
		[Parameter()]
		[String]$InsertHeaderRow
	)
	if ($NotMatch -eq ".+") {
		$Columns = $Source[0] | Select-Object * | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match "($Match)" }
	} else {
		$Columns = $Source[0] | Select-Object * | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -notmatch "($NotMatch)" }
	}
	$DataTable = New-Object System.Data.DataTable
	foreach ($Column in $Columns.Name) {
		$DataTable.Columns.Add("$($Column)") | Out-Null
	}

	#Create header Row
	if ($InsertHeaderRow) {
		$Row = $DataTable.NewRow()
		foreach ($Column in $Columns.Name) {
			$Row["$($Column)"] = $InsertHeaderRow
		}
		$DataTable.Rows.Add($Row)
		$rowCount = $Source.Count + 1
	} else {
		$rowCount = $Source.Count
	}

	#For each row (entry) in source, build row and add to DataTable.
	foreach ($Entry in $Source) {
		$Row = $DataTable.NewRow()
		foreach ($Column in $Columns.Name) {
			$Row["$($Column)"] = if ($Entry.$Column -ne $null) { ($Entry | Select-Object -ExpandProperty $Column) -join ', ' } else { $null }
		}
		$DataTable.Rows.Add($Row)
	}
	#Validate source column and row count to DataTable
	if ($Columns.Count -ne $DataTable.Columns.Count) {
		throw "Conversion failed: Number of columns in source does not match data table number of columns"
	} else {
		if ($rowCount -ne $DataTable.Rows.Count) {
			throw "Conversion failed: Source row count not equal to data table row count"
		}
		#The use of "Return ," ensures the output from function is of the same data type; otherwise it's returned as an array.
else {
			Return, $DataTable
		}
	}
}

function ConvertTo-IPSubnet {
    <#
      .Synopsis
        Converts an IP Address and Subnet Mask to IP Subnet
      .Description
        ConvertTo-IPSubnet takes an IP Address and Subnet Mask and returns an IP Subnet.
      .Parameter IPAddress
        An IP Address to convert.
      .Parameter SubnetMask
        A Subnet Mask to convert.
    #>

	[CmdLetBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$IPAddress,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$SubnetMask
	)
	begin {
		function toBinary ($dottedDecimal) {
			$dottedDecimal.split('.') | ForEach-Object{ $binary += $([convert]::ToString($_, 2).PadLeft(8, '0')) }
			return $binary
		}

		function toDottedDecimal ($binary) {
			do {
				$dottedDecimal += '.' + [string]$([convert]::ToInt32($binary.Substring($i, 8), 2)); $i += 8
			}
			while ($i -le 24)
			return $dottedDecimal.Substring(1)
		}
	}
	process {
		$ipBinary = toBinary $IPAddress
		$smBinary = toBinary $SubnetMask
		$netBits = $smBinary.indexOf("0")

		#Validate the subnet mask
		if (($smBinary.length -ne 32) -or ($smBinary.Substring($netBits).Contains('1') -eq $true)) {
			Write-Warning 'Subnet Mask is invalid!'
			Exit
		}

		#Validate the IP address
		if (($ipBinary.length -ne 32) -or ($ipBinary.Substring($netBits) -eq '00000000') -or ($ipBinary.Substring($netBits) -eq '11111111')) {
			Write-Warning 'IP Address is invalid!'
			Exit
		}

		#Return network information
		Return [PSCustomObject]@{
			'NetworkID' = $(toDottedDecimal $($ipBinary.Substring(0, $netBits).padright(32, '0')))
			'FirstAddress' = $(toDottedDecimal $($ipBinary.Substring(0, $netBits).padright(31, '0') + '1'))
			'LastAddress' = $(toDottedDecimal $($ipBinary.Substring(0, $netBits).padright(31, '1') + '0'))
			'BroadcastAddress' = $(toDottedDecimal $($ipBinary.Substring(0, $netBits).padright(32, '1')))
		}
	}
	end { }
}

function Execute-SqlQuery {
	<#
	.SYNOPSIS
	Connects to SQL Database and returns query results

	.DESCRIPTION
	Connects to specified SQL Server/Database using Integrated Windows Authentication
	Runs specified SQL Query and Returns a DataRecord object

	.PARAMETER sqlServer
	Name of the SQL Server that hosts the Xperience database

	.PARAMETER databaseName
	Name of the database that hosts the Xperience updates

	.PARAMETER  sqlQuery
	SQL query to return package data from Xperience database (T-SQL syntax)

	.EXAMPLE
	PS C:\> Execute-SqlQuery -sqlServer sql01 -databaseName Xperience -sqlQuery 'SELECT * FROM Package'
	Connects to Xperience database on sql01 and returns DataRecord containing results

	.EXAMPLE
	PS C:\> $records = Execute-SqlQuery -sqlServer sql01 -databaseName Xperience -sqlQuery 'SELECT * FROM Package'
	Connects to Xperience database on sql01 and returns DataRecord containing results

	.INPUTS
	System.String, System.String, System.String

	.OUTPUTS
	System.Data.Common.DataRecordInternal

	.NOTES
	Created By: Chris Brucker
	Creation Date: 20160811
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.String]$sqlServer,
		[Parameter(Position = 1, Mandatory = $true)]
		[System.String]$databaseName,
		[Parameter(Position = 2, Mandatory = $true)]
		[System.String]$sqlQuery
	)

	begin { }
	process {
		try {
			# Open ADO.NET Connection to SQL Server and Database using Windows authentification.
			$adoConnection = New-Object -TypeName Data.SqlClient.SqlConnection
			$adoConnection.ConnectionString = "Data Source=$sqlServer;Initial Catalog=$databaseName;Integrated Security=True;"
			$adoConnection.Open()
			Write-Verbose -Message "Connected to database [$databaseName] on server [$sqlServer]"

			# Query the SQL Database, timeout after 2 minutes (120 seconds)
			$sqlCmd = New-Object -TypeName Data.SqlClient.SqlCommand -ArgumentList $sqlQuery, $adoConnection
			$sqlCmd.CommandTimeout = 120

			# Build a data adapter so that we can fill the data table
			$sqlAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter
			$sqlAdapter.SelectCommand = $sqlCmd
			Write-Verbose -Message "Executing Query [$sqlQuery]"

			$dataTable = New-Object -TypeName System.Data.DataTable
			$sqlAdapter.Fill($dataTable) | Out-Null
			Write-Verbose -Message 'Filled DataTable'

			# Close connections and cleanup objects
			$sqlAdapter.Dispose()
			$sqlCmd.Dispose()
			$adoConnection.Close()
			Remove-Variable -Name adoConnection, sqlCmd, sqlAdapter

			[gc]::Collect()

			return $dataTable
		} catch {
			throw $_.Exception.Message
		}
	}
	end { }
}

function Get-CMDistributionPointBySubnet {
    <#
      .Synopsis
        Returns the ConfigMgr Distribution
      .Description
        Get-CMDistributionPointBySubnet takes a Subnet ID and returns the Distribution Point and Site Code associated with that Subnet.
	  .Parameter SubnetID
        A Subnet ID or list of Subnet ID's.
	  .Parameter Boundaries
        Array of Boundary Information, optional parameter
      .Parameter Path
        Path to CSV file containing boundary information, optional parameter
    #>

	[CmdLetBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string[]]$SubnetIDs,
		[Parameter(ParameterSetName = 'Boundaries')]
		[ValidateNotNullOrEmpty()]
		$Boundaries,
		[Parameter(ParameterSetName = 'Path')]
		[ValidateNotNullOrEmpty()]
		[string]$Path
	)
	Begin {
		if ($Path) {
			$Boundaries = Import-Csv -Path $Path
		}
		$DPFound = $false
	}
	Process {
		Foreach ($Boundary in $Boundaries) {
			if (-not $DPFound) {
				Foreach ($Subnet in $SubnetIDs) {
					## VPN/Wireless Subnets; Set Distribution Point for each data center
					if ($Subnet -eq '10.15.4.0' -or $Subnet -eq '10.14.8.0') {
						$DPFound = $false
						$onWireless = $true
						$siteCode = 'DC'
						$location = 'Primay DataCenter'
						$distributionPoint = $argoServer
					} elseif ($Subnet -eq '10.31.8.0' -or $Subnet -eq '10.30.8.0') {
						$DPFound = $false
						$onWireless = $true
						$siteCode = 'DC2'
						$location = 'Secondary DataCenter'
						$distributionPoint = 'SCCMServer'
					} elseIf ($Boundary.Value -eq $Subnet) {
						$DPFound = $true
						Write-Verbose "Found a Match for [$Boundary.Value]"
						Return $Boundary
					}
				}
			}
		}
	}
	End {
		If ($onWireless) {
			$properties = @{
				'SiteCode' = $siteCode;
				'Location' = $location;
				'DistributionPoint' = $distributionPoint
			}
			Return New-Object -TypeName psobject -Property $properties
		}
		If (-not $DPFound) {
			$properties = @{
				'SiteCode' = '';
				'Location' = '';
				'DistributionPoint' = $argoServer
			}
			Return New-Object -TypeName psobject -Property $properties
		}
	}
}

function Reset-LogFile {
<#
	.Synopsis
		Reset's Log File if it is greater than 1MB in size
	.Description
		Gets size of specified log file and deletes it if is greater than
		specified size.
	.Parameter Path
		Location of Log File
	.Parameter MaxFileSize
		File Size
	.PARAMETER FileSizeUnit
		Specify B for Bytes, KB for KiloBytes, MB for MegaBytes
	.EXAMPLE
		PS C:\> Reset-LogFile -Path 'C:\Logs\log.log' -MaxFileSize '1' -FileSizeUnit 'MB'
	.NOTES
		Additional information about the function.
#>

	[CmdLetBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[ValidateScript({ Test-Path -Path $_ })]
		[string]$Path,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.Version]$MaxFileSize,
		[Parameter(Mandatory = $true)]
		[ValidateSet('B', 'KB', 'MB')]
		[string]$FileSizeUnit
	)
	begin {
		$logFile = Get-Item -Path $Path
	}
	process {
		# Convert File Size to appropriate unit
		switch ($FileSizeUnit) {
			'MB' { [System.Version]$fileSize = "{0:N2}" -f ($logFile.Length / 1MB) }
			'KB'{ [System.Version]$fileSize = "{0:N2}" -f ($logFile.Length / 1KB) }
			default { [System.Version]$fileSize = "{0:N2}" -f $logFile.Length }
		}

		if ($fileSize -gt $MaxFileSize) {
			Write-Verbose -Message "Log file will be removed because it is [$fileSize] [$FileSizeUnit] in size which is greater than [$MaxFileSize] [$FileSizeUnit]."
			Remove-Item -Path $Path -Force
		} else {
			Write-Verbose -Message "Log file is [$fileSize] [$FileSizeUnit] in size which is less than [$MaxFileSize] [$FileSizeUnit]."
		}
	}
	end { }
}

function Set-ArgoWWSID {
<#
	.SYNOPSIS
		Allocates a WWSID for a workstation

	.DESCRIPTION
		This script will add the workstation if it does not already exist in the database.
		It will allocate in the database a single WWSID for the workstation. Then output a csv file to the specified destination.

	.PARAMETER ComputerName
		Name of the computer

	.PARAMETER Region
		ArgoKeys Region, defaults to 001

	.PARAMETER Branch
		Branch ID

	.PARAMETER Environment
		ArgoKeys Environment (Production, Test or Training)

	.PARAMETER DistributionPoint
		Name of the ConfigMgr Distribution Point that we will pull content from

	.PARAMETER SQLInstance
		Name of the SQL Instance/Server that we will connect to

	.PARAMETER SQLDB
		Name of the SQL Database that we will connect to

	.PARAMETER CsvFolderDestination
		Path to the folder that we will create the ARGOKEYS file in.

	.EXAMPLE
		Set-ArgoWWSID -ComputerName L87MF4VZ -BranchId 001 -Environment Production -DistributionPoint SR1SC100 -SQLInstance SQL01 -CsvFolderDestination C:\ADS\Client\BAT

	.Outputs
		None

	.NOTES
		Additional information about the function.
#>
	Param (
		[Parameter(Mandatory = $true)]
		[string]$ComputerName,
		[string]$Region = '001',
		[Parameter(Mandatory = $true)]
		[string]$Branch,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Production', 'Test', 'Training')]
		[string]$Environment,
		[Parameter(Mandatory = $true)]
		[string]$DistributionPoint,
		[string]$SQLInstance,
		[string]$SQLDB = 'FFB_ArgoKeysTracker',
		[Parameter(Mandatory = $true)]
		[string]$CsvFolderDestination
	)

	# Create folder path if not exists
	New-Item -Path $CsvFolderDestination -ItemType directory -Force | Out-Null;

	# Set Full file path
	$CsvFullPath = "$CsvFolderDestination\ARGOKEYS"

	Write-Log -Text "We are attempting to fetch a WWSID using the following information:
						SQL Instance:		[$sqlInstance]
						WorkstationName: 	[$ComputerName]
						Branch:			[$Branch]
						EnvironmentName:	[$Environment]
						DistributionPoint:	[$distributionPoint]
						csvFile:		[$CsvFolderDestination\ARGOKEYS]"

	$ServerQuery = "EXEC [dbo].[uspUpdateArgoWWSID]
					@RegionId = '$Region',
					@BranchId = '$Branch',
					@WorkStationName = '$ComputerName',
					@EnvironmentName = '$environment'"

	$argoInfo = Execute-SqlQuery -sqlServer $sqlInstance -databaseName $SQLDB -sqlQuery $ServerQuery

	if ($argoInfo.RC -eq 0) {
		$argoKeysFile = [pscustomobject]@{ 'Computer Name' = $ComputerName; 'Region' = $Region; 'Branch' = $Branch; 'WWS ID' = $argoInfo.WWSID; 'PRINTER' = ''; 'Syncserver' = $distributionPoint }
		$argoKeysFile | ConvertTo-Csv -NoTypeInformation | ForEach-Object { $_ -replace '"', '' } | Out-File -FilePath $CsvFullPath -Encoding ascii
		$lblOutput.Text = "WWSID [$($argoInfo.WWSID)] has been allocated from the available pool to [$ComputerName]. There are [$($argoInfo.RemainingWWSIDs)] WWSIDs remaining for Branch [$branch]."
		Write-Log -Text "WWSID [$($argoInfo.WWSID)] has been allocated from the available pool to [$ComputerName]. There are [$($argoInfo.RemainingWWSIDs)] WWSIDs remaining for Branch [$branch]."
	} elseif ($argoInfo.RC -eq -1) {
		$lblError.Text = "There are no available Argo IDs for Branch: [$Branch] in the [$SQLDB] database. Please call the IT Support Center."
		Write-Log -Text "There are no available Argo IDs for Branch: [$Branch] in the [$SQLDB] database. Please call the IT Support Center."
		return -1
	} else {
		$lblError.Text = "Error while attempting to assign a WWSID for [$Workstation] in the [$SQLDB] database. Please call the IT Support Center."
		Write-Log -Text "Error while attempting to assign a WWSID for [$Workstation] in the [$SQLDB] database. Please call the IT Support Center."
		return 1
	}
}

function Test-BDPConnection {
	<#
	.SYNOPSIS
		Determines if BDP is online.
	.DESCRIPTION
		If BDP is online, it sets the DistributionPoint to itself. If it offline, it sets
		the DistributionPoint to the Argo Server
	.PARAMETER  BDPName
		The name of the Distribution Point for the location.
	.EXAMPLE
		PS C:\> Test-BDPConnection -BDPName 'D7H1VPNW'
		This example shows how to call the Test-BDPConnection function with named parameters.
	.EXAMPLE
		PS C:\> Test-BDPConnection 'D7H1VPNW'
		This example shows how to call the Test-BDPConnection function with positional parameters.
	.INPUTS
		System.String
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.String]$BDPName
	)
	process {
		$DPOnline = Test-Connection -ComputerName $BDPName -Count 2 -ErrorAction SilentlyContinue
		if ($DPOnline) {
			$global:distributionPoint = $BDPName
		} else {
			$global:distributionPoint = $argoServer
		}
	}
}

function Write-Log {
<#
.SYNOPSIS
    Writes output to the console and log file simultaneously
.DESCRIPTION
	This functions outputs text to the console and to the log file specified in the XML configuration.
	The date, time and installation phase is pre-pended to the text, e.g. [30-07-2013 11:27:07] [Initialization] "Deploy Application script version is [2.0.0]"
.EXAMPLE
	Write-Log -Text "This is a custom message..."
.PARAMETER Text
	The text to display in the console and to write to the log file
.PARAMETER PassThru
	Passes the text back to the PowerShell pipeline
#>
	Param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $True)]
		[array]$Text,
		[switch]$PassThru = $false
	)
	Process {
		$Text = $Text -join (" ")
		$currentDate = (Get-Date -UFormat "%m-%d-%Y")
		$currentTime = (Get-Date -UFormat "%T")
		$logEntry = "[$currentDate $currentTime] $Text"
		Write-Host $logEntry
		# Create the Log directory and file if it doesn't already exist
		If (!(Test-Path -Path $logFile -ErrorAction SilentlyContinue)) { New-Item $logFile -ItemType File -ErrorAction SilentlyContinue | Out-Null }
		Try {
			"$logEntry" | Out-File $logFile -Append -ErrorAction SilentlyContinue
		} Catch {
			$exceptionMessage = "$($_.Exception.Message) `($($_.ScriptStackTrace)`)"
			Write-Host "$exceptionMessage"
		}
		If ($PassThru -eq $true) {
			Return $Text
		}
	}
}
#endregion