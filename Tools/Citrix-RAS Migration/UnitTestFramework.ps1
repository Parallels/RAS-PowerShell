<#
.SYNOPSIS
This is a basic unit testing framework.

.DESCRIPTION
This is a basic unit testing framework.

.EXAMPLE
. "./UnitTestFramework.ps1"
$TestSuiteResults = @()
$TestSuiteResults += TestSuite -name "My Test Suite" {
	{
		UnitTest "My UnitTest" {
			throw "Failing unit test"
		}
	}
	{
		UnitTest "UnitTest with RAS session" {
			RASSessionContext {
				$folder = Get-PubFolder
				Assert ($folder) -message "Failed to get folder."
				Assert ($folder.Name -eq "TestFolder") -message "Unexpected folder name" -expected "TestFolder" -actual $folder.Name
			} # Remove-RASSession is called and clean up is done.
		}
	}
	{
		UnitTest "Some other test" {
			try	{
				$folder = Get-PubFolder
			}
			catch {
				Assert ($false) -message "Something wrong." -exception $_.Exception
			}
		}
	}
}

foreach ($TestSuiteResult in $TestSuiteResults) {
    Log -type "INFO" -message $TestSuiteResult.Header
    foreach ($test in $TestSuiteResult.UnitTests) {
        Log -type "INFO" -message "[UnitTest: $($test.TestName)]" -testInfo $test.Info -testStatus $test.Status
    }
    Log -type "INFO" -message $TestSuiteResult.Footer
}

.NOTES
TestSuite returns a dictionary with the name of the TestSuite and the results of all the executed unit tests.
UnitTest function is not required to be run in TestSuite.
RASSessionContext requires the Defaults.ps1 file to be created and Username, Password, and Server to be defined.
if the $Server is not defined in Defaults.ps1 $Server will be assigned  to localhost
#>

. "./Utilities.ps1"

function TestSuite ([string] $name, [scriptblock[]] $unitTests) {

	if (-not $name) {
		throw "Test suite must have a name."
	}

	if (-not $unitTests) {
		throw "Test suite must have at least 1 unit test."
	}

	[uint16]$failed = 0;
	[uint16]$passed = 0;
	[float]$testSuiteIntegrity = 0;
	$tests = @()

	log -type "INFO" -message "Started '$name' test suite."
	foreach ($test in $unitTests) {
		try {
			$testData = $test.Invoke()
		}
		catch [Exception] {
			Log -type "ERROR" -message "$($_.Exception.Message)"
			$failed++
			throw
		}
		$tests += $testData
		$passed += ([bool]$testData.Status -eq $true)
		$failed += ([bool]$testData.Status -eq $false)
	}

	$testSuiteIntegrity = ($passed / ($passed + $failed) * 100)
	if ($testSuiteIntegrity -ge 0 -and $testSuiteIntegrity -lt 100) {
		$status = $false
	}
	else {
		$status = $true
	}
	log -type "INFO" -message "Test success rate: $($testSuiteIntegrity)%." -fnName "TestSuite" -testStatus $status
	log -type "INFO" -message "'$name' test suite completed. Number of tests ran $($passed + $failed)." -fnName "TestSuite"

	# Log -type "INFO" -message "[$("="*20) [$($name.ToUpper()) TEST RESULTS] $("="*20)]"
	# foreach ($test in $tests) {
	#     Log -type "INFO" -message "[UnitTest: $($test.TestName)][Info: $($test.Info)]" -testStatus $test.Status
	# }

	return @{
		Header    = "[$("="*20) [$($name.ToUpper()) TEST RESULTS] $("="*20)]"
		UnitTests = $tests
		Footer    = "[$("="*20) [$($passed)/$($passed+$failed) Tests passed.] $("="*20)]"
	}
}

function UnitTest([string] $name, [scriptblock] $fn) {
	if (-not $name) {
		throw "Unit test must have a name."
	}

	if (-not $fn) {
		throw "Callback 'fn' must not be null."
	}

	Log -type "INFO" -message "[UnitTest:$($name)][STARTED]" -fnName "UnitTest"
	try {
		$fn.Invoke()
	}
	catch {
		Log -type "INFO" -message "[UnitTest:$($name)][FAILED]" -fnName "UnitTest" -exception $_.Exception
		return @{
			TestName = $name
			Status   = $false
			Info     = $_.Exception.Message
		}
		# throw

	}
	Log -type "INFO" -message "[UnitTest:$($name)][SUCCESS]" -fnName "UnitTest"
	return @{
		TestName = $name
		Status   = $true
		Info     = $null
	}
}

function CleanupRASSessionContext () {

	$sites = Get-Site
	$servers = Get-RDS
	$groups = Get-RDSGroup
	$apps = Get-PubRDSApp
	$folders = Get-PubFolder

	if ($sites[0].Id -ne 1) {
		Log -type "INFO" -message "Removing RAS Sites ..."
		foreach ($site in $sites) {
			try {
				Remove-Site -InputObject $site
				Log -type "INFO" -message "Removed site $($site.Name)."
			}
			catch {
				Log -type "ERROR" -message "Failed to remove site" -exception $_.Exception
			}
		}
	}

	Log -type "INFO" -message "Removing RDS servers ..."
	foreach ($server in $servers) {
		try {
			Log -type "INFO" -message "Removed RDS server $($server.Server)"
			Remove-RDS -InputObject $server -NoUninstall

		}
		catch {
			Log -type "ERROR" -message "Failed to remove RDS" -exception $_.Exception
		}
	}

	Log -type "INFO" -message "Removing RAS groups ..."
	foreach ($group in $groups) {
		try {
			Remove-RDSGroup -InputObject $group
			Log -type "INFO" -message "Removed RAS group $($group.Name)"
		}
		catch {
			Log -type "ERROR" -message "Failed to remove RAS group." -exception $_.Exception
		}
	}

	Log -type "INFO" -message "Removing RAS published applications"
	foreach ($app in $apps) {
		try {
			Remove-PubRDSApp -InputObject $app
			Log -type "INFO" -message "Removed application $($app.Name)"
		}
		catch {
			Log -type "EROR" -message "Failed to remove RDS application." -exception $_.Exception
		}
	}

	Log -type "INFO" -message "Removing published folders"
	$childFolders = @()
	foreach ($folder in $folders) {
		try {
			[PSAdmin.PubFolder]$folder = $folder
			Remove-PubFolder -InputObject $folder
			Log -type "INFO" -message "Removed published folder $($folder.Name)."
		}
		catch {
			Log -type "WARNING" -message "Failed remove published folder." -exception $_.Exception
		}
	}

	try {
		Invoke-Apply
	}
	catch {
		Log -type "ERROR" -message "Failed to apply setting on RAS." -exception $_.Exception
	}
}

function RASSessionContext ([scriptblock] $fn) {
	if (-not $fn) {
		throw "Call back function cannot be null!"
	}
	Remove-RASSession
	try {
		Log -type "INFO" -message "Creating new RAS Session ..."
		if (-not $Server) {
			$Server = "localhost"
			Log -type "WARNING" -message "'Server' was not defined. Using '$Server' as server ..."
		}
		$Password = ConvertTo-SecureString -String $Password -AsPlainText -Force
		New-RASSession -Username $Username -Password $Password -Server "$Server"
		Log -type "INFO" -message "Created RAS session"
	}
	catch {
		Log -type "ERROR" -message "Failed to create RAS session. Error was: $($_.Exception.Message)"
		throw $_.Exception.Message
	}

	try {
		$fn.Invoke()
	}
	catch {
		# CleanupRASSessionContext
		Remove-RASSession
		Log -type "INFO" -message "Removing RAS session"
		throw
	}
	# CleanupRASSessionContext
	Remove-RASSession
	Log -type "INFO" -message "RAS session removed"
}


function Assert($expression, [string] $message, [Exception] $innerException = $null, $actual = $null, $expected = $null) {
	if (-not $expression) {
		$stackinfo = GetCmdletCallDetails
		if ($innerException) {
			$message += " InnerException: $($innerException.Message)"
		}
		if ($actual -ne $null -and $expected -ne $null) {
			$message += "[Actual: $($actual)][Expected: $($expected)]"
		}
		throw "[$($stackinfo.CallerName)][line: $($stackinfo.LineNumber[2])] ---> $($message)"
	}
}