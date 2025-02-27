<#	
	SCRIPT:   ActiveBatch_Instance_Counter_Calculations_v2.ps1
	
	DESC..:   - The script within this document is only for --estimating-- your current usage. 
				The is a PowerShell script, and it is not designed to make any changes to your environment. 
				The script will connect to the Job Scheduler and calculate a minimum number of annual executions. 
				Since it cannot consider instances that have already been purged from the system, you will need to increase the annual count to account for peak execution times. 
				The script is provided 'as-is'.
#>


param 
(
	# hostname of the JSS server.  Default 'localhost'
	[string]$jssHost = "$([system.net.dns]::gethostentry($env:computername).hostname)",
	
	# number of days period to calculate for.  Default last '30' days.
	[int]$daysPeriod = 30
)


# script's execution
write-host "`n$('─'*100)`nInstance Execution Usage Estimator Script" -f 14;

# request Jss hostname if not local
write-host "`n* Connect to this Scheduler server [$jssHost] ?  " -f 13;
$isLocalhost = read-host "> Enter (yes/no) ";
if ($isLocalhost -in ('n','no','0')) {
    $jssHost = read-host "`nEnter Scheduler's server hostname`n> JSS-Hostname ..";
}

try {
    # Scheduler name here is set to localhost but can be changed if the PowerShell module is deployed elsewhere.
    write-host "`nConnecting to scheduler [ Host: $jssHost ]..." -f 14;
	$header = " > ThisHost .......:   {0} `n > ThisUser .......:   {1} `n > CONNECTED-JSS ..:   {2} `n > JSS-VERSION ....:   {3} `n > CONN-STRING ....:   {4} `n";
	try {
		$isAbatModule = $false;
		$jss = Connect-AbatJss -JobScheduler "$jssHost";
		$jssConnStr = ( $jss.GetPropertyValue("Services.Database.ConnectionString") );
		$isAbatModule = $true;
		write-host ( "$header" -f $env:computername, $(whoami), $jss.Endpoint, $jss.Version, $jssConnStr ) -f 14;
	} catch {
		$isAbatModule = $false;
		write-host "AbatModule Connection Failed:  $_ " -f 13;
		write-host "Attempting connection via COM..." -f 7;
		$jssCom = New-Object -ComObject('ActiveBatch.AbatJobScheduler'); "`n"; 
		$jssCom.Connect("$jssHost"); 
		$jssConnStr = ( $jssCom.DbConnectionString );
		write-host ( "$header" -f $env:computername, $(whoami), $jssCom.MachineName, $jssCom.ProductVersion, $jssConnStr ) -f 14;
	}
}
catch {
    write-host "*** Unable connect to the Job Scheduler on [ $jssHost ]" -f 13;
    throw $_;
}

# Get today's date minus 30. Change this number if the result from Difference in Days is less than 1 from this number
$todayMinus = (Get-Date).AddDays(-($daysPeriod+1));

# Get all instances from the associated Scheduler from the root within the last $todayMinus days
write-host "`nRetrieving scheduler instances from date [ $($todayMinus.Date) ]. " -f 11;

try {
	if ($jssConnStr -notmatch 'Provider=|OleDb') { throw "Non_Supported_Provider: $jssConnStr" }
	
	write-host " |─► Attempting to retrive instances via OleDb provider..." -f 14;
	$connectionString = $jssConnStr;
	$conn = New-Object System.Data.OleDb.OleDbConnection($connectionString)
	$conn.open();
	
	$sqlCmd   = (
		"select count(1) 'InstCount' " +
		"`n	,(select top 1 creationtime from instanceProperties with (nolock) " +
		"`n		where BeginExecutionTime between '$($todayMinus.Date)' and '$($todayMinus)') 'CreationTime' " +
		"`nfrom instanceProperties with (nolock) " +
		"`nwhere BeginExecutionTime >= '$($todayMinus.Date)' "
	);
	$readcmd  = New-Object system.Data.OleDb.OleDbCommand($sqlCmd,$conn);
	$readcmd.CommandTimeout = '300';
	
	$sqlData  = New-Object system.Data.OleDb.OleDbDataAdapter($readcmd);
	$datTable = New-Object system.Data.datatable;
	
	# populate data-table (e.g., $datTable.Rows[0][0])
	[void]$sqlData.fill($datTable);
	$conn.close();
	
	# response object
	$instances = [psCustomObject]@{
		Count  = ( $datTable.InstCount );
		CreationTime = ( $datTable.CreationTime );
	}
}
catch {
	write-host " |─► *** Failure occured on OleDb request:  [ $_ ]" -f 13;
	write-host " |─► Attempting to retrive instances via AbatModule instead. This will take some time, please wait...`n" -f 14;
	try {
		# abatmodule instances method
		$abInstances = ( Get-AbatInstances -JobScheduler $jss -completed -ObjectKey "/" -StartDate $todayMinus.date -Limit 100000000 );
		
		# response object
		$instances = [psCustomObject]@{
			Count  = ( $abInstances.Count );
			CreationTime = ($abInstances[$($abInstances.Count-1)].CreationTime);
		}
	}
	catch {
		throw [system.exception]::New("AbatModule Failure, exiting script:  $_");
	}
}


# Estimates and Calculations
$instanceEstimate = ($instances.count);
$earliestInstance = ($instances.CreationTime);
$totDaysSince     = ((Get-Date) - $earliestInstance.date);
$annualEstimate   = ([math]::ceiling( ($instanceEstimate / [decimal]::ceiling($totDaysSince.totaldays)) * 365));
$daysLow          = $(if ($totDaysSince.days -le 0){ 0.9 } else { $totDaysSince.days } );
$annualEstimate2  = ([math]::ceiling( ($instanceEstimate / ($daysLow)) * 365));
$numFrm           = "{0:#,###,###,###}";


# output formatting functions
function centerText ($txt, $chr='─')
{
	$pads = ("$chr" * ((100-4-$txt.length)/2));
	return  ("$('█'*100)`n$pads[ $txt ]$pads");
}

function frmPad($msg, $val, $pad=45)
{
	return " ► $( "$msg ".padright($pad, '.') ) :   $($val)";
}

# Main output █─►
write-host "";
centerText -txt "Usage Estimate Summary" | write-host -f 11;
write-host (
	"`n"+ ( frmPad -msg "The Earliest Instance Date is" -val $earliestInstance ) ,
	"`n"+ ( frmPad -msg "Difference in Days from Today is" -val $totDaysSince.days ) ,
	"`n"+ ( frmPad -msg "Total Est. Instance Count for [ $($totDaysSince.days) ] days" -val ($numFrm -f $instanceEstimate) ) ,
	"`n"+ ( frmPad -msg "Total Estimated Annual Executions (1-year)" -val ($numFrm -f $annualEstimate +" ~ "+ $numFrm -f $annualEstimate2) ) ,
    "`n`n$('█'*100)"
) -f 11;




# end of scripts
return;


