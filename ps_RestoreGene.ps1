<#--------------------------------------------------------------------------
.SYNOPSIS
    
Restore Generator v8.05 (2016-11-19)
(C) 2012, Paul Brewer
                        
Feedback: paulbrewer@yahoo.co.uk
Updates: https://paulbrewer.wordpress.com/ps_restoregene/
User Guide: https://paulbrewer.wordpress.com/2016/08/05/restoregeneversion6/
       
                       
.DESCRIPTION
Calls stored procedure sp_RestoreGene on the Primary server to get a RESTORE script which is executed on the Standby server
Updates: https://paulbrewer.wordpress.com/2013/10/12/database-restore-automation/
                       
V3.5 - geostock - Updated to include a 'ConsoleFeedback' parameter, defaulted to on. Change to off to suppressed write-host in SQL Agent jobs.
V3.6 - Richard - Variable declare fix, where database name has a space or full stop.
     - EXIT this script if there are no new log backups to restore, don't THROW
V3.9 - SQLMongo fix for hyphen in database name.
V4.0 - Domingo fix for default backup paths
V5.0 - Change default for WithStandby to 0, NoRecovery is the default
V5.1 - Mick Pollock Azure URL backup path support and string replace functionality
V5.2 - June 17th, 2016 - Add DebugMode parameter, writes restore script to xml log file, does not execute commands
V5.4 - July 3rd, 2016 - New parameter for FILESTREAM files, for WITH MOVE parameter, jgkiran
V5.8 - WITH MOVE secondary file names no longer included as separate lines, code tidying modification
     - Requires Version 5.8 or higher of the stored procedure sp_RestoreGene
V6.0 - Improved error handling / feedback
V6.1 - Add Throw on Error
V6.4 - Publish SQL Server Central
V6.41 - PoSh Driver Last LSN minor fix.
V6.44 - Kill connections to Target DB Name if supplied, Bug fix identified by Rod.
V6.5  - New 'Drop Database After Restore' parameter for use with CHECKDB and automated backup verification, John Lee
V6.54 - New Parameter to exclude differential and log backups, for development environment refreshes, Mehedi Amin
V6.57 - New parameters to exclude backup device types, remove variable declares/RAISEERROR, Lars Rasmussen 
V8.04 - April 15th.2018 - Error handling, checking pending restore commands, Thomas Wolf
V8.05 - December 2020 - Import correct SQL Module, Thomas Wolf
                       
---------------------------------------------------------------------------#>
#Script Parameters
Param(
                       
        #Name and path for the restore log
        [Parameter(Mandatory=$true)]
        $RestoreLog,
                       
        #Primary server, SQL instance
        [Parameter(Mandatory=$true)]
        $PrimaryServer,
                       
        #Standby server, SQL instance
        [Parameter(Mandatory=$true)]
        $StandbyServer,
                       
        #DBName server, SQL instance
        [Parameter(Mandatory=$true)]
        $DBName,
                       
        #The TargetDBName can override of the restored db name, defaults to DBName is needed.
        $TargetDBName = $null,
                       
        #The WithReplace parameter is needed if over-wrtting the source of the database backup (PrimaryServer=StandbyServer)
        # and if no tail log backup was taken.
        $WithReplace = "0",
                       
        #The WithMoveDataFiles parameter allows the restored database to use a different path for database files.
        $WithMoveDataFiles = $null,
                       
        #The WithMoveLogFile parameter allows the restored database to use a different path for its log file.
        $WithMoveLogFile = $null,
                       
        #The WithMoveFileStreamFile parameter allows the restored database to use a different path for its FileStream files.
        $WithMoveFileStreamFile = $null,
                       
        #The FromFileFullUNC parameter allows overriding the drive & folder path of the full backup files, possibly with network share name.
        $FromFileFullUNC = $null,
                       
        #The FromFileDiffUNC parameter allows overriding the drive & folder path of the diff backup files, possibly with network share name.
        $FromFileDiffUNC = $null,
                       
        #The FromFileLogUNC parameter allows overriding the drive & folder path of the log backup files, possibly with network share name.
        $FromFileLogUNC = $null,
                       
        #The StopAt parameter allows overriding the date / time recovery point to historic backup files, defaults to most current.
        $StopAt = $null,
                       
        #The StandBy parameter allows overriding WITH STANDBY so database is readable.
        $StandBy = "0",
                       
        #The WithRecovery parameter allows overriding WITH RECOVERY, default is NORECOVERY .
        $WithRecovery = "0",
                       
        #The WithCHECKDB parameter executes CHECKDB, only possible in conjunction WITH RECOVERY .
        $WithCHECKDB = "0",
                       
        #The LogShippingInitialize parameter performs a full, diff, log(s) recovery if 1, else outstanding logds only if 1
        $LogShippingInitialize = "1",
                       
        # Removed in V6.57, causes issues in edge cases
        #The Log_Reference parameter value is written to the SQL Error Log
        #$Log_Reference = $null,
                       
        #The KillConnections parameter will kill restore blocking SPID's if 1
        $KillConnections = "1",
                       
        #If run interactively, change to "0" for SQL Agent Jobs.
        $ConsoleFeedback = "1",
                 
        # Credentials for Azure Blog Storage
        $BlobCredential = $null,
                 
        # RestoreScript String Find
        $RestoreScriptReplaceThis = $null,
                 
        # Restore Script String Replace
        $RestoreScriptWithThis = $null,
                
        # Log restore commands, do not execute
        $DebugMode = $null,
    
        # Drop the database after the restore
        $DropDatabaseAfterRestore = $null,
    
        # Drop the database after the restore
        $ExcludeDiffAndLogBackups = $null,
    
        # Include backup device type "7"
        $IncludeDeviceType7 = "1",
    
        # Include backup device type "102"
        $IncludeDeviceType102 = "1",
    
        # Include backup device type "2"
        $IncludeDeviceType2 = "1",
    
        # Include backup device type "9"
        $IncludeDeviceType9 = "1"  
                       
   )
                       
# Defaults and Overrides
if ($FromFileDiffUNC -eq $null) {$FromFileDiffUNC = $FromFileFullUNC}
if ($FromFileLogUNC-eq $null) {$FromFileLogUNC= $FromFileFullUNC}
if ($StopAt -eq $null) {$StopAt = Get-Date -Format s}
                
# ==============================================================================
# Open a connection to the primary server
$SQLConnectionPrimary = New-Object System.Data.SqlClient.SqlConnection
try
{
    $SQLConnectionPrimary.ConnectionString = "Server=" + $PrimaryServer + ";Database=master;Integrated Security=True"
    $SQLConnectionPrimary.Open()
}
catch
{
    throw "Error : Connection to Primary server cannot be established"
}
                       
# ==============================================================================
# Open a connection to the standby server,to execute RESTORE commands
$SQLConnectionStandby = New-Object System.Data.SqlClient.SqlConnection
try
{
    $SQLConnectionStandby.ConnectionString = "Server=" + $StandbyServer + ";Database=master;Integrated Security=True"
    $SQLConnectionStandby.Open()
}
catch
{
    throw "Error : Connection to Standby server cannot be established"
}
                       
# ==============================================================================
# Check for connections blocking the restore on the standby server database
if ($KillConnections -eq 0)
{
  if ($TargetDBName -eq $null) {$activeconnections = "SELECT * FROM sys.sysprocesses WHERE dbid = DB_ID('" + $DBName + "')"}
  elseif ($TargetDBName -ne $null) {$activeconnections = "SELECT * FROM sys.sysprocesses WHERE dbid = DB_ID('" + $TargetDBName + "')"}
      
  $check = Invoke-Sqlcmd -Query $activeconnections -Database "master" -ServerInstance $StandbyServer
  if ($check.spid -ne $null)
  {throw "Error : Active connections to the database are blocking the restore on the standby server"}
}
                       
# ==============================================================================
#  Function to kill restore blocking SPID's on standby
function f_killconnections
{
  if ($TargetDBName -eq $null) {$KillQueryConstructor = "SELECT ';KILL ' + CAST(spid AS VARCHAR(4)) + '' FROM sys.sysprocesses WHERE spid > 50 AND dbid = DB_ID('" + $DBName + "')"}
  elseif ($TargetDBName -ne $null) {$KillQueryConstructor = "SELECT ';KILL ' + CAST(spid AS VARCHAR(4)) + '' FROM sys.sysprocesses WHERE spid > 50 AND dbid = DB_ID('" + $TargetDBName + "')"}
      
  $KillCommands = Invoke-Sqlcmd  -QueryTimeout 6000 -Query $KillQueryConstructor -Database "master" -ServerInstance $StandbyServer
  foreach ($KillCommand in $KillCommands)
  {
    if ($KillCommand -ne $NULL)
    {
      $KillCommand[0] | Out-Default
      $Result = Invoke-Sqlcmd -Query $KillCommand[0] -Database "master" -ServerInstance $StandbyServer
      $Result | Out-Default
    }
  }
}        
                       
# ==============================================================================
# Get Previous Restore Log details
$StartTime = Get-Date -Format s
                       
if ($LogShippingInitialize -eq "1") {$LogShippingStartTime = $StartTime}
                       
if ($LogShippingInitialize -eq "0")
{
  [xml]$PreviousLog = Get-Content -Path $RestoreLog
  $LogShippingStartTime = $PreviousLog.LogData.LogShippingStartTime.Description
  $LogShippingLastLSN = $PreviousLog.LogData.LogShippingLastLSN.Description
  $PreviousWithRecovery = $PreviousLog.LogData.WithRecovery.Description
  if ($LogShippingLastLSN -eq $null) {throw "Error : Previous log file LastLSN is invalid or not found"}
  if ($PreviousWithRecovery -eq "1") {throw "Error : WITH RECOVERY has been run on the standby server"}
}
                       
# ==============================================================================
# Restore Log File processing, record start runtime details
# Create a new XML File with  root node
[System.XML.XMLDocument]$oXMLDocument=New-Object System.XML.XMLDocument 
                       
# New Node
[System.XML.XMLElement]$oXMLRoot=$oXMLDocument.CreateElement("LogData") 
                       
# Append as child to an existing node
$oXMLDocument.appendChild($oXMLRoot) 
                       
# Add a Attribute
$oXMLRoot.SetAttribute("description","ps_LogShippingLight") 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("StartTime"))
$oXMLSystem.SetAttribute("Description",$StartTime) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("WithRecovery"))
$oXMLSystem.SetAttribute("Description",$WithRecovery) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("DBName"))
$oXMLSystem.SetAttribute("Description",$DBName) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("TargetDBName"))
$oXMLSystem.SetAttribute("Description",$TargetDBName) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("WithMoveDataFiles"))
$oXMLSystem.SetAttribute("Description",$WithMoveDataFiles) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("WithMoveLogFiles"))
$oXMLSystem.SetAttribute("Description",$WithMoveLogFile) 
             
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("WithMoveFileStreamFile"))
$oXMLSystem.SetAttribute("Description",$WithMoveFileStreamFile) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("FromFileFullUNC"))
$oXMLSystem.SetAttribute("Description",$FromFileFullUNC) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("FromFileDiffUNC"))
$oXMLSystem.SetAttribute("Description",$FromFileDiffUNC) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("FromFileLogUNC"))
$oXMLSystem.SetAttribute("Description",$FromFileLogUNC) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("StopAt"))
$oXMLSystem.SetAttribute("Description",$StopAt) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("LogShippingStartTime"))
$oXMLSystem.SetAttribute("Description",$LogShippingStartTime) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("Standby"))
$oXMLSystem.SetAttribute("Description",$StandBy) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("WithCHECKDB"))
$oXMLSystem.SetAttribute("Description",$WithCHECKDB) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("PrimaryServer"))
$oXMLSystem.SetAttribute("Description",$PrimaryServer) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("StandbyServer"))
$oXMLSystem.SetAttribute("Description",$StandbyServer) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("RestoreLog"))
$oXMLSystem.SetAttribute("Description",$RestoreLog) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("LogShippingInitialize"))
$oXMLSystem.SetAttribute("Description",$LogShippingInitialize ) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("KillConnections"))
$oXMLSystem.SetAttribute("Description",$KillConnections) 
                 
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("BlobCredential"))
$oXMLSystem.SetAttribute("Description",$BlobCredential) 
                 
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("RestoreScriptReplaceThis"))
$oXMLSystem.SetAttribute("Description",$RestoreScriptReplaceThis) 
                 
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("RestoreScriptWithThis"))
$oXMLSystem.SetAttribute("Description",$RestoreScriptWithThis)
                
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("DebugLog"))
$oXMLSystem.SetAttribute("Description",$DebugLog)
    
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("DropDatabaseAfterRestore"))
$oXMLSystem.SetAttribute("Description",$DropDatabaseAfterRestore)
    
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("ExcludeDiffAndLogBackups"))
$oXMLSystem.SetAttribute("Description",$ExcludeDiffAndLogBackups)             
    
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("IncludeDeviceType7"))
$oXMLSystem.SetAttribute("Description",$IncludeDeviceType7) 
    
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("IncludeDeviceType102"))
$oXMLSystem.SetAttribute("Description",$IncludeDeviceType102) 
    
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("IncludeDeviceType2"))
$oXMLSystem.SetAttribute("Description",$IncludeDeviceType2) 
    
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("IncludeDeviceType9"))
$oXMLSystem.SetAttribute("Description",$IncludeDeviceType9) 
    
                       
# ==============================================================================
#Snapin for the Invoke-SQLCmd cmdlet
Add-PSSnapin SqlServerCmdletSnapin100 -ErrorAction SilentlyContinue
 
# Thomas Wolf, December 2020
#Import-Module SQLPS -DisableNameChecking
 if (!(Get-Module -Name "SqlServer")) 
{
  Import-Module SQLPS -DisableNameChecking
}
 
 
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Set-Location -Path $dir
                       
# ==============================================================================
# Construct the Restore Gene stored procedure call
                
if ($LogShippingInitialize  -eq "0") 
{
  $LogShippingLight_EXEC = "EXEC dbo.sp_RestoreGene @LogShippingLastLSN = '" + $LogShippingLastLSN + "', @LogShippingStartTime= '" + $LogShippingStartTime + "',@Database = '" + $DBName + "', @TargetDatabase = '" + $TargetDBName + "',@WithMoveDataFiles = '" + $WithMoveDataFiles + "', @WithMoveLogFile = '" + $WithMoveLogFile + "', @WithMoveFileStreamFile = '" + $WithMoveFileStreamFile + "', @FromFileFullUNC = '" + $FromFileFullUNC + "', @FromFileDiffUNC = '" + $FromFileDiffUNC + "', @FromFileLogUNC= '" + $FromFileLogUNC+ "', @StopAt = '" + $StopAt + "', @StandbyMode = '" + $Standby + "', @WithReplace = '" + $WithReplace + "' , @WithRecovery = '" + $WithRecovery + "', @WithCHECKDB = '" + $WithCHECKDB + "', @DropDatabaseAfterRestore = '" + $DropDatabaseAfterRestore + "' , @BlobCredential = '" + $BlobCredential + "', @RestoreScriptReplaceThis = '" + $RestoreScriptReplaceThis + "', @RestoreScriptWithThis = '" + $RestoreScriptWithThis + "'" + ", @SetSingleUser = '" + $KillConnections + "'"  + ", @ExcludeDiffAndLogBackups = '" + $ExcludeDiffAndLogBackups + "'"  + ", @IncludeDeviceType7 = '" + $IncludeDeviceType7 + "'"  + ", @IncludeDeviceType102 = '" + $IncludeDeviceType102 + "'"  + ", @IncludeDeviceType2 = '" + $IncludeDeviceType2 + "'"  + ", @IncludeDeviceType9 = '" + $IncludeDeviceType9 + "'"
}
else
{
  $LogShippingLight_EXEC = "EXEC dbo.sp_RestoreGene @Database = '" + $DBName + "', @TargetDatabase = '" + $TargetDBName + "',@WithMoveDataFiles = '" + $WithMoveDataFiles + "', @WithMoveLogFile = '" + $WithMoveLogFile + "', @WithMoveFileStreamFile = '" + $WithMoveFileStreamFile + "', @FromFileFullUNC = '" + $FromFileFullUNC + "', @FromFileDiffUNC = '" + $FromFileDiffUNC + "', @FromFileLogUNC= '" + $FromFileLogUNC+ "', @StopAt = '" + $StopAt + "', @StandbyMode = '" + $Standby + "', @WithReplace = '" + $WithReplace + "' , @WithRecovery = '" + $WithRecovery + "', @WithCHECKDB = '" + $WithCHECKDB + "', @DropDatabaseAfterRestore = '" + $DropDatabaseAfterRestore + "' , @BlobCredential = '" + $BlobCredential + "', @RestoreScriptReplaceThis = '" + $RestoreScriptReplaceThis + "', @RestoreScriptWithThis = '" + $RestoreScriptWithThis + "'" + ", @SetSingleUser = '" + $KillConnections + "'"  + ", @ExcludeDiffAndLogBackups = '" + $ExcludeDiffAndLogBackups + "'"  + ", @IncludeDeviceType7 = '" + $IncludeDeviceType7 + "'"  + ", @IncludeDeviceType102 = '" + $IncludeDeviceType102 + "'"  + ", @IncludeDeviceType2 = '" + $IncludeDeviceType2 + "'"  + ", @IncludeDeviceType9 = '" + $IncludeDeviceType9 + "'"
}
                    
if ($ConsoleFeedback -eq "1") {
  Write-Host "-----------------------------------------"
  Write-Host "RestoreGene Batch Execution Starting"                
  Write-Host "-----------------------------------------"
  Write-Host $LogShippingLight_EXEC
}
                
# ==============================================================================
# Execute the sp_RestoreGene stored procedure on the primary server
try
{
    $LogShippingLight_Results = Invoke-SQLCmd -Query $LogShippingLight_EXEC -QueryTimeout 6000 -Database "master" -ServerInstance $PrimaryServer
}
catch
{             
    if ($LogShippingLight_Results -eq $null)
    {
        #throw "Error : No backups files pending restore"
        if ($ConsoleFeedback -eq "1") 
        {
            Write-Host "-----------------------------------------"
            Write-Host "No new backups found, nothing to restore"
        }   
    }
   
    Write-Host "-----------------------------------------"
    Write-Host "Restore Gene Encountered the Error Below "
    Write-Host "-----------------------------------------"
    Write-Host $Error
    Write-Host "-----------------------------------------"
    Throw "Error in Restore Gene shown above"
}
                       
# Save correctly sequenced restore commands to a hash table
$hash=@{}
foreach ($command in $LogShippingLight_Results)
{
  $hash.Add($command.SortSequence,($command.TSQL))
}
$pendingcmds = $hash.GetEnumerator() | Sort-Object -Property key
            
# Save highest LSN for Restore Log file
foreach ($LSN in $LogShippingLight_Results)
{
  if ($LogShippingLastLSN -lt $LSN.Last_LSN) {$LogShippingLastLSN = $LSN.Last_LSN}
}
                
if ($pendingcmds.Value -eq $null) {
      Write host"-----------------------------------------------------".
      Write-Host "No Backup Files to restore."
      Write host"-----------------------------------------------------".
      Throw "No Backup Files to restore."
}
  
                       
# ==============================================================================
# Execute the RESTORE commands on the standby server
foreach ($pendingcmd in $pendingcmds)
{
  if ($KillConnections -eq 1) {f_killconnections}
    
  # Removed in V6.57
  #$DBName_nospaces = $DBName -replace(" ","_")                  # doesn't like spaces 
  #$DBName_nospaces = $DBName_nospaces -replace("\.","_")            # doesn't like full stops
  #$DBName_nospaces = $DBName_nospaces -replace("-","_")            # doesn't like hyphens
  #$cmd = ";DECLARE @msg_" + $DBName_nospaces + " VARCHAR(1000) " + $pendingcmd.value
       
                       
  if ($ConsoleFeedback -eq "1") 
  {
    Write-Host "-----------------------------------------"
    Write-Host $pendingcmd.value
  }
                       
  try
  {
    if ($DebugMode -eq "1") 
    {
        $DebugLog = $DebugLog + $pendingcmd.value
    }
    else
    {
        $SQLCommand = New-Object System.Data.SqlClient.SqlCommand($pendingcmd.value, $SQLConnectionStandby)
        $SQLCommand.CommandTimeout=65535
        $SQLCommand.ExecuteScalar()
    }
  }
  catch
  {
      Write-Host "-----------------------------------------"
      Write-Host "Restore Gene Encountered the Error Below "
      Write-Host "-----------------------------------------"
      Write-Host $Error
      Write-Host "-----------------------------------------"
      Throw "Error in Restore Gene shown above"
  }                    
  sleep -Seconds 1
}
                       
if ($ConsoleFeedback -eq "1") {
  Write-Host "-----------------------------------------"
  Write-Host "RestoreGene Batch Execution Complete"
  Write-Host "-----------------------------------------"
   
}
                       
# ==============================================================================
# Restore Log File final processing, record completion details
$EndTime = Get-Date -Format s
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("EndTime"))
$oXMLSystem.SetAttribute("Description",$EndTime) 
                       
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("LogShippingLastLSN"))
$oXMLSystem.SetAttribute("Description",$LogShippingLastLSN)
                
[System.XML.XMLElement]$oXMLSystem=$oXMLRoot.appendChild($oXMLDocument.CreateElement("RestoreScript"))
$oXMLSystem.SetAttribute("DebugLog",$DebugLog)
                       
$oXMLDocument.Save($RestoreLog) 