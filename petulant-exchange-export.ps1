param(  [string]$logpath = "C:\Exchange Exports\ExportLog.csv",            #Parameter to manually specify log file path and filename.
        [switch]$quiet,                                                    #Run without outputting messages to the screen, Useful as debug messages screw up pipeline. 
        [int]$BatchSize = 4,                                               #Manually Specify Batch Size, Default 4
        [string]$ExchangeExportShare = "C:\Exchange Exports",              #Manually specify export location. Exchange Trusted Subsystem must have permissions to this NTFS fileshare.
        [string]$EventLogSource = "ExchangeMailboxExport",                 #Name for the script event log. 
        [switch]$RemovePSTBeforeExport = $false,                           #Remove any existing PST files before export? 
        [switch]$SkipExistingPST = $true,                                  #Check Directory for existing PST Files.
        [int]$ErrorTolerancePerMailbox = 50,                               #Tolerance for corruption in Mailboxes. Max 50
        [string]$SMTPFromAddress = "test@test.org",                        #From Address for Logging Messages
        [string]$SMTPToAddress = "test@test.org",                          #To Address for Logging Messages
        [string]$SMTPServer = "smtp.test.org"                                   #SMTP Server for Logging Messages
)


#Load Exchange PS Snapin and begin an Exchange Powershell Session. 
add-pssnapin Microsoft.Exchange.Management.PowerShell.Snapin
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1' Connect-ExchangeServer -auto -ManagementShell | Out-Null

####################################################################################################
#
# petulant-exchange-export
#
# Requirements before running. 
#
# Needs to be run from an elevated Powershell session, in the context of a 
# service account with local admin rights on the host machine, and the ability to send as 
# it.support@rslqld.org. This service account needs to also have the Exchange RBAC role Import Export.
#
# New-ManagementRoleAssignment -Role "Mailbox Import Export" -User "Username"
#
# A network share must also be created with full access rights granted to
# "Exchange Trusted Subsystem" this is the share that all the PST files
# will be exported to. If you need to change or move the share, ensure that the new
# share has these access rights. 
#
# Skip existing PST is not a safe resume, if for whatever reason the last export did not complete
# completely (server shutdown, significant corruption) recommend setting $RemovePSTBeforeExport to $True
#
# Common Parameters
#
# -logpath <directory>             Manually specify log output directory
#
# -quiet                           Switch, suppresses most console output
#
# -BatchSize <integer>             Manually specify how large the mailbox batch is. 
#
# -ExchangeExportShare <directory> Manually specify the Exchange Export Share
#
# -EventLogSource <Name>           Override the default name for events in the Event Log
#
# -RemovePSTBeforeExport           Switch, Removes contents of Exchange Export Share before export.
#
####################################################################################################


function main()
{

#Check to see if anyone else is exporting, and break to prevent any cross contamination. 
#check-currentexports

#Get all mailboxes to export. 

$MailboxList = Get-Mailbox

#Create System.Collections Objects.
$BatchMailboxes = New-Object System.Collections.Stack

$TestMailboxes = New-Object System.Collections.ArrayList

$MailboxesToExport = New-Object System.Collections.Stack

        #Remove PST files before export if required.
        if($RemovePSTBeforeExport -eq $true)
        {
            Remove-PST
        }
       
       
        #Grab a list of all files in Export Share
        $ExistingPST = Get-ChildItem $ExchangeExportShare -Recurse

        #Add mailboxes to Stack
        foreach ($Mailbox in $MailboxList)
        {
            
            #Extra Logic to restore from next pst
            if($SkipExistingPST)
                {
                    
                    $SkipPST = $false
                        
                        #Grab a name from an exported PST file, check it against all files in target director.
                        foreach ($PST in $ExistingPST)
                            {
                                $TempName = $PST.name
                                $TempName = $TempName.Split(".")
                                if($Mailbox.name -eq $TempName[0])
                                    {
                                        $SkipPST = $true
                                    }
                            }

                            if($SkipPST -eq $True)
                                {
                                    #PST Exists, Log and Skip Export
                                    log-export "$Mailbox" "Export Skipped, PST Exists."   
                                }
                                else
                                {
                                    #Push mailbox to be exported
                                    $MailboxesToExport.Push($Mailbox.Name)
                                }
                }
            else
                {
                    #Push Mailbox to be exported
                    $MailboxesToExport.Push($Mailbox.Name)

                }

        }
    $MailboxesToExport
       
       #While there are still mailboxes to export, continue. 
       While($MailboxesToExport.Count -gt 0)
       {
            #Create a batch.
            Create-Batch
           
            #Export that batch.
            Export-Batch
            
            #Check that the batch has completed.
            Check-Batch
       }
      

        if($MailboxesToExport.Count -eq 0)
        {
            Send-EmailMessage $SMTPToAddress "All Mailboxes Exported, See log for details"
        }
        else
        {
            log-export "Script Running" "Error. MailboxesToExport Less than Zero"
        }
}

##############################################################################
#.SYNOPSIS
# Remove all items in the Export share directory
#
#.DESCRIPTION
# Recursively removes all items in the export share directory
#
#.EXAMPLE
# Remove-PST
# 
##############################################################################
function Remove-PST()
{
    #Remove all items in the Export share directory
    Get-ChildItem $ExchangeExportShare -Recurse | Remove-Item
}

##############################################################################
#.SYNOPSIS
# Creates a new batch of Mailboxes. 
#
#.DESCRIPTION
# Creates a Batch of Mailboxes, by popping $BatchSize Mailboxes off of $MailboxesToExport
#
#.EXAMPLE
# Create-Batch
# 
##############################################################################
function Create-Batch()
{
    #Create Enumerator and clear System.Collections Objects
    $BatchEnumerator = 0
    $BatchMailboxes.Clear()
    $TestMailboxes.Clear()

    #If the number of mailboxes is 0 or less, break
    if ($MailboxesToExport.Count -lt 1)
    {
        log-export "Complete" "All Mailboxes Exported"
        break
    }
    else
    {
        #If the next batch will be the correct batch size.
        if ($MailboxestoExport.Count -ge $BatchSize)
        {
                #Create a batch, which includes a Stack and an Array list to test. 
               while($BatchEnumerator -lt $BatchSize)
               {
                    $NameTemp = $MailboxesToExport.Pop()
                    $BatchMailboxes.push($NameTemp)
                    $TestMailboxes.add($NameTemp)
                    $BatchEnumerator++    
   
                }  
        }        
        else #for batches less than full size. 
        {
            #Create a batch, which includes a Stack and an Array list to test. 
             while($BatchEnumerator -lt $MailboxesToExport.Count)
               {
                    $NameTemp = $MailboxesToExport.Pop()
                    $BatchMailboxes.push($NameTemp)
                    $TestMailboxes.add($NameTemp)
                    $BatchEnumerator++  
    
               }  

        }
    }
 }

##############################################################################
#.SYNOPSIS
# Exports a new batch of Mailboxes. 
#
#.DESCRIPTION
# Iterates through the current batch and creates a new mailbox export request
# for each mailbox in the batch. 
#
#.EXAMPLE
# Export-Batch
# 
##############################################################################
 function Export-Batch()
 {
  #Zeroise the Enumerator
  $CheckEnumerator = 0
        
        
        while($CheckEnumerator -lt $BatchSize)
        {

            if($BatchMailboxes.Count -gt 0)
            {
                #Pop off the next Mailbox from this batch
                $MBOXNAME = $BatchMailboxes.pop() 

                #Export the mailbox to the mailboxname.pst
                        New-MailboxExportRequest -Mailbox $MBOXNAME -Filepath "$ExchangeExportShare\xx$MBOXNAME.pst" #-ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable ExportRequestError
                
                   #If error tolerance is required, set the export to allow it.
                    if($ErrorTolerancePerMailbox -gt 0)
                        {
                            Get-MailboxExportRequest | Set-MailboxExportRequest -BadItemLimit $ErrorTolerancePerMailbox -AcceptLargeDataLoss
                        }

                        #Increment Enumerator
                        $CheckEnumerator++

                                #Handle Errors
                                if ($ExportRequestError)
                                    {
                                        #Parse the whole error log, and export it to required logging sources
                                        ParseErrors $ExportRequestError

                                        #Log the fact that errors have been encountered during export.
                                        log-export "Ended with Errors" "Ending With Errors: Error Exporting Mailboxes"

                                        #End the script prematurely. 
                                        break
                                    }
            }
       }
 
 }

##############################################################################
#.SYNOPSIS
# Checks the current batch versus a fresh output to see if the batch is complete
#
#.DESCRIPTION
# Creates a fresh Get-MailboxExport Request. Iterates through all current batch 
# mailboxes and checks each for completion. 
#
#.EXAMPLE
# Check-Batch
# 
##############################################################################
 function Check-Batch()
 {
    
            while ($TestMailboxes.count -gt 0)                                  
            {
                sleep -Seconds 120

                #Request fresh export info from the exchange server
                $MailboxExportText = Get-MailboxExportRequest -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable GetMailboxErrors

               
                foreach ($Entry in $MailboxExportText)
                {
                                   
                        #Convert the current AD Entry to a string
                        $ComparisonText = "$Entry.Mailbox"
                    
                        #Split the name from the mailbox export type
                        $ComparisonString = $ComparisonText.Split("\")

                            #If the entry is complete, check each name
                            if($Entry.Status -contains "Completed")
                            {
                                    
                                    foreach ($Mailbox in $TestMailboxes.ToArray())
                                     {

                                        if($ComparisonString[0].endswith($Mailbox))
                                            { 
                                                $TestMailboxes.Remove($Mailbox)
                                                $Entry | Remove-MailboxExportRequest -Confirm: $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable RemoveMailboxErrors
                                                log-export "$ComparisonString[0]"
                                            }
                        
                                     }
             
             
                            } 
                            if($Entry.Status -contains "Failed")
                            {
                                    
                                    foreach ($Mailbox in $TestMailboxes.ToArray())
                                     {

                                        if($ComparisonString[0].endswith($Mailbox))
                                            { 
                                                $TestMailboxes.Remove($Mailbox)
                                                $Entry | Remove-MailboxExportRequest -Confirm: $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable RemoveMailboxErrors
                                                Send-EmailMessage $SMTPToAddress "$Mailbox Export Failed, Please review logs for details"
                                                log-export "$ComparisonString[0]" "Failed"
                                            }
                        
                                     }
             
             
                            }             

                }
                    #Check for Errors
                    if ($GetMailboxErrors)
                    {
                        #Parse errors to log, and break
                        ParseErrors($GetMailboxErrors)
                        log-Export "ExportFailure" "Could Not Generate Export, Ending"
                        Send-EmailMessage $SMTPToAddress "Could Not Generate Export, Ending"
                        break
                    }
                    if ($RemoveMailboxErrors)
                    {
                        #Parse errors to log, and break
                        ParseErrors($RemoveMailboxErrors)
                        log-Export "ExportFailure" "Could Not Remove Export, Ending"
                        Send-EmailMessage $SMTPToAddress "Could Not Generate Export, Ending"
                        break
                    }
            }
                  
 }
##############################################################################
#.SYNOPSIS
# Logs data in all required locations
#
#.DESCRIPTION
# Takes a Name and optionally an Error Message and creates log output for this data.
#
#.PARAMETER $fName
# Name of Log Entry
#
#.PARAMETER $errMessage
# Optional. Default = OK. Error Message Text to pass to log outputs. 
#
#.EXAMPLE
# log-export "administrator@test.org"
# 
##############################################################################
 function log-export([string]$fName, [string]$errMessage = "OK")
 {
        
        #Get Current Date
        $logdate = Get-Date

        #Ensure that the Event Log exists.
        New-EventLog –LogName Application –Source $EventLogSource -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
       
        #Parse Inputs to a PSObject
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty "Date"($logdate)
        $obj | Add-Member NoteProperty "Mailbox"($fName)
        $obj | Add-Member NoteProperty "Status"($errMessage)
        
        if (!($quiet))
            {   
            #Write output to console, only if verbose. 
            Write-Output $obj
            }
        
        #Write to dedicated logfile         
        
            #Write to the Event Log
            if($errMessage -eq "OK")
            {
                Write-EventLog –LogName Application –Source $EventLogSource –EntryType Information –EventID 1 –Message "$fName Mailbox Exported $errMessage on $logdate”
            } 
            else
            {
                Write-EventLog –LogName Application –Source $EventLogSource –EntryType Information –EventID 1 –Message "$fName Error: $errMessage on $logdate”
            }  
 }
 
 
##############################################################################
#.SYNOPSIS
# Checks Exchange Server for current exports.
# 
#
#.DESCRIPTION
# Checks for previous Mailbox Export Requests, and backs off if any are found.
# This includes completed tasks, as these records may have not been checked by those
# who created them. To preserve this data, the script breaks.
#
#.EXAMPLE
#check-currentexports
#
#
##############################################################################
 function check-currentexports()
 {
    #Get a fresh export request
    $MailboxExportRequestCheck = Get-MailboxExportRequest
        
        
        if($MailboxExportRequestCheck)
            {
                #If there are existing exports, break the script to prevent interrupt and log.
                log-export "ScriptStartup" "Failed Due to Existing Export Requests. Please Clean these up and try again."
                Send-EmailMessage $SMTPToAddress "Failed Due to Existing Export Requests. Please Clean these up and try again."
                break
            }
            else
            {
                #Log that the exchange server is clear and that the 
                log-export "ScriptStartup" "Mailbox Exports are clear, beginning export of all mailboxes."
            }

 }
##############################################################################
#.SYNOPSIS
# Send a quick email message
# 
#
#.DESCRIPTION
# Creates an email message from the parameters and sends it via the assigned 
# SMTP server.
#
#.EXAMPLE
# Send-EmailMessage "matt.osborne@test.org" "What are the haps my friend" 
#
#
##############################################################################
 function Send-EmailMessage([string]$EmailAddress,[string]$MessageBody, [string]$SMTPServerAddress = $SMTPFromAddress)
 {
        #Take values, send email message. 
        send-mailmessage -to $EmailAddress -from $SMTPFromAddress -Subject "Important Message from Exchange Export Script" -body $MessageBody  -smtpserver $SMTPServerAddress -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -ErrorVariable EmailError
    
        #Parse any errors with sending email. 
        ParseErrors($EmailError)
 
 }


##############################################################################
#.SYNOPSIS
# Takes a list of errors created by -ErrorVariable and parses it for useful 
# information.
#
#.DESCRIPTION
# Checks $ErrorList for PathTooLong and SystemAccess Exceptions, based on that 
# it determines whether to include or ignore these errors, and forwards them 
# on to the relvant output function if required.
#
#.PARAMETER $ErrorList
# List of errors to parse
#
#.EXAMPLE
# CheckErrors $ErrorList
# 
##############################################################################
function ParseErrors ([PSObject] $ErrorList)
{

 # Check to see if there have been any errors
    if(($ErrorList))
    {
               
        #Test logging
        Log-Export $logPath "Errors Have Occurred" -force –ErrorAction SilentlyContinue –ErrorVariable logFail;
        
            #Test to see that logging did not produce any errors
            if (($logFail))
            {
                #Request a new logging directory
                (Write-Host "Logging Failed, please specify a working directory for log file")
                break
            }
    
                #Check each error message for report worthy instances, dump all errors to logfile
                $ErrorList | ForEach-Object{$obj = New-Object PSObject
                       
                        log-export "Unknown"  $_.Exception.GetType().FullName 
                        
                        #Append error to Log File and Event Log
                        }

            # Inform user of errors 
            if(!($quiet))
            {
                Write-Host "There were errors, please view the log file at $logPath for more information" -foregroundcolor "red"
            }
    }
    else
    {
        if(!($quiet))
        {
            Write-Host "No Errors" -foregroundcolor "green"
        }
    }
}

 main   