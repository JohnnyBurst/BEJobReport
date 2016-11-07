<#
    .SYNOPSIS    
        
    This script will connect to Symantec BackupExec server and using PowerShell, it will collect the following  information about the backup jobs:

        -Job Name
        -Selection Summary
        -Storage
        -Start Time
        -Elapsed Time
        -Job Status
        -Media Label
        -Total Data Size Bytes
        -Job Rate MB Per Minute
        -Script Filters

    The script has two filters :

        -Name Like Expression: You can say for example give me jobs where the name contains "*OffSite*"
        -Time Expression: You can either use:
            -Since X Days : like give me all jobs happened since X days
            -From Last Job Run: give me the last job run information for the jobs
        
    --------------
    Script Info
    --------------

        Script Name                :         BackupExec Info (Get-BEJOBs)
        Script Version             :         1.1
        Author                     :         Ammar Hasayen   
        Blog                       :         http://ammarhasayen.com         
        Twitter                    :         @ammarhasayen
        Email                      :         me@ammarhasayen.com

    --------------
    Copy Rights
    --------------

        
     .LINK
     My Blog
     http://ammarhasayen.com

    

     .DESCRIPTION 
      
        The script will generate three log files:

            -Info Log : will help you track what the script is doing.
            -Error Log : in case of errors
            -Detailed Log : contains detailed information about each job run

        The script will also generate a nice HTML table that contains the list of jobs and their information

        The script has an option to send these info via an email if you choose to configure SMTP settings via one of the script parameters.
        
         The challenge i faced writing this script is querying the Media Label field because this field is represented by XML file returned from the Get-JobLog.

        So if you write Get-Job |Get-JobHistory |GetJobLog

        then you will have XML file with the media label information there. I had to do some string operations to extract the media label information.

        The script should be running from within the BackupExec server and it is tested with BackupExec 2014 only.    
        
          

     .PARAMETER ScriptFilesPath
     Path to store script files like ".\" to indicate current directory or full path like C:\myfiles

	.PARAMETER SendMail
	 Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
	
	.PARAMETER MailFrom
	 Email address to send from.
	
	.PARAMETER MailTo
	 Email address to send to.
	
	.PARAMETER NameLike 
	 Expression filter to filter on Job names. Example is "*OffSite*" to filter for any job with the word "OffSite" in the job name.

    .PARAMETER Days
     Filter jobs happening in the past X days.This parameter cannot be used with the -FromLastJobRun switch parameter

     .PARAMETER FromLastJobRun
     Switch parameter. When used, the script will bring only the last run job instances. This switch parameter cannot be used with the (-Days) parameter.



    .EXAMPLE
     Get BackupExec Jobs and send email with the results
     \Get-BEJobs.ps1 -ScriptFilesPath .\  -SendMail:$true -MailFrom noreply@contoso.com  -MailTo me@contoso.com  -MailServer smtp.contoso.com

     .EXAMPLE
     Get BackupExec Jobs happening in the last 3 days
     \Get-BEJobs.ps1 -ScriptFilesPath .\  -Days 3

     .EXAMPLE
     Get BackupExec Jobs with the name containing "*Yearly*"
     \Get-BEJobs.ps1 -ScriptFilesPath .\  -NameLike "*Yearly*"

     .EXAMPLE
     Get BackupExec Jobs with the name containing "*Yearly*" and only returning the last job run results
     \Get-BEJobs.ps1 -ScriptFilesPath .\  -NameLike "*Yearly*" -FromLastJobRun

     .EXAMPLE
     Get BackupExec Jobs with the name containing "*Yearly*" happening last week
     \Get-BEJobs.ps1 -ScriptFilesPath .\  -NameLike "*Yearly*" -Days 7


   
   #>

#region parameters

    [cmdletbinding()]

    param(
        [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Path to store script files like c:\ ')][string]$ScriptFilesPath = ".\",
        [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false)][string]$NameLike = "*",
        [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false)][string]$Days = 3,
        [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,ParameterSetName="FromLastJobRun")][Switch]$FromLastJobRun,	
        [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Send Mail ($True/$False)')][bool]$SendMail=$false,
	    [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail From')][string]$MailFrom,
	    [parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail To')]$MailTo,
	    [parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Server')][string]$MailServer	    
        )
     

#endregion parameters


#region Module Functions

    #region helper functions

        function _sendEmail {
     
           param($from,$to,$subject,$smtphost,$ScriptFilesPath,$InfoFullPath,$ErrorFullPath) 

               
                $varerror = $ErrorActionPreference
                $ErrorActionPreference = "Stop"

                #region files inventory
                    
                    $htmlFileName = Join-Path $ScriptFilesPath "BEJobHTML.html"

                    $files = @()

                    foreach ($file in (Get-ChildItem $ScriptFilesPath)) {
                        if ($file.name -like "*detail*" ) {
                            $files += $file.Name
                        }
                    } 

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Directory to look at attachment $ScriptFilesPath"
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Files to attached count = $($files.count)"

                #endregion files inventory


                #region prepare SMTP Client

                    $msg = new-object Net.Mail.MailMessage            
                    $msg.From = $from
                    $msg.To.Add($to)
                    $msg.Subject = $subject
                    $msg.Body = Get-Content $htmlFileName 
                    $msg.isBodyhtml = $true 

                #endregion prepare SMTP Client

                
                #region add attachments

                    #protection threshold
                    if ($files.count) {
                        if($files.count -lt 7) {
                            $varattach = $files.count
                        }else {
                            $varattach = 7
                        }

                        for ($i=0; $i -lt $varattach; $i++) {

                            try {

                                $file = $files[$i]
                                $file = Join-Path $ScriptFilesPath $file
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - File  = $file "
                                $attachment = new-object Net.Mail.Attachment($file)
                                $msg.Attachments.Add($attachment)

                            }catch {
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - Error attaching File : $file "
                                Write-CorpError -myError $_  -Info "[Module Closure - Fail attaching file" -mypath $ErrorFullPath
                            }
                        }
                    } else {
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - No files to attach "
                        
                    }

                #endregion add attachments


                #region send email
                    
                    try{
                        $smtp = new-object Net.Mail.SmtpClient($smtphost)
                        $smtp.Send($msg)
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Email Sent to $to"
                    }catch {
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Fail to send email... check Error log"
                        Write-CorpError -myError $_  -Info "[Module Closure - Could not send email" -mypath $ErrorFullPath
                        _status "      Could not send email... check Info log for detials" 2 
                    }finally {$msg.dispose()}

                #endregion send email


                $ErrorActionPreference = $varerror 
                
               

        }  # function _sendEmail

        
        function Get-CorpModule {

            <#.Synopsis
            Imports a module if not imported.

            .DESCRIPTION
            Imports a module if not imported.



            Versioning:
                - Version 1.0 written 24  November 2013 : First version
    



            .PARAMETER Name
            string representing module name

            .EXAMPLE
            PS C:\> Get-CorpModule  -name ActiveDirectory


            .Notes
            Last Updated             : Nov 24, 2013
            Version                  : 1.0 
            Author                   : Ammar Hasayen (Twitter @ammarhasayen)
            Email                    : me@ammarhasayen.com
            based on                 : 

            .Link
            http://ammarhasayen.com


            .INPUTS
            string

            .OUTPUTS
            CIMInstance


            #>

                [cmdletbinding()]
    


                Param(

                    [Parameter(Position = 0, `
                               ValueFromPipelineByPropertyName = $true, `
                               ValueFromPipeline = $true, `
                               HelpMessage = "Enter module name",
                               ParameterSetName ='Parameter Set 1') ]

                    [string]$name
    
                )

                Begin {
            

                        # Function Get-CorpModule BEGIN Section

                        Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"  

                        Write-verbose -Message ($PSBoundParameters | out-string)



                } # Function Get-CorpModule BEGIN Section



                Process {

                        # Function Get-CorpModule PROCESS Section

                        if (-not(Get-Module -name $name)) {

                            if (Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) { 
                 
                                Import-Module -Name $name 

                                Write-Verbose -Message "module finished importing "

                                write-output $true 

                             } #end if module available then import 

                            else { # module not available 

                                Write-Verbose -Message "module not available "

                                write-output $false 

                            } # module not available 


                        } # end if not module 

                        else { # module already loaded 

                            Write-Verbose -Message "module already loaded "

                            write-output $true 
                        } 



                } # Function Get-CorpModule PROCESS Section


                End {
                        # Function END Section
            
                         Write-Verbose -Message "Function Get-CorpModule  Ends"
         
                         
           

                } # Function Get-CorpModule END Section


            } # Function Get-CorpModule 


        function get-timestamp {

           get-date -format 'yyyy-MM-dd HH:mm:ss'

         } # function get-timestamp


        function _status {

          param($text,$code)       
        
            if ($code -eq 1) {        
                write-host
                write-host "$text"  -foreground cyan
            }
            if ($code -eq 2) {        
            write-host "$text"  -foreground Magenta  
            } 
            if ($code -eq 3) {        
            write-host "$text"  -foreground White 
            } 
            if ($code -eq 0) {        
            write-host "$text"  -foreground red 
            }   

     } # function _status 


        function Log-Start{
            <#
            .SYNOPSIS
            Creates log file
 
            .DESCRIPTION
            Writes initial logging data
 
            .PARAMETER $LogFullPath
            Mandatory. File name and path name to log file. Eaxmple : C:\temp\myfile.log
  
 
            .INPUTS
            Parameters above
 
            .OUTPUTS
            Log file created
 
            .NOTES
            Version:        1.0
            Author:         Luca Sturlese
            Creation Date:  10/05/12
            Note:           modified by Ammar Hasayen
 
            Version:        1.1
            Author:         Luca Sturlese
            Creation Date:  19/05/12
            Purpose/Change: Added debug mode support
 
            .EXAMPLE
            Log-Start -LogFullPath "C:\Windows\Temp\mylog.log"
            #>
    
            [CmdletBinding()]
  
            Param ([Parameter(Mandatory=$true)][string]$LogFullPath)
   
            Process{    
   
    
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
            Add-Content -Path $LogFullPath -Value "Started processing at [$([DateTime]::Now)]."
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
            Add-Content -Path $LogFullPath -Value ""
   
  
            #Write to screen for debug mode
            Write-Debug "***************************************************************************************************"
            Write-Debug "Started processing at [$([DateTime]::Now)]."
            Write-Debug "***************************************************************************************************"
            Write-Debug ""    
            }
        } # function Log-Start


        function Log-Write{
            <#
            .SYNOPSIS
            Writes to a log file
 
            .DESCRIPTION
            Appends a new line to the end of the specified log file
  
            .PARAMETER LogFullPath
            Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
  
            .PARAMETER LineValue
            Mandatory. The string that you want to write to the log
      
            .INPUTS
            Parameters above
 
            .OUTPUTS
            None
 
            .NOTES
            Version:        1.0
            Author:         Luca Sturlese
            Creation Date:  10/05/12
            Purpose/Change: Initial function development
            Note:           Modified by Ammar Hasayen
  
            Version:        1.1
            Author:         Luca Sturlese
            Creation Date:  19/05/12
            Purpose/Change: Added debug mode support
 
            .EXAMPLE
            Log-Write -LogFullPath "C:\Windows\Temp\Test_Script.log" -LineValue "This is a new line which I am appending to the end of the log file."
            #>
  
          [CmdletBinding()]
  
          Param ([Parameter(Mandatory=$true)][string]$LogFullPath, [Parameter(Mandatory=$true)][string]$LineValue)
  
    
          Process{
    
    
            Add-Content -Path $LogFullPath -Value $LineValue
  
            #Write to screen for debug mode
            Write-Debug $LineValue

          }

    } # function Log-Write


        function Log-Finish{
            <#
            .SYNOPSIS
            Write closing logging data & exit
 
            .DESCRIPTION
            Writes finishing logging data to specified log and then exits the calling script
  
            .PARAMETER LogFullPath
            Mandatory. Full path of the log file you want to write finishing data to. Example: C:\Windows\Temp\Test_Script.log
 
            .INPUTS
            Parameters above
 
            .OUTPUTS
            None
 
            .NOTES
            Version:        1.0
            Author:         Luca Sturlese
            Creation Date:  10/05/12
            Purpose/Change: Initial function development
            Note:           Modified by Ammar Hasayen
    
            Version:        1.1
            Author:         Luca Sturlese
            Creation Date:  19/05/12
            Purpose/Change: Added debug mode support
  
            Version:        1.2
            Author:         Luca Sturlese
            Creation Date:  01/08/12
            Purpose/Change: Added option to not exit calling script if required (via optional parameter)
 
            .EXAMPLE
            Log-Finish -LogFullPath "C:\Windows\Temp\Test_Script.log" 

            #>
  
            [CmdletBinding()]
  
            Param ([Parameter(Mandatory=$true)][string]$LogFullPath)
  
            Process{
            Add-Content -Path $LogFullPath -Value ""
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
            Add-Content -Path $LogFullPath -Value "Finished processing at [$([DateTime]::Now)]."
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
  
            #Write to screen for debug mode
            Write-Debug ""
            Write-Debug "***************************************************************************************************"
            Write-Debug "Finished processing at [$([DateTime]::Now)]."
            Write-Debug "***************************************************************************************************"
  
       
             }
        } # function Log-Finish


        function Write-CorpError {
            
            [cmdletbinding()]

            param(
                [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Error Variable')]$myError,	
	            [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Additional Info')][string]$Info,
                [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Log file full path')][string]$mypath,
	            [switch]$ViewOnly

                )

                Begin {
       
                    function get-timestamp {

                        get-date -format 'yyyy-MM-dd HH:mm:ss'
                    } 

                } #Begin

                Process {

                    if (!$mypath) {

                        $mypath = " "
                    }

                    if($myError.InvocationInfo.Line) {

                    $ErrorLine = ($myError.InvocationInfo.Line.Trim())

                    } else {

                    $ErrorLine = " "
                    }

                    if($ViewOnly) {

                        Write-warning @"
                        $(get-timestamp)
                        $(get-timestamp): $('-' * 60)
                        $(get-timestamp):   Error Report
                        $(get-timestamp): $('-' * 40)
                        $(get-timestamp):
                        $(get-timestamp): Error in $($myError.InvocationInfo.ScriptName).
                        $(get-timestamp):
                        $(get-timestamp): $('-' * 40)       
                        $(get-timestamp):
                        $(get-timestamp): Line Number: $($myError.InvocationInfo.ScriptLineNumber)
                        $(get-timestamp): Offset : $($myError.InvocationInfo.OffsetLine)
                        $(get-timestamp): Command: $($myError.invocationInfo.MyCommand)
                        $(get-timestamp): Line: $ErrorLine
                        $(get-timestamp): Error Details: $($myError)
                        $(get-timestamp): Error Details: $($myError.InvocationInfo)
"@

                        if($Info) {
                            Write-Warning -Message "More Custom Info: $info"
                        }

                        if ($myError.Exception.InnerException) {

                            Write-Warning -Message "Error Inner Exception: $($myError.Exception.InnerException.Message)"
                        }

                        Write-warning -Message " $('-' * 60)"

                     } #if($ViewOnly) 

                     else {
                     # if not view only 
        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): $('-' * 60)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):  Error Report"        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error in $($myError.InvocationInfo.ScriptName)."        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Line Number: $($myError.InvocationInfo.ScriptLineNumber)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Offset : $($myError.InvocationInfo.OffsetLine)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Command: $($myError.invocationInfo.MyCommand)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Line: $ErrorLine"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error Details: $($myError)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error Details: $($myError.InvocationInfo)"
                        if($Info) {
                            Log-Write -LogFullPath $mypath -LineValue  "$(get-timestamp): More Custom Info: $info"
                        }

                        if ($myError.Exception.InnerException) {

                            Log-Write -LogFullPath $mypath -LineValue  "$(get-timestamp) :Error Inner Exception: $($myError.Exception.InnerException.Message)"
            
                        }    

                     }# if not view only

               } # End Process

        } # function Write-CorpError


        function Test-corpIsWsman{

            <#.Synopsis
            Tests if a computer is running WS-MAN and verifies the version of WS-MAN if found.

            .Description
            Test-corpIsWsman takes a computer account, returns information about the support for WSMAN and verifies the version of WS-MAN if found.

            Returns object with two properties :

                - "wsman_supported" is a Boolean that can be 
                    > $true : if the computer is accessible over wsman
                    > $false : if the computer is not accessible over wsman

                - "wsman_version" is a string that can be 
                * "None"      : if we cannot reach the computer using WSMAN
                * "wsman v2" : if we can reach the computer  using WSMAN, and it is version 2.0
                * "wsman v3" : if we can reach the computer  using WSMAN, and it is version 3.0



            Versioning:
                - Version 1.0 written 5  November 2013 : returns an integer that can be 0 "no wsman" , 2 "wsman v2" , 3 "wsman v3"
                - Version 2.0 written 13 November 2013 : returns an object with two properties "wsman_version" and "Wsman_supported" 



            .PARAMETER Computername
            String value representing the computer to test. You can use the following aliases for this parameter
            "Computer","Name","MachineName"

            .Example
            Getting support for WS-MAN on PC1
            PS C:\>Test-corpIsWsman -ComputerName "PC1"

            .Example
            Getting support for WS-MAN on PC1 using -Name parameter alias
            PS C:\>Test-corpIsWsman -Name "PC1"

            .Example
            Getting support for WS-MAN on PC1 using -Computer parameter alias
            PS C:\>Test-corpIsWsman -Computer "PC1"

            .Example
            Running the function without any parameters will default to the localhost as computername
            PS C:\>Test-corpIsWsman

            .Example
            Running the function with error action SilentlyContinue. You will still get an object back if something went wrong like computer does not exist. The function will throw exception always to allow you to catch it if you want, or you can use EA (SilentlyContinue) to get the object result back without noticing the exception.
            PS C:\>Test-corpIsWsman  -Computer "PC1" -EA SilentlyContinue

            .Example
            Get computer names from text file and pipeline the output to Test-corpIsWsman
            PS C:\> get-content computers.txt | Test-corpIsWsman

            .Example
            Get-ADComputer will produce "Name" property, convert it to "ComputerName" so it can be used accross the pipeline
            PS C:\> Get-ADcomputer PC1 |select @{Name="ComputerName";Expression={$_.Name}} |Test-corpIsWsman

            .Notes
            Last Updated             : Nov 13, 2013
            Version                  : 2.0 
            Author                   : Ammar Hasayen (@ammarhasayen)
            based on                 : Jeffery Hicks script(@JeffHicks)

            .Link
            http://ammarhasayen.com
            .Link
            http://jdhitsolutions.com/blog/2013/04/get-ciminstance-from-powershell-2-0

            #>

            [cmdletbinding()]

                Param(

                [Parameter(Position = 0,
                           ValueFromPipeline = $true,
                           ValueFromPipelineByPropertyName = $true)]

                [alias("name","machinename","computer")]

                [string]$Computername=$env:computername
                )


                Begin {
        
                    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"
          
                    #a regular expression pattern to match the ending
                    #gets last digit, then the dot before it, then the digit before the dot
                    [regex]$rx="\d\.\d$"

                    $objparam = @{ wsman_supported = $false
                                   wsman_version   = "None"
                                 }

                }#Test-corpIsWsman function Begin Section


                Process {
        
                    Try {

                        Write-Verbose -Message "testing WSMAN on $Computername using Test-corpIsWsman Function"

                        $result = Test-WSMan -ComputerName $Computername -ErrorAction Stop

                        Write-Verbose -Message "WSMAN is accessible on $Computername"

                        $objparam["wsman_supported"] = $true 

                    }#try


                    Catch {

                        #Write the error to the pipeline if the computer is offline
                        #or there is some other issue

                        Write-Verbose -Message  "Cannot connect to $Computername using WSMAN."

                        Write-Verbose -Message "$ComputerName cannot be accessed using WSMAN ..."

                        New-Object -TypeName PSObject `
                                   -Property $objparam  

                        write-Error $_.exception.message
            
                    }#catch

        
        
                    if ($result) {
            
                        Write-Verbose -Message "Checking WSMAN version on $Computername"

                        $m = $rx.match($result.productversion).value

                             if ($m -eq '3.0') {
                
                                Write-Verbose -Message "$ComputerName running WSMAN version 3.0 ..."
                    
                                $objparam["wsman_version"] = "wsman v3"
                   
                                New-Object -TypeName PSObject `
                                           -Property $objparam
                             }

                             else {
                
                                Write-Verbose -Message "$ComputerName running WSMAN version 2.0 ..."
               
                                $objparam["wsman_version"] = "wsman v2"
                
                                New-Object -TypeName PSObject `
                                       -Property $objparam                
                            }        
        
                    }#end if($result)



                }#Test-corpIsWsman function "Process Section"


                End {

                    Write-Verbose -Message "End Test-corpIsWsman Function"

                }#Test-corpIsWsman function "End Section"



        }  # function Test-corpIsWsman


        function Get-CorpUptime {           
    

            [cmdletbinding()]

                Param(
                [Parameter(Position = 0,
                           ValueFromPipeline = $true,
                           ValueFromPipelineByPropertyName = $true)]
                [alias("Name","MachineName","Computername","Host","Hostname")]
                [string]$Computer,

                [Parameter(Position = 1,
                           ValueFromPipeline = $true,
                           ValueFromPipelineByPropertyName = $true)]
    
                [bool]$UsePSRemote

                )


                Begin {   
     
                # Begin block for Get-CorpUpTime 

                    $now = [DateTime]::Now  

   
                } # Get-CorpUpTime Begin block



                Process {
                # Process block for Get-CorpUpTime

        
            
                        #Write-Verbose -Message "Computer : $computer - Start Processing"

                        $paramOutput = @{computername = $computer}

                        try {

                            # getting data

                            if ($UsePSRemote) {

                                $operatingSystem = Get-CorpCimInfo `
                                                    -Class Win32_OperatingSystem `
                                                    -Computername $computer  `
									                -Property lastbootuptime `
									                -EA Stop
                            }else {

                                $operatingSystem = Get-WmiObject `
                                                    -Class Win32_OperatingSystem `
                                                    -Computername $computer  `
									                -Property lastbootuptime `
									                -EA Stop

                            }

                            # when the connectivity method is not DCOM using (Get-WmiObject), then the returned data is already in DatTime format, so no need to convert.

                            # when the connectivity method is DCOM using Get-WmiObject), then the returned value is string and needs converting.

                            try {
                                # just in case you are using Get-WmiObject not Get-CorpCimInfo
                                $boottime=[Management.ManagementDateTimeConverter]::ToDateTime($operatingSystem.LastBootUpTime) 
                            }
                            catch{
                                $boottime = $operatingSystem.LastBootUpTime
                           }
                          
                            $uptime = New-TimeSpan -Start $boottime -End $now

                
                            # building output object

                            $paramOutput.Add("upTime",$uptime)

                            $objOutput = New-Object -TypeName psobject  `
                                                    -Property $paramOutput
                 

                            Write-Output -InputObject $objOutput

                        } #end try

                        catch {
                            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [Get-CorpUpTime] Error : $computer - fail to get Up Time information, returning N/A"


                            #Write-Verbose -Message "Computer : $computer - fail to get information, returning n/a"

                            # building output object

                            $paramOutput.Add("upTime","n/a")

                            $objOutput = New-Object -TypeName psobject  `
                                                    -Property $paramOutput


                            Write-Output -InputObject $objOutput



                        }# end catch

                        #Write-Verbose -Message "Computer : $computer - END Processing"

      



            } # Get-CorpUpTime Process block



        End{
        #End block for Get-CorpUpTime
            
    
    
        } # Get-CorpUpTime End block



    } # funtion Get-CorpUpTime


        function Get-CorpCimInfo {

            <#
            .Synopsis
            Creates on-the-fly sessions to retrieve WMI data


            .Description
            The Get-CimInstance cmdlet in PowerShell 3.0 can be used to retrieve WMI information from a remote computer using the WSMAN protocol instead of the legacy WMI service that uses DCOM and RPC.
            However, the remote computers must be running PowerShell 3.0 and the latest version of the WSMAN protocol.

            When querying a remote computer,the script will test if the remote computer can be reached using WS-MAN, and will test if it is running PowerShell 3.0 and the latest WS-MAN.
            If this is the case, Get-CIMInstance setups a temporary CIMSession.

            However, if the remote computer is running PowerShell 2.0 and accessible via WSMAN, then PowerShell remoting is used to get data using Get-WMIObject over PowerShell invoke remoting command.
            If the remote computer cannot be contacted using WS-MAN, DCOM is used (CIMSession with a CIMSessionOption to use the DCOM protocol)

            A switch parameter is added to the script called (-LegacyDCOM), which will instruct the script to use (Get-WMIObject) instead of CIMSession with DCOM as a CIMSessionOption. If you add this switch parameter, and the script failback to using DCOM, then a normal Get-WMIObject will be used to get data. The benifit for this is the type of the returned object which is (ManagementObject) in case of Get-Wmiobject. This type of object can be pipelined to Invoke-wmiMethod if needed.This is the only reason to use this switch

            The script has parameter switches to turn off any connectivity method for more customization.Using -DisableWsman3 switch, will disable connecting to remote machines using WSMAN 3.0.etc.

            Managed exceptions will be thrown wisely to allow calling function to catch them and act accordingly. Exceptions will not break the pipeline if you catch exceptions.

            Running the script with Verbose mode will give a lot of information for you to look at.

            This script is essentially a wrapper around Get-CimInstance and PowerShell remoting to make it easier to query data from a mix of computers that have different levels of WS-MAN support.

            Before returning the results back as an object, we will add a property called (CorpComputer) to identify from where the data get generated.

            If (IncludeConnectivityInfo) switch is used, the returned object will also contains a property called (CorpConnectivityMethod) to show which method is used to collect the data

            Both properties can be viewed by pipelining the returned object with (Select *)



            .PARAMETER Class
            Aliase for this parameter is (ClassName). This is the CIM or WMI class to query. You cannot use this paramater with (Query) parameter at same time. (Class) and (Query) are defined in different parameter sets.

            .PARAMETER Computername
            Aliases for this parameter are (Computer,Host,Hostname,Machine,"MachineName"). Default value for this parameter is localhost.

            .PARAMETER Filter
            String to filter CIM or WMI classes.

            .PARAMETER Property
            Reduce the size of returned data by returning only subset of the properties.

            .PARAMETER NameSpace
            CIM/WMI name space to be used. Default is root\cimv2.

            .PARAMETER Query
            Native wmi query to be executed in the remote computer.You can not use this parameter with (Class) parameter.

            .PARAMETER KeyOnly
            Switch parameter to return key values only.Only applicable during CIM sessions, not PowerShell remoting.

            .PARAMETER $OperationTimeoutSec
            Only applicable during CIM sessions, not PowerShell remoting.

            .PARAMETER Shallow
            Only applicable during CIM sessions, not PowerShell remoting.

            .PARAMETER Disablewsman3
            Switch to disable connecting using WS-MAN 3.0 using native Get-CIMInstance over WSMAN.

            .PARAMETER Disablewsman2
            Switch to disable connecting using WS-MAN remoting (PowerShell Remoting).

            .PARAMETER DisableDCOM
            Switch to disable connecting using DCOM (RPC).

            .PARAMETER ShowRunTime
            Switch to show script execution time when running in verbose mode.

            .PARAMETER NoProgressBar
            Switch to hide Progress Bar.By default, a progress bar will be shown to indicate the script progress.

            .PARAMETER includeConnectivityInfo.
            #This switch will add a property to the returned object that indicates the type of method used when retrieving information from each remote computer.

            .PARAMETER LegacyDCOM
            A switch parameter that will instruct the script to use (Get-WMIObject) instead of CIMSession with DCOM as a CIMSessionOption. If you add this switch parameter, and the script failback to using DCOM, then a normal Get-WMIObject will be used to get data. The benifit for this is the type of the returned object which is (ManagementObject) in case of Get-Wmiobject. This type of object can be pipelined to Invoke-wmiMethod if needed.This is the only reason to use this switch.




            .Example
            Get computer names from pipeline.
            PS C:\> get-content computers.txt | Get-CorpCimInfo -class win32_logicaldisk -filter "drivetype=3"

            .Example
            Disable the use of DCOM to access information
            PS C:\> Get-CorpCimInfo -Class Win32_Bios "Localhost","Host1" -DisableDCOM

            .Example
            Only enable DCOM to access information using Get-CIMSession with DCOM as a CIMSessionOption
            PS C:\> Get-CorpCimInfo -Class Win32_Bios "Localhost","Host1" -DisableWsman3 -DisableWsman2

            .Example
            Only enable DCOM to access information and force the script to use the legacy (Get-WMIObject) method
            PS C:\> Get-CorpCimInfo -Class Win32_Bios "Localhost","Host1" -DisableWsman3 -DisableWsman2 -LegacyDCOM

            .Example
            Use Query parameter insted of ClassName
            PS C:\> Get-CorpCimInfo -Query "SELECT * from Win32_Process WHERE name LIKE 'p%'" -Computername "Host1" 

            .Example
            Using Start-Job to pass array of servers to get-corpCimInfo
            [String[]] $list = @("localhost","localhost")

             $job = Start-Job -ScriptBlock { 
                        param ( [String[]] $list )
                        $list | % {Get-CorpCimInfo -Class win32_bios -computerName $_ }
                     } -ArgumentList (,$list)
              Wait-Job -Job $job | Out-Null
              Receive-Job $job


            Versioning:
                - Version 1.0 written  5 November 2013 : flowControl returns string for $protocolToUse
                - Version 2.0 written 13 November 2013 : - FlowControl returns an object for $protocolTouse
                                                         - Better exception handling when getting return values from Test-corpIsWsman
                                             




            .Notes
            Last Updated: Nov 11, 2013
            Version     : 2.0
            Author      : Ammar Hasayen (@ammarhasayen)
            Based on    : Jeffery Hicks script (@JeffHicks)

            .Link
            http://ammarhasayen.com

            .Link
            Get-CimInstance
            New-CimSession
            New-CimsessionOption

            .Inputs
            string

            .Outputs
            CIMInstance

            #>

            [cmdletbinding()]

                Param(

                [Parameter(Mandatory=$true, `
                           HelpMessage="Enter a class name", `
                           ValueFromPipelineByPropertyName=$true, `
                           ParameterSetName="classOption")]
                [alias("ClassName")]
                [ValidateNotNullorEmpty()]
                [string]$Class,

                 #you cannot provide class parameter and query parameter at the same time
                 [Parameter( Mandatory=$true, `
                             HelpMessage="Enter a Wmi Query", `
                             ValueFromPipelineByPropertyName=$true, `
                             ParameterSetName="queryOtpion")]
                 [string]$Query,

                [Parameter(Position=1, `
                           ValueFromPipelineByPropertyName=$true, `
                           ValueFromPipeline=$true, `
                           HelpMessage="Enter one or more computer names separated by commas.") ]
                [ValidateNotNullorEmpty()]
                [alias("Computer","host","hostname","machine","machinename")]
                [string[]]$Computername=$env:computername,

                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [string]$Filter,

                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [string[]]$Property,

                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [ValidateNotNullorEmpty()]
                [string]$Namespace="root\cimv2",
   

                #only available via cim session, when PowerShell remoting is used, this parameter
                #will be ignored
                [switch]$KeyOnly,

                #only available via cim session, when PowerShell remoting is used, this parameter
                #will be ignored
                [uint32]$OperationTimeoutSec,

                #only available via cim session, when PowerShell remoting is used, this parameter
                #will be ignored
                [switch]$Shallow,

                #use this switch to disable Get-CIMSession over WSMAN 3.0 Connectivity
                [switch]$DisableWsman3,

                #use this switch to disable WSMAN remoting (PowerShell Remoting)
                [switch]$DisableWsman2,

                #use this switch to disable DCOM Connectivity 
                [switch]$DisableDcom,

                #use this switch to enabel DCOM legacy Get-WMIObject when failing back to DCOM
                [switch]$LegacyDCOM,

                 #use this switch to show script execution time when running in verbose mode
                [switch]$ShowRunTime,

                [switch]$NoProgressBar,
                #Switch to hide Progress Bar

                [switch]$includeConnectivityInfo
                #This switch will add a property to the returned object that indicates the type of method used when retrieving information from each remote computer

               #[Parameter(Mandatory=$false)]               
               #[System.Management.Automation.Credential()]$Credential
               #uncomment this parameter if you need to supply a credential

                )#end function parameters


                Begin{
            
                        if ($PSBoundParameters.ContainsKey("ShowRunTime")){
                            #Start stop watch
                            $Watch  =  [System.Diagnostics.Stopwatch]::StartNew()
                        }


                        if (!(($PSBoundParameters.ContainsKey("query") ) -OR ($PSBoundParameters.ContainsKey("class"))) ) {
                            Throw " You should provide either Class parameter or Query Parameter"
                        }

                        if (($PSBoundParameters.ContainsKey("query") ) -AND ($PSBoundParameters.ContainsKey("class")) ) {
                            Throw " You cannot provide both Class parameter and Query Parameter"
                        }

            
                        Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"  

                        Write-verbose -Message ($PSBoundParameters | out-string)
    
                        #defining the methods to use when connecting to the remote computer
                        #by default all methods will be performed
                        #by default, wsman3 will be tried first, if it fails, then wsman2, then dcom

                        $propertiesMethods = @{wsman3 = $true;
                                               wsman2 = $true;
                                               dcom   = $true
                                              }
                        $Methods = New-Object -TypeName PSObject -Property $propertiesMethods
                                         
                        if ($PSBoundParameters.ContainsKey("disableWsman3"))  {$Methods.wsman3 = $false}
                        if ($PSBoundParameters.ContainsKey("disableWsman2"))  {$Methods.wsman2 = $false}
                        if ($PSBoundParameters.ContainsKey("disabledcom"))    {$Methods.dcom   = $false}

                        #check if wsman3,wsman2 and dcom are all disabled via parameter switches
                        if (!($Methods.dcom  -OR $Methods.wsman2 -OR $Methods.wsman3 )){
                           Throw  "You cannot disable all test types"
                        }


                        #printing out available tests
                        Write-verbose -Message "Tests available for the script : $(if($Methods.wsman3){"WS-MAN3"} if($Methods.wsman2){"WS-MAN2"} if($Methods.DCOM){"DCOM"})" 
             
          
                        function Get-corpProgress {
	            
		                    param($PercentComplete,$status)
		    
		                    Write-Progress -activity "Get-corpCimInfo Script" `
                                          -percentComplete ($PercentComplete) `
                                          -Status $status

	                    }#end Get-corpProgress function

            
                    
            
                        function flowControl {
                        
                        #route the script execution depending on the connectivity methods available 
                       
                            Param (
                                [Parameter(Mandatory=$true)]
                                [string]$Computername,

                                [Parameter(Mandatory=$true)]
                                [object]$wsSupport,
                    
                                [Parameter(Mandatory=$true)]
                                [object]$Methods,

                                [Parameter(Mandatory=$true)]
                                [object]$ProtocolToUse
                             )

                             Begin {   #function flowControl

                             Write-Verbose -Message "+++++++Entering flowControl function"

                             }   #begin function flowControl

                             Process {#function flowControl
                    
                                #Possibility 1:

                                if ($ProtocolToUse.Method -like "Nothing") {
                    
                                    if ($Methods.wsman3 -and ($wsSupport.wsman_version -like "wsman v3")) {
                                        $ProtocolToUse.Method = "WSMAN 3.0"
                                         Return $ProtocolToUse
                            
                                    }

                                    elseif ($Methods.wsman2 -and ($wsSupport.wsman_supported)) {
                                        $ProtocolToUse.Method = "WSMAN 2.0"
                                         Return $ProtocolToUse
                                    } 

                                    elseif ($Methods.dcom) {
                                        $ProtocolToUse.Method = "dcom"
                                        Return $ProtocolToUse
                                    }

                                    else {
                                        Write-Verbose -Message "None of the available connectivity methods works with $Computername"  
                                        Write-Error "+++++++Exiting script with managed exception for you to catch... cannot connect to $Computername using any configured method"
                                    }

                                }#end ($ProtocolToUse.Method -like "Nothing")  
                    
                    
                                #Possibility 2:
                                if ($ProtocolToUse.Method -like "WSMAN 3.0") { #means WSMAN 3.0 conenctivity failed
                        
                                    if ($Methods.wsman2 -AND ($wsSupport.wsman_supported) )  {
                                        $ProtocolToUse.Method = "WSMAN 2.0"
                                        Write-Verbose -Message "failing back to $($ProtocolToUse.Method) for $Computername"
                                        Return $ProtocolToUse
                                    } 

                                    elseif($Methods.dcom) {
                                        $ProtocolToUse.Method = "dcom"
                                        Write-Verbose -Message "failing back to $($ProtocolToUse.Method) for $Computername"
                                        Return $ProtocolToUse
                                    } 
                          
                                    else {
                                        Write-Verbose -Message "no failback method for $Computername.. +++++++Exiting"
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $Computername using any configured method"
                                    }



                              }#end if ($ProtocolToUse.Method -like "WSMAN 3.0")



                                #Possibility 3:
                               if ($ProtocolToUse.Method -like "WSMAN 2.0") { #means WSMAN 2.0 conenctivity failed
                        
                                    if ($Methods.dcom) {
                                        $ProtocolToUse.Method = "dcom"
                                        Return $ProtocolToUse
                                    }

                                     else {
                                        Write-Verbose -Message "no failback method for $Computername.. +++++++Exiting"
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $Computername using any configured method"
                                    }

                               }#end if ($ProtocolToUse.Method -like "WSMAN 2.0")


                     


                             }#process function flowControl

                             End { #function flowControl

                                Write-Verbose -Message "Protocol to be used $($ProtocolToUse.Method)"

                                Write-Verbose -Message "+++++++Exiting flowControl function"

                             }#end block of function flowControl 
                

                        }#end Function flowControl

                   }#end Get-CorpCisInfo function begin block
    



                Process {  #process block for Get-corpCisInfo function

                        Write-Verbose -Message "Processing $($computername.count) computer(s)"

                        [int]$counter = 0 #used to show progress only



                        foreach ($computer in $computername) {       
     
                            if (!($PSBoundParameters.ContainsKey("NoProgressBar"))) {

                                $counter += 1

                                Get-corpProgress -PercentComplete (($counter/($Computername.count))*100) `
                                                 -Status "getting information from $computer"
                             }#end if

                
                            #This is a variable to hold the decision on what method to be used to connect to the remote computer. The flowControl function is the only one the can change this object
                            #All routing decisions are taken based on the value of this object
                
                            $protocolToUse = New-Object -TypeName psobject `
                                                        -Property @{Method = "Nothing"}
                 
                            Write-Verbose -Message "Processing $computer"

                
                            #First thing to do is to hand control to the flowControl funtion
                            #to do that, we will evaluate WSMAN support first using a variable $wsSupport
                            #WSMAN support means testing if the remote computer is accessible via WSMAN
                                
                            #clearing error to get only error data from Test-corpIsWsman function
                            $error.clear()
                
                            if($Methods.wsman3 -OR $Methods.wsman2) {

                               $wsSupport= Test-corpIsWsman `
                                                -ComputerName $computer `
                                                -ErrorAction SilentlyContinue
                  
                            }#end if

                            else {

                                 $objparam = @{ wsman_supported = $false
                                                wsman_version   = "None"
                                              }   
                                 $wsSupport = New-Object -TypeName PSObject `
                                   -Property $objparam 
                            }



                            #building flowControl function parameters:
                            $paramflowControl = @{ ComputerName  = $computer;
                                                   Methods       = $Methods;
                                                   ProtocolToUse = $protocolToUse;
                                                   wsSupport     = $wsSupport
                                                 }
                                 
                            #we will lt the flowControl function determin what method to use to connect
                            $protocolToUse = flowControl @paramflowControl

                            #Since exception is not thrown, then there is a method to be evaluated
            
                            #hashtable of parameters for New-CimSession or remoting
                            #adding computername and EA

                            $sessParam = @{Computername=$computer;ErrorAction='Stop'}

                            #credentials?
                            if (($PSBoundParameters.ContainsKey("Credential")))
                                {

                                if ($credential) {
                                    Write-Verbose -Message "Adding alternate credential for CIMSession"
                                    $sessParam.Add("Credential",$Credential)
                                }#end if if ($credential)

                            }#end if


                             #WSMAN 3.0 method
                             if ($protocolToUse.Method -like "WSMAN 3.0") {
                                 
                                 Write-Verbose -Message "trying $($protocolToUse.Method) on $computer"

                                 Try {               
                                     $session = $null
                                     $session = New-CimSession @sessParam
                                     Write-Verbose -Message "Session using $($protocolToUse.Method) is created on $computer"
                                 }#end try

                                 Catch {
                                     Write-Warning "Failed to create a CIM session to $computer using $($protocolToUse.Method)"
                                     Write-Warning $_.Exception.Message
                                     #setting alternative method

                                     #building flowControl function parameters:
                                     $paramflowControl = @{ ComputerName  = $computer;
                                                            Methods       = $Methods;
                                                            ProtocolToUse = $protocolToUse;
                                                            wsSupport     = $wsSupport
                                                           }
                                 
                                     #we will lt the flowControl function determin what method to use to connect
                                     $protocolToUse = flowControl @paramflowControl                           
                                                              
                                 }#end catch

                                 if ($session){

                                     #create the parameters to pass to Get-CIMInstance
                                     $paramHash=@{
                                                  CimSession= $session
                                                  }

                                     $cimParams = "Filter","KeyOnly","Shallow","OperationTimeOutSec","Namespace","Query","Class"

                                     foreach ($param in $cimParams) {

                                        if ($PSBoundParameters.ContainsKey($param)) {

                                             Write-Verbose -Message "Adding $param for CIM command : $computer"
                                             $paramhash.Add($param,$PSBoundParameters.Item($param))
                                         } #if

                                     }#foreach param 
                    
                     

                                     #execute the query
                                     Write-Verbose -Message "Querying $class using $($protocolToUse.Method) on $computer"
                    
                                     Try {
                                        $obj = Get-CimInstance @paramhash -EA Stop
                                        #we add a property called "corpComputer" that will help identify the machine returning this info
                                        #this become useful when doing PowerShell Jobs (start-Job)
                                        $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer
                            
                                        if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                     -MemberType NoteProperty `
                                                     -Name "corpConnectivityMethod" `
                                                     -Value $protocolToUse.Method
                                        }
                                        $obj
                                     }#end try

                                     Catch {
                        
                                         Write-Verbose -Message " Error while Querying $class on $computer via $($protocolToUse.Method)" 

                                         #writing out exception
                                         Write-Warning $_.Exception.Message

                                         #setting alternative method
                                         #building flowControl function parameters:
                                         $paramflowControl = @{ ComputerName  = $computer;
                                                                Methods       = $Methods;
                                                                ProtocolToUse = $protocolToUse;
                                                                wsSupport     = $wsSupport
                                                               }
                                 
                                         #we will lt the flowControl function determin what method to use to connect
                                         $protocolToUse = flowControl @paramflowControl      

                                      }#end catch

                                      Finally {
                                         Write-Verbose "Removing CIM 3.0 Session from $computer"
                                         if ($session) {Remove-CimSession $session}
                                      }#end finally

                                 }#end if($session)
                
                             }#end if($protocolToUse.Method -like "WSMAN 3.0")      

                  
                             #WSMAN 2.0 method
                             if ($($protocolToUse.Method) -like "WSMAN 2.0") {
                
                                 $paramHash=@{}
                                 $wsParams = "Filter","Namespace","Query","Property","class"
                
                                 foreach ($param in $wsParams) {
                    
                                    if ($PSBoundParameters.ContainsKey($param)) {
                        
                                        Write-Verbose -Message "Adding $param for PS remoting command : $computer"
                                        $paramhash.Add($param,$PSBoundParameters.Item($param))

                                    } #end if ($PSBoundParameters.ContainsKey($param))

                                }#foreach ($param in $wsParams       

                                #execute the query
                                Write-Verbose -Message "Querying $class using $($protocolToUse.Method) on $computer"
                
               
                                Try {  

                                    $wssession = $null
                         
                                    $wssession = New-PSSession @sessParam
                                  
                                    $obj = Invoke-Command -Session $wssession `
                                                          -ScriptBlock{param($x) Get-WmiObject @x} `
                                                          -ArgumentList $paramhash `
                                                          -EA Stop

                                    #we add a property called "corpComputer" that will help identify the machine returning this info
                                    #this become useful when doing PowerShell Jobs (start-Job)
                                    $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer

                                    if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                     -MemberType NoteProperty `
                                                     -Name "corpConnectivityMethod" `
                                                     -Value $protocolToUse.Method
                                        }

                                    $obj
                                    Write-Verbose -Message "Done : querying $class using $($protocolToUse.Method) on $computer"

                                }#end try

                
                                Catch {

                                    Write-Verbose -Message " Error while Querying $class on $computer via $($protocolToUse.Method)"

                                    #writing out exception
                                    Write-Warning $_.Exception.Message
                                        
                                    #setting alternative method

                                    #building flowControl function parameters:
                                    $paramflowControl = @{ ComputerName  = $computer;
                                                           Methods       = $Methods;
                                                           ProtocolToUse = $protocolToUse;
                                                           wsSupport     = $wsSupport
                                                          }
                                 
                                     #we will lt the flowControl function determin what method to use to connect
                                     $protocolToUse = flowControl @paramflowControl  

                                }#end catch  

                                Finally {

                                    Write-Verbose -Message "Removing PSSession from $computer"
                                    if($wssession) {Remove-PSSession $wssession}

                                }#end Finally   
                          
                            }#end if ($protocolToUse.Method -like "WSMAN 2.0")



                            #dcom method
                            if($protocolToUse.Method -like "dcom")  {
                
                
                                 write-verbose "trying $($protocolToUse.Method) on $computer"

                                 if ($PSBoundParameters.ContainsKey("LegacyDCOM")) {
                                 #using legacy Get-WMIObject command

                                    write-verbose "using Legacy DCOM"

                                    $paramHash=@{
                                                 ComputerName = $computer
                                                }
                        
                                    $dcomLegacyParams = "Filter","Namespace","Query","Property","class"

                                    foreach ($param in $dcomLegacyParams) {
                       
                                        if ($PSBoundParameters.ContainsKey($param)) {
                       
                                             Write-Verbose -Message "Adding $param for Legacy dcom command : $computer"
                        
                                            $paramhash.Add($param,$PSBoundParameters.Item($param))

                                        }#if

                                    } #foreach param 

                                     #execute the query
                                    Write-Verbose "Querying $class using legacy $($protocolToUse.Method) on $computer"
                                    Try {

                                        $obj = Get-WmiObject @paramHash -EA Stop
                                        #we add a property called "corpComputer" that will help identify the machine returning this info
                                        #this become useful when doing PowerShell Jobs (start-Job)

                                        $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer

                                        if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                            -MemberType NoteProperty `
                                                            -Name "corpConnectivityMethod" `
                                                            -Value $protocolToUse.Method
                                         }#end if

                                        $obj #return object
                       
                                    }#end Try

                                     Catch {

                                        Write-Verbose -Message " Error while Querying $class on $computer via legacy $($protocolToUse.Method)" 

                                        #writing out exception
                                        Write-Warning $_.Exception.Message

                                        #Write-Error  exception and existing
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $computer using any configured method"
                                                  
                        
                                     }#end Catch

                                     Finally {
                                     }#end Finally 

                                 }#end if ($PSBoundParameters.ContainsKey("LegacyDCOM"))

                                 else {
                                 #use Get-CIMInstance with CIMSessionOption (Protocol Dcom)
                                 $opt = New-CimSessionOption -Protocol Dcom
                                 $sessparam.Add("SessionOption",$opt)  


                                 Try {               
                                    $session = $null
                                    $session = New-CimSession @sessParam
                                    write-verbose -Message "Session using $($protocolToUse.Method) is created on $computer"
                                  }#end try

                                 Catch {
                                    Write-Warning "Failed to create a CIM session to $computer using $($protocolToUse.Method)"
                                    Write-Warning "No other actions will be performed"
                                    Write-Warning $_.Exception.Message  
                                    Write-Error "+++++++Exiting script with managed exception for you to catch... cannot connect to $computer using any configured method"  
                                 }#end catch

                                 if($session){


                                    #create the parameters to pass to Get-CIMInstance
                                    $paramHash=@{
                                         CimSession= $session
                                                }


                                    $cimParams = "Filter","KeyOnly","Shallow","OperationTimeOutSec","Namespace","Query","Property","class"

                                    foreach ($param in $cimParams) {
                       
                                        if ($PSBoundParameters.ContainsKey($param)) {
                       
                                             Write-Verbose -Message "Adding $param for CIM command : $computer"
                        
                                            $paramhash.Add($param,$PSBoundParameters.Item($param))

                                        }#if
                                    } #foreach param 
                    
                     

                                    #execute the query
                                    Write-Verbose "Querying $class using $($protocolToUse.Method) on $computer"
                                    Try {

                                        $obj = Get-CimInstance @paramhash -EA Stop
                                        #we add a property called "corpComputer" that will help identify the machine returning this info
                                        #this become useful when doing PowerShell Jobs (start-Job)

                                        $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer

                                        if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                            -MemberType NoteProperty `
                                                            -Name "corpConnectivityMethod" `
                                                            -Value $protocolToUse.Method
                                         }#end if

                                        $obj #return object
                       
                                    }#end Try

                                     Catch {

                                        Write-Verbose -Message " Error while Querying $class on $computer via $($protocolToUse.Method)" 

                                        #writing out exception
                                        Write-Warning $_.Exception.Message

                                        #Write-Error  exception and existing
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $computer using any configured method"
                                                  
                        
                                     }#end Catch

                                     Finally {

                                        Write-Verbose -Message "Removing CIM Session from $computer"
                                        if($session){Remove-CimSession $session}

                                     }#end Finally 

                                }#end if($session)
               
                             }#end else



                             }#end if($ProtocolToUse.Method -like "dcom") for dcom    
  

                             Write-Verbose -Message "Finish processing $computer. Last Method used is $($protocolToUse.Method)"

                       }#end foreach ($computer in $computername)
 
                }#end function process block


                End { 
                     Write-Verbose -Message "Script Get-CorpCimInfo Ends"
         
                     if ($PSBoundParameters.ContainsKey("ShowRunTime")) {
                         #Stop and display stop watch
            
                         Write-Verbose -Message "Script run time (Minutes:seconds:milliseconds): $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString()):$($Watch.Elapsed.MilliSeconds.ToString())"
                     }#end if ($PSBoundParameters.ContainsKey("ShowRunTime"))

                }# end function end block



        } # function Get-CorpCimInfo function       


        function _screenheadings {

            Cls
            write-host 
            write-host 
            write-host 
            write-host "--------------------------" 
            write-host "Script Info" -foreground Green
            write-host "--------------------------"
            write-host
            write-host " Script Name : BackupExec Info (Get-BEJOBs)"  -ForegroundColor White
            write-host " Author      : Ammar Hasayen @ammarhasayen" -ForegroundColor White         
            write-host " Version     : 1.1"   -ForegroundColor White
            write-host
            write-host "--------------------------" 
            write-host "Script Release Notes" -foreground Green
            write-host "--------------------------"
            write-host                 
            write-host            
            Write-Host " http://ammarhasayen.com"  -ForegroundColor White
            write-host  
            write-host "--------------------------" 
            write-host "Script Start" -foreground Green
            write-host "--------------------------"
            Write-Host
            
    } # function _screenheadings


        function _screenFooter {           
            param($path)
                
                write-host 
                write-host 
                write-host 
                write-host "--------------------------" 
                write-host "Script Log Files" -foreground Green
                write-host "--------------------------"
                write-host
                write-host " Three log file are created under $path :"  
                write-host "    - Info Log"  -ForegroundColor Red 
                write-host "    - Detailed Log"  -ForegroundColor Red 
                write-host "    - Error" -ForegroundColor Red               
                write-host             
                write-host "--------------------------" 
                write-host "Script Ends" -foreground Green
                write-host "--------------------------" 
            
        } # function __screenFooter


        function Write-DetailedJobStatisLog {
            
            param( $obj_bag, $DetailedFullPath)

            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): "
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): -------------"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Detailed Info"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): -------------"

            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Name  : $($obj_bag.Name)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Storage  : $($obj_bag.Storage)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Selection Summary : $($obj_bag.SelectionSummary)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Start Time : $($obj_bag.StartTime)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Elapsed Time : $($obj_bag.ElapsedTime)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Status : $($obj_bag.JobStatus)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Media Label : $($obj_bag.Medialabel )"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Total Bytes : $($obj_bag.TotalDataSizeBytes)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Job Rate MB Per Min : $($obj_bag.JobRateMBPerMinute)"

            } # function Write-DetailedJobStatisLog

    #endregion helper functions


    #region GUI functions

        function _drawme{

            param($array_unsorted)

            $array = $array_unsorted |Sort-Object -Property StartTime

                [int]$C=0
                $Output ="<html>
                    <body>
                    <font size=""1"" face=""Arial,sans-serif"">
                    <h1 align=""center"">BackupExec Job Report</h3>
                    <h3 align=""center"">Generated $((Get-Date).ToString())</h5>
                    </font>"


                        $Output+="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">	
	                    <tr align=""center"" bgcolor=""#6699FF"">"
	
	                    $Output+="<th><font color=""#ffffff"">#</font></th>
                                    <th><font color=""#ffffff"">Name</font></th>
                                      <th><font color=""#ffffff"">Storage</font></th>                
                                    <th><font color=""#ffffff"">StartTime</font></th>
                                    <th><font color=""#ffffff"">ElapsedTime</font></th>
                                    <th><font color=""#ffffff"">JobStatus</font></th>
                                    <th><font color=""#ffffff"">TotalDataSizeBytes </font></th>              
                                    <th><font color=""#ffffff"">JobRateMBPerMinute </font></th>
                                    <th><font color=""#ffffff"">SelectionSummary </font></th>
                                    <th><font color=""#ffffff"">Medialabel</font></th>"
				
	                    $Output+="</tr>"
	                    $AlternateRow=0;
 


                        foreach ($job in $array )
	                    {	$C++

		                    $Output+="<tr"
		                    if ($AlternateRow)
		                    {
			                    $Output+=" style=""background-color:#dddddd"""
			                    $AlternateRow=0
		                    } else
		                    {
			                    $AlternateRow=1
		                    }
		
		                    $Output+=">"
	
	                        $Output+="<td  align=""center""><strong>$C</strong></td>"
		                    $Output+="<td  align=""center""><strong>$($job.Name)</strong></td>"	
                            $Output+="<td  align=""center""><strong>$($job.Storage)</strong></td>"
                            $Output+="<td  align=""center"">$($job.StartTime )</td>"	
                            $Output+="<td  align=""center"">$($job.ElapsedTime)</td>"
                	
                	
		                    if ( $job.JobStatus  -like "*error*") {
                                $Output+="<td align=""center""><font color=""#FF0099""><Strong>$($job.JobStatus)</Strong></font></td>"
                            }else{
                                $Output+="<td  align=""center""><font color=""#339900""><strong>$($job.JobStatus)</strong></font></td>"	
                            }

                            	
                            $Output+="<td  align=""center"">$($job.TotalDataSizeBytes )</td>"	
                            $Output+="<td  align=""center"">$($job.JobRateMBPerMinute)</td>"
                            $Output+="<td  align=""center"">$($job.SelectionSummary)</td>"
                            $Output+="<td  align=""center""><strong>$($job.Medialabel)</strong></td>"
                
                

                            $Output+="</tr>"
		
	       
	

                        }
                        $Output+="</table><br />"

                $Output


         } # function _drawme

    #endregion gui functions

#endregion Module Functions


#region Module Factory

            
    #region create directory
    
            try{
                $ScriptFilesPath = Convert-Path $ScriptFilesPath -ErrorAction Stop
            }catch {
                Write-CorpError -myError $_ -ViewOnly -Info "[Creating Files - Validating log files path] Sorry, please check the sript path name again"
                Exit
                throw " [Creating Files  - Creating files] Validating log files path] Sorry, please check the script path name again "
            }

            $ScriptFilesPath = Join-Path $ScriptFilesPath "BEScriptOutputFiles"

            if(Test-Path $ScriptFilesPath ) {
                try {
                        Remove-Item $ScriptFilesPath -Force -Recurse -ErrorAction Stop

                    }catch {

                        Write-CorpError -myError $_ -ViewOnly -Info "[Creating Files - Deleting old working directory] Could not delete directory $ScriptFilesPath"
                        Exit
                        throw "[Creating Files - Deleting old working directory] Could not delete directory $ScriptFilesPath "

                    }
            }
    
    
            if(!(Test-Path $ScriptFilesPath )) {
                try {
                    New-Item -ItemType directory -Path $ScriptFilesPath -ErrorAction Stop
                }catch{
                    Write-CorpError -myError $_ -ViewOnly -Info "[Creating Files - Creating working directory] Could not delete directory $ScriptFilesPath"
                    Exit
                    throw "[Creating Files - Creating working directory] Could not delete directory $ScriptFilesPath "

                }
            }  
    
    
    #endregion create directory
                
    #region Create files
            
        _status "    1.1 Creating Files" 1
            
        #region ErrorLog            
           
            try{                
                $ErrorLogFile = "ReportError.log"    
                $ErrorFullPath = Join-Path $ScriptFilesPath  $ErrorLogFile -ErrorAction Stop
            }catch {
                Write-CorpError -myError $_ -ViewOnly -Info "[Create Files - Validating log files] Sorry, please check the script path name again"
                Exit
                throw " [Create Files - Creating files] Sorry, please check the script path name again "

            }
             
            #Check if file exists and delete if it does

            If((Test-Path -Path $ErrorFullPath )){
                try {
                Remove-Item -Path $ErrorFullPath  -Force -ErrorAction Stop
                }catch {
                    Write-CorpError -myError $_ -ViewOnly -Info "[Create Files - Deleting old log files]"
                    Exit
                    throw " [Create Files - Creating files] Sorry, but the script could not delete log file on $ErrorFullPath "
                }
            }
    
            #Create Error file

            try {
                New-Item -Path $ScriptFilesPath -Name $ErrorLogFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Create Files - Creating log files]"
                 Exit
                throw "  [Create Files - Creating files] Sorry, but the script could not create log file on $ScriptFilesPath [Module 3 Factory - Creating files] "
            }

            #initiate error log file

            Log-Start -LogFullPath $ErrorFullPath

            Write-verbose -Message "[Create Files] : Error Log File created $ErrorFullPath"

        #endregion ErrorLog

        #region InfoLog

            try{
                $InfoLogFile = "ReportInfo.log"    
                $InfoFullPath =  Join-Path $ScriptFilesPath  $InfoLogFile -ErrorAction Stop
            }catch{
                Write-CorpError -myError $_ -ViewOnly -Info "[Create Files - Validating log files] Sorry, please check the script path name again"
                Exit
                throw " [Create Files - Creating files] Sorry, please check the sript path name again "

            }

            #Check if file exists and delete if it does
            If((Test-Path -Path $InfoFullPath )){
                try {
                    Remove-Item -Path $InfoFullPath  -Force -ErrorAction Stop
                }catch {
                    Write-CorpError -myError $_ -ViewOnly -Info "[Create Files - Deleting old log files]"
                    Exit
                    throw "[Create Files] : Sorry, but the script could not delete log file on $InfoFullPath [Module 3 Factory - Creating files] "
                }
            }
    
            #Create Info file
            try {
                New-Item -Path $ScriptFilesPath -Name $InfoLogFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Create Files - Creating log files]"
                 Exit                 
                 throw "[Create Files] : Sorry, but the script could not delete log file on $ScriptFilesPath [Module 3 Factory - Creating files] "
            }

           
            #initiate Info log file

             Log-Start -LogFullPath $InfoFullPath

             Write-verbose -Message "[Create Files] : Info Log File created $InfoLogFile"
            

            

        #endregion Info Log

        #region DetailedLog            
           
            $DetailedLogFile = "ReportDetailed.log"    
            try {
                $DetailedFullPath =  Join-Path $ScriptFilesPath  $DetailedLogFile -ErrorAction Stop
            }catch {
                Write-CorpError -myError $_ -ViewOnly -Info "[Creating Files - Validating log files] Sorry, please check the sript path name again"
                Exit

            }
             
            #Check if file exists and delete if it does

            If((Test-Path -Path $DetailedFullPath )){
                try {
                Remove-Item -Path $DetailedFullPath  -Force -ErrorAction Stop
                }catch {
                     Write-CorpError -myError $_ -ViewOnly -Info "[Creating Files - Deleting old log files]"
                     Exit
                     throw " [Creating Files - Creating files] Sorry, but the script could not delete log file on $DetailedFullPath "
                }
            }
    
            #Create Detailed file

            try {
                New-Item -Path $ScriptFilesPath -Name $DetailedLogFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Creating Files - Creating log files]"
                 Exit                
                throw "  [Creating Files - Creating files] Sorry, but the script could not create log file on $ScriptFilesPath [Creating Files - Creating files] "
            }

            #initiate  Detailed file log file

            Log-Start -LogFullPath $DetailedFullPath

            Write-verbose -Message "[Creating Files] : Detailed Log File created $DetailedFullPath"

        #endregion DetailedLog        
    

    #endregion Create files

    #region initial log info 
            
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Info log file $($InfoFullPath) is created successfully"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Error log file $($ErrorFullPath) is created successfully" 
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Error log file $($DetailedFullPath) is created successfully"      


        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Starting $($MyInvocation.Mycommand)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): ($PSBoundParameters | out-string)"
        
        if($PSVersionTable.PSVersion.major){
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): PowerShell Host Version :$($PSVersionTable.PSVersion.major)"
        }

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Script Options : "
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - Name Like Filter = ""$NameLike"""
        if ($PSBoundParameters.ContainsKey("FromLastJobRun")) {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - Option Used : FromLastJobRun"
            $LastRunOption = $true 
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - Option Used = Days ($Days)"
            $LastRunOption = $false
        }
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - Scipt Path = $ScriptFilesPath"
        

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module  Creating Files] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module Factory]"

        
 
    #endregion initial log info

    #region Screen Headings

            Write-Verbose -Message "Info : Starting $($MyInvocation.Mycommand)"  

            Write-verbose -Message ($PSBoundParameters | out-string)

            _screenheadings

            #region general Info
                Write-Output ""
                Write-Output " Script Parameters:"
                Write-Output "   - Filters applied to Get-BEJOB is -Name ""$NameLike""" 

                if ($LastRunOption) {
                    Write-Output "   - Option Used : From Last Job Run"
                    
                }else {
                     Write-Output "   - Option Used : Days since backup = $days days."
                     
                }
                
                Write-Output ""
             
            #endregion #region general Info

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Clearing the console screen and displaying script headings"

           

            _status " 1. Creating Files and Directories" 2 

        #endregion   Screen Heading    

    #region Variables

    
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Creating Global Variables"
        
             Write-Output ""
             _status " 2. Creating global variables" 2

             
             if ((Get-CorpModule -name BEMCLI)) {
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):          BEMCLI Module is loaded  " 

             }else {
                _status " ----------BEMCLI Module cannot be loaded---------- " 0
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):          BEMCLI Module is loaded  " 
             Exit
             }    


             #region variables

                 
                 $today = get-date
                 $temp = 0 - $Days
                 $since = $today.AddDays($temp)

                 $array_obj = @()
    
                 $array_bag = @()

                 $obj_prop = @{ Name = ""
                                SelectionSummary = ""
                                Storage = ""
                                JobHistory = @()                    
                                }
                 
                 $bag_prop = @{  Name = ""
                                 SelectionSummary = ""
                                 Storage =""
                                 StartTime = ""
                                 ElapsedTime = ""                     
                                 JobStatus = ""
                                 Medialabel = ""   
                                 TotalDataSizeBytes  = "" 
                                 JobRateMBPerMinute = ""                   
                             }
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):           variables created"

             #endregion variables

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module  Factory] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): " 


    #endregion variables


#endregion Module Factory


#region Module Process

        
        Write-Output ""
        _status " 3. Processing Jobs" 2
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module  Process] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): " 

    #region Get-BEJob
        
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): " 
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Getting Backup Jobs using Get-BEJob -Name ""$NameLike"""
        _status "          Getting Backup Jobs using Get-BEJob -Name ""$NameLike"""  1  
        
        
        [array]$Jobs = @(Get-BEJob -Name $NameLike )       
        

        if (!$jobs -or $jobs.count -eq 0) {
        
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): No Jobs are returned.. Exiting"
            _status " ----------NO JOBS ARE RETURNED (Get-BEJOB---------- " 0  
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "
            Exit 
        }
        else {

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Jobs returned = $($Jobs.count)"
            Write-Output " "
            _status "          Jobs returned = $($Jobs.count)      " 1  
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "            

        }

        Write-Output ""
        _status " 4. Processing Jobs Details" 2
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Jobs List:"

        foreach ($job in $jobs) {

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):      $($job.name)"

            $obj = New-Object -TypeName PSObject -Property $obj_prop
        
            $obj.Name = $job.name
        
            $obj.SelectionSummary = $job.SelectionSummary 

            $obj.storage = $job.storage
        
            if ($PSBoundParameters.ContainsKey("FromLastJobRun")) {
                $obj.JobHistory = $job |Get-BEJobHistory -FromLastJobRun
            }else{
                $obj.JobHistory = $job |Get-BEJobHistory -FromStartTime $since -ToStartTime $today
            } 
        
            $array_obj += $obj  
        }
    
    #endregion Get-BEJob

    #region create array of custom objects

         _status "          Inspecting each job for Job History and logs      " 1  
         foreach ($obj in $array_obj) {
 
           If ($obj.JobHistory) {

                    #if the job has history

                    $history_Array = $obj.Jobhistory
      
                       foreach ($history in $history_Array ) {

                           $obj_bag = New-Object -TypeName PSObject -Property $bag_prop 

                           $obj_bag.Name = $obj.name
                           $obj_bag.SelectionSummary =  $obj.SelectionSummary
                           $obj_bag.Storage =  $obj.Storage             
                           $obj_bag.StartTime = $history.StartTime
                           $obj_bag.ElapsedTime = $history.ElapsedTime  
                           $obj_bag.JobStatus =  $history.JobStatus         
                           $obj_bag.TotalDataSizeBytes =$history.TotalDataSizeBytes
                           $obj_bag.JobRateMBPerMinute  = $history.JobRateMBPerMinute
                           
                           
                           
                           $tempErrorPref = $ErrorActionPreference
                           $ErrorActionPreference = 'Stop'
                           try{
                               $log =  $history | Get-BEJobLog 
                               $myindex = $log.IndexOf("Media Label:") 
                               $tape = $log.Substring($myindex+13) 
                               $TapeValue=($tape -split '[\r\n]') |? {$_} 
                               $obj_bag.Medialabel = $TapeValue[0] 
                           }catch{
                                
                               $obj_bag.Medialabel = "N/A w E" 
                               Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Error getting Job Log (Get-BEJobLog) for $($obj_bag.name)"
                               Write-Output " screw $($obj_bag.name)"
                           }finally{
                           $ErrorActionPreference = $tempErrorPref
                           }
                           
                           $array_bag += $obj_bag

                           Write-DetailedJobStatisLog $obj_bag  $DetailedFullPath  

                       } # foreach ($history in $history_Array )       
          
   
          }# If ($obj.JobHistory)

           else { #if the job does not have history
                          
                           $obj_bag = New-Object -TypeName PSObject -Property $bag_prop 
                           $obj_bag.Name = $obj.name
                           $obj_bag.SelectionSummary =  $obj.SelectionSummary 
                           $obj_bag.Storage =  $obj.Storage 
                           $obj_bag.StartTime = "N/A"
                           $obj_bag.ElapsedTime = "N/A" 
                           $obj_bag.JobStatus = "N/A"
                           $obj_bag.Medialabel = "N/A"
                           $obj_bag.TotalDataSizeBytes = "N/A"
                           $obj_bag.JobRateMBPerMinute = "N/A"
                           $array_bag += $obj_bag    

                           Write-DetailedJobStatisLog $obj_bag  $DetailedFullPath  
   
            }#end else  
   
           
 
 
        }#foreach ($obj in $array_obj)

        Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Number of Job Instances inspected : $($array_bag.count )"
        _status "          Number of Job Instances inspected : $($array_bag.count )     " 1  

    #endregion create array of custom objects



#endregion Module Process


#region Module GUI

    $HTMLReport = Join-Path $ScriptFilesPath  "BEJobHTML.Html" 

    $HTML = _drawme $array_bag

    $HTML | Out-File $HTMLReport


#endregion Module GUI


#region Module Closure
    
    
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "  
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "  
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Module Closure"
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "  
    
    
    if ($SendMail) {
            Write-Output ""
            _status " 4. Send Email" 2
            $subject = "BackupExec Job Report _$(Get-Date -f 'yyyy-MM-dd')"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Sending Email :"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - To : $MailTo"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - From : $MailFrom "
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - SMTP Host : $MailServer"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Subject : $subject"        
            _sendEmail $MailFrom $MailTo $subject $MailServer $ScriptFilesPath $InfoFullPath $ErrorFullPath
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Skipping Send Email as (SendEmail) parameter was not supplied"  

        }   
    
    
    _screenFooter $ScriptFilesPath


    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "  
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Script Ends"
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): " 

    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): $('-' * 60)"
    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Script Ends"
    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): " 

    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): $('-' * 60)"
    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Script Ends"
    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): "

#endregion Module Closure
       
       
   


















