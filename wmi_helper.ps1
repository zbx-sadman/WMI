<#                                          
    .SYNOPSIS  
        Return instances of HP Insight Management WBEM objects as LLD-JSON for Zabbix
        Also returns value of specified by index item of array that fetched by WMI query

    .DESCRIPTION
        Return instances of HP Insight Management WBEM objects as LLD-JSON for Zabbix
        Also returns value of specified by index item of array that fetched by WMI query

    .NOTES  
        Version: 0.9.1
        Name: WMI helper
        Author: zbx.sadman@gmail.com
        DateCreated: 
        Testing environment: HP DL360e, Windows Server 2008R2 SP1, Powershell 2.0;
                             HP DL360p, Windows Server 2012 R2, PowerShell 4.

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with WMI objects :
            Discovery - Make Zabbix's LLD JSON;
            Get       - Get value of metric or array item.

    .PARAMETER Namespace
        Namespace to WMI query

    .PARAMETER Query
        WMI Query

    .PARAMETER Idx
        Array index

    .PARAMETER ConsoleCP
        Codepage of Windows console. Need to properly convert output to UTF-8

    .PARAMETER DefaultConsoleWidth
        Say to leave default console width and not grow its to $CONSOLE_WIDTH

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        powershell.exe -NoProfile -ExecutionPolicy "RemoteSigned" -File "wmi_helper.ps1" -Action "Get" -Namespace "ROOT\HPQ" -Query "select OperationalStatus from HP_Processor where DeviceID='Proc 2'" -Idx "0" -Verbose -defaultConsoleWidth 

        Description
        -----------  
        Return Value of first item of OperationalStatus array that fetched from HP_Processor class of ROOT\HPQ namespace.
        Verbose messages is enabled.

    .EXAMPLE 
        ... "wmi_helper.ps1" -Action "Get" -Namespace "ROOT\HPQ" -Query "select * from HP_Processor" -defaultConsoleWidth 

        Description
        -----------  
        Show all instances (and its metrics) that is obtained on query execution.
#>


Param (
   [Parameter(Mandatory = $False)] 
   [ValidateSet('Discovery', 'Get')]
   [string]$Action,
   [Parameter(Mandatory = $True)]
   [String]$Namespace,
   [Parameter(Mandatory = $True)]
   [String]$Query,
   [Parameter(Mandatory = $False)]
   [String]$Idx,
   [Parameter(Mandatory = $False)]
   [String]$ErrorCode,
   [Parameter(Mandatory = $False)]
   [String]$ConsoleCP,
   [Parameter(Mandatory = $False)]
   [Switch]$DefaultConsoleWidth
);

#Set-StrictMode -Version Latest

# Set US locale to properly formatting float numbers while converting to string
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"
# Width of console to stop breaking JSON lines
Set-Variable -Option Constant -Name "CONSOLE_WIDTH" -Value 512

####################################################################################################################################
#
#                                                  Function block
#    
####################################################################################################################################

#
#  Define names of params in class instances to be use its as LLD macros
#
Function Define-LLDMacros {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject
   );
   $Class = $InputObject[0].__CLASS;
   # return array of parameter names, that will be used as LLD macros
   Switch ($Class) {
     #  'HP_DiskDrive' { @("DEVICEID", "NAME"); }
     #  'HP_Processor' { @("DEVICEID", "NAME"); }
     'HP_MemoryModule'         { @("NAME", "MANUFACTURER", "TAG"); }
     # HP_EthernetPort is HP_WinEthernetPort class in real
     'HP_WinEthernetPort'      { @("DEVICEID", "CAPTION", "PORTTYPE"); }
     # HP_EthernetPort is HP_WinEthernetPort  class in real
     'HP_WinNumericSensor '    { @("DEVICEID", "NAME", "UPPERTHRESHOLDCRITICAL", "NUMERICSENSORTYPE"); }
     'HPSA_ArrayController'    { @("NAME", "ELEMENTNAME"); }
      Default                  { @("DEVICEID", "NAME"); }
   }  
}

#
#  Prepare string to using with Zabbix 
#
Function PrepareTo-Zabbix {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [String]$ErrorCode,
      [Switch]$NoEscape,
      [Switch]$JSONCompatible
   );
   Begin {
      # Add here more symbols to escaping if you need
      $EscapedSymbols = @('\', '"');
      $UnixEpoch = Get-Date -Date "01/01/1970";
   }
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         If ($Null -Eq $Object) {
           # Put empty string or $ErrorCode to output  
           If ($ErrorCode) { $ErrorCode } Else { "" }
           Continue;
         }
         # Need add doublequote around string for other objects when JSON compatible output requested?
         $DoQuote = $False;
         Switch (($Object.GetType()).FullName) {
            'System.Boolean'  { $Object = [int]$Object; }
            'System.DateTime' { $Object = (New-TimeSpan -Start $UnixEpoch -End $Object).TotalSeconds; }
            Default           { $DoQuote = $True; }
         }
         # Normalize String object
         $Object = $( If ($JSONCompatible) { $Object.ToString().Trim() } else { Out-String -InputObject (Format-List -InputObject $Object -Property *)});

         If (!$NoEscape) { 
            ForEach ($Symbol in $EscapedSymbols) { 
               $Object = $Object.Replace($Symbol, "\$Symbol");
            }
         }

         # Doublequote object if adherence to JSON standart requested
         If ($JSONCompatible -And $DoQuote) { 
            "`"$Object`"";
         } else {
            $Object;
         }
      }
   }
}

#
#  Convert incoming object's content to UTF-8
#
Function ConvertTo-Encoding ([String]$From, [String]$To){  
   Begin   {  
      $encFrom = [System.Text.Encoding]::GetEncoding($from)  
      $encTo = [System.Text.Encoding]::GetEncoding($to)  
   }  
   Process {  
      $bytes = $encTo.GetBytes($_)  
      $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)  
      $encTo.GetString($bytes)  
   }  
}

#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
Function Make-JSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [array]$ObjectProperties, 
      [Switch]$Pretty
   ); 
   Begin   {
      [String]$Result = "";
      # Pretty json contain spaces, tabs and new-lines
      If ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } Else { $CRLF = $Tab = $Space = ""; }
      # Init JSON-string $InObject
      $Result += "{$CRLF$Space`"data`":[$CRLF";
      # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
      $itFirstObject = $True;
   } 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) {
         # Skip object when its $Null
         If ($Null -Eq $Object) { Continue; }

         If (-Not $itFirstObject) { $Result += ",$CRLF"; }
         $itFirstObject=$False;
         $Result += "$Tab$Tab{$Space"; 
         $itFirstProperty = $True;
         # Process properties. No comma printed after last item
         ForEach ($Property in $ObjectProperties) {
            If (-Not $itFirstProperty) { $Result += ",$Space" }
            $itFirstProperty = $False;
            $Result += "`"{#$Property}`":$(PrepareTo-Zabbix -InputObject $Object.$Property -JSONCompatible)";
         }
         # No comma printed after last string
         $Result += "$Space}";
      }
   }
   End {
      # Finalize and return JSON
      "$Result$CRLF$Tab]$CRLF}";
   }
}

$Result = 0;
If ([string]::IsNullOrEmpty($Query)) { 
   Write-Verbose "$(Get-Date) No query given";
   exit; 
}

#If ([string]::IsNullOrEmpty($NameSpace)) { $NameSpace = "ROOT\HPQ"; }

If ([string]::IsNullOrEmpty($Idx)) { $Idx = 0; }


Write-Verbose "$(Get-Date) Taking instances...";

# Prepare object lists

$Objects = Get-WmiObject -Computer "." -Query $Query -NameSpace $NameSpace;
#$Objects | fl *
#exit

Write-Verbose "$(Get-Date) Object(s) fetched, begin processing its with action: '$Action'";
$Result = $(
   # if no object in collection: 1) JSON must be empty; 2) 'Get' must be able to return ErrorCode
   Switch ($Action) {
      'Discovery' {
         # Discovery given object, make json for zabbix
         Write-Verbose "$(Get-Date) Class of object(s) is '$($Objects[0].__CLASS)'";
         $ObjectProperties = Define-LLDMacros -InputObject $Objects;
         Write-Verbose "$(Get-Date) Generating LLD JSON";
         Make-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
      }
      'Get' {
         # Select value if single metric or array's item
         If ($Null -ne $Objects) { 
            Write-Verbose "$(Get-Date) Select value of single object or array item that selected by WMI query";
            $Objects = $( If (1 -lt $Objects.Length) {
                             $Objects;
                          } ElseIf ($($Objects.Properties).IsArray) {
                             $($Objects.Properties).Value[$Idx];
                          } else {
                             $($Objects.Properties).Value;
                         }
            );
         }
         PrepareTo-Zabbix -InputObject $Objects -ErrorCode $ErrorCode;
      }
   }
);

# Convert string to UTF-8 if need (For Zabbix LLD-JSON with Cyrillic chars for example)
if ($consoleCP) { 
   Write-Verbose "$(Get-Date) Converting output data to UTF-8";
   $Result = $Result | ConvertTo-Encoding -From $consoleCP -To UTF-8; 
}

# Break lines on console output fix - increase console width to $CONSOLE_WIDTH chars
if (!$defaultConsoleWidth) { 
   Write-Verbose "$(Get-Date) Changing console width to $CONSOLE_WIDTH";
   mode con cols=$CONSOLE_WIDTH; 
}

Write-Verbose "$(Get-Date) Finishing";

$Result;
