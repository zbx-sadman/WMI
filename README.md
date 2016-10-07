#WMI Helper 

This is a little Powershell script that helps to use WMI with Zabbix

Actual release 0.9.1

Tested on:
- HP DL360e, Windows Server 2008R2 SP1, Powershell 2.0;
- HP DL360p, Windows Server 2012 R2, PowerShell 4.

Actions:
- _Discovery_ - Make Zabbix's LLD JSON from WMI instances, which fetched with WMI query;
- _Get_       - Get value of WMI object or array item;


###How to use standalone

        # Show value of first item of OperationalStatus array that fetched from HP_Processor class of ROOT\HPQ namespace.
        # Verbose output is enabled.
        powershell.exe -NoProfile -ExecutionPolicy "RemoteSigned" -File "wmi_helper.ps1" -Action "Get" -Namespace "ROOT\HPQ" -Query "select OperationalStatus from HP_Processor where DeviceID='Proc 2'" -Idx "0" -Verbose -defaultConsoleWidth 


###How to use with Zabbix
1. Just include [zbx\_wmi\_helper.conf](https://github.com/zbx-sadman/wmi/tree/master/Zabbix_Templates/zbx_wmi_helper.conf) to Zabbix Agent config;
2. Put _wmi\_helper.ps1_ to _C:\zabbix\scripts_ dir. If you want to place script to other directory, you must edit _zbx\_wmi\_helper.conf_ to properly set script's path; 
3. Set Zabbix Agent's / Server's _Timeout_ to more that 3 sec (may be 10 or 30);
4. Import [template(s)](https://github.com/zbx-sadman/wmi/tree/master/Zabbix_Templates) to Zabbix Server;
5. Be sure that Zabbix Agent worked in Active mode - in template can be used 'Zabbix agent(active)' poller type. Otherwise - change its to 'Zabbix agent' and increase value of server's StartPollers parameter;
6. Enjoy.

**Note #1**
Do not try import Zabbix v2.4 template to Zabbix _pre_ v2.4. You need to edit .xml file and make some changes at discovery_rule - filter tags area and change _#_ to _<>_ in trigger expressions. I will try to make template to old Zabbix.

**Note #2**
Make sure that all doublequotes is escaped, if its used in query string with Zabbix. Otherwise you will see error messages instead expected values.

###Hints
- Edit or add cases on Switch() in Define-LLDMacros() function to specify Zabbix's LLD macro names;
- To see all instances and its metrics just do right query with _Get_ command: `... "wmi_helper.ps1" -Action "Get" -Namespace "ROOT\HPQ" -Query "select * from HP_Processor" -defaultConsoleWidth`;
- You can use WMI Helper to bring values '1' or '0' instead Windows's 'True' or 'False' to Zabbix's Data Item. Example: `wmi.helper[get,root\hpq,select PoweredOn from HP_ProcessorChip where Name=\"Proc 1\"]`;
- To get on Zabbix Server side properly UTF-8 output when have non-english (for example Russian Cyrillic) symbols in Computer Group's names, use  _-consoleCP **your_native_codepage**_ command line option. For example to convert from Russian Cyrillic codepage (CP866): _... "wsus_miner.ps1"  ... -consoleCP CP866_;
- If u need additional symbol escaping in LLD JSON - just add one or more symbols to _$EscapedSymbols_ array in _PrepareTo-Zabbix_ function;
- To measure script runtime use _Verbose_ command line switch.
