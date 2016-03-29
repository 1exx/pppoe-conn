@echo off
rem
rem ########################################
rem ###### PPPoE Connection generator ######
rem ########################################
rem :: Need test on Win8
setlocal

rem :: -------------------------------------
rem :: pppName - name of created connection
rem :: pppCall - service name
set pppName=MyISP
set pppCall=myisp

rem :: -------------------------------------
rem :: Check Windows Version
rem :: 5.1 = XP
rem :: 6.0 = Vista or Server 2K8
rem :: 6.1 = Win7 or Server 2K8R2
rem :: 6.2 = Win8 or Server 2K12
rem :: 6.3 = Win8.1 or Server 2K12R2
rem :: 10.0 = Win10
rem :: 0.0 = Unknown or Unable to determine
rem :: --------------------------------------
for /f "tokens=4-5 delims=[.XP " %%i in ('ver') do set VERS=%%i.%%j
if "%vers%" == "10.0" set VERSION=10
if "%vers%" == "6.3"  set VERSION=9
if "%vers%" == "6.2"  set VERSION=8
if "%vers%" == "6.1"  set VERSION=7
if "%vers%" == "6.0"  set VERSION=6
if "%vers%" == "5.1"  set VERSION=5

rem :: Get device id for WAN Miniport (PPPoE)
rem :: seems it necessary to work
rem :: wmic nic where (name like "%PPPoE%" or name like "%PPPOE%") get deviceid /format:textvaluelist.xsl
for /f "tokens=2 delims==" %%i in ('wmic nic where ^(name like "%%PPPoE%%" or name like "%%PPPOE%%"^) get deviceid /format:textvaluelist.xsl') do set IDX=%%i

rem :: Creating Network Profile Folder to store dialer profile, if its not there already
if %version% == 5 (
	mkdir "%USERPROFILE%\Application Data\Microsoft\Network\Connections\Pbk" 2> nul
	cd "%USERPROFILE%\Application Data\Microsoft\Network\Connections\Pbk\"
) else (
	mkdir %userprofile%\AppData\Roaming\Microsoft\Network\Connections\pbk 2> nul
	cd "%userprofile%\AppData\Roaming\Microsoft\Network\Connections\Pbk\"
)

rem :: Creating pppoe profile
type nul > rasphone.pbk
(
echo [%pppName%]
echo Encoding=1

rem :: WinXP - none
rem :: Win7  - PBVersion=1
rem :: Win8  - PBVersion=2 ???
rem :: Win81 - PBVersion=3
rem :: Win10 - PBVersion=4
if %version% GTR 5 (
	if %version% GEQ 9 (
		if %version% GEQ 10 ( 
			echo PBVersion=4
		) else (
			echo PBVersion=3
		)
	) else (
 		echo PBVersion=1
 	)
)

echo Type=5
echo AutoLogon=0
echo UseRasCredentials=0

rem :: TODO: generate manually
rem :: LowDateTime=-582619296
rem :: HighDateTime=30327625
echo DialParamsUID=6326979
echo Guid=359A05BF4C146640949F56383A0F18F5

if %version% == 5 echo BaseProtocol=1

echo VpnStrategy=0

rem :: ExcludedProtocols=8 - disabled IPv6
if not %version% == 5 (
	echo ExcludedProtocols=8
) else (
	echo ExcludedProtocols=3
)

echo LcpExtensions=1
echo DataEncryption=8
echo SwCompression=1
echo NegotiateMultilinkAlways=0

if %version% == 5 (
	echo SkipNwcWarning=0
	echo SkipDownLevelDialog=0
)

echo SkipDoubleDialDialog=0

rem :: XP = 1
if %version% == 5 (
	echo DialMode=1
) else (
	echo DialMode=0
)

echo OverridePref=15
echo RedialAttempts=0
echo RedialSeconds=60
echo IdleDisconnectSeconds=0
echo RedialOnLinkFailure=1
echo CallbackMode=0
echo CustomDialDll=
echo CustomDialFunc=
echo CustomRasDialDll=

if not %version% == 5 (
	echo ForceSecureCompartment=0
	echo DisableIKENameEkuCheck=0
)

echo AuthenticateServer=0
echo ShareMsFilePrint=0
echo BindMsNetClient=0
echo SharedPhoneNumbers=0
echo GlobalDeviceSettings=0
echo PrerequisiteEntry=
echo PrerequisitePbk=

if %version% == 5 (
	echo PreferredPort=
	echo PreferredDevice=
) else (
	echo PreferredPort=PPPoE%idx%-0
	echo PreferredDevice=WAN Miniport ^(PPPOE^)
)

echo PreferredBps=0
echo PreferredHwFlow=0
echo PreferredProtocol=0
echo PreferredCompression=0
echo PreferredSpeaker=0
echo PreferredMdmProtocol=0
echo PreviewUserPw=1
echo PreviewDomain=0
echo PreviewPhoneNumber=0
echo ShowDialingProgress=1
echo ShowMonitorIconInTaskBar=1

rem:: MS-CHAPv2 only
rem:: WinXP: 768
rem:: Win7:  512
rem:: Win81: 512
rem:: Win10: 512
if %version% == 5 (
	echo CustomAuthKey=-1
	echo AuthRestrictions=768
	echo TypicalAuth=1
) else (
	echo CustomAuthKey=0
	echo AuthRestrictions=512
)

echo IpPrioritizeRemote=1

if not %version% == 5 echo IpInterfaceMetric=0

echo IpHeaderCompression=0
echo IpAddress=0.0.0.0
echo IpDnsAddress=0.0.0.0
echo IpDns2Address=0.0.0.0
echo IpWinsAddress=0.0.0.0
echo IpWins2Address=0.0.0.0
echo IpAssign=1
echo IpNameAssign=1

if %version% == 5 echo IpFrameSize=1006

echo IpDnsFlags=0
echo IpNBTFlags=0
echo TcpWindowSize=0

rem:: what does it mean?
echo UseFlags=3

echo IpSecFlags=0
echo IpDnsSuffix=

if %version% GTR 5 (
	echo Ipv6Assign=1
	echo Ipv6Address=::
	echo Ipv6PrefixLength=0
	echo Ipv6PrioritizeRemote=1
	echo Ipv6InterfaceMetric=0
	echo Ipv6NameAssign=1
	echo Ipv6DnsAddress=::
	echo Ipv6Dns2Address=::
	echo Ipv6Prefix=0000000000000000
	echo Ipv6InterfaceId=0000000000000000
	echo DisableClassBasedDefaultRoute=0
	echo DisableMobility=0
	echo NetworkOutageTime=0
	echo ProvisionType=0
	echo PreSharedKey=
	echo CacheCredentials=1
)

rem :: doesnt have win8 to test
if %version% GEQ 9 (
	echo NumCustomPolicy=0
	echo NumEku=0
	echo UseMachineRootCert=0
	echo NumServers=0
	echo NumRoutes=0
	echo NumNrptRules=0
	echo AutoTiggerCapable=0
	echo NumAppIds=0
	echo NumClassicAppIds=0
	if %version% GEQ 10 (
		echo SecurityDescriptor=
		echo ApnInfoProviderId=
		echo ApnInfoUsername=
		echo ApnInfoPassword=
		echo ApnInfoAccessPoint=
		echo ApnInfoAuthentication=1
		echo ApnInfoCompression=0
		echo DeviceComplianceEnabled=0
		echo DeviceComplianceSsoEnabled=0
		echo DeviceComplianceSsoEku=
		echo DeviceComplianceSsoIssuer=
		echo FlagsSet=0
	)
	echo DisableDefaultDnsSuffixes=0
	echo NumTrustedNetworks=0
	echo NumDnsSearchSuffixes=0
	echo PowershellCreatedProfile=0
	echo ProxyFlags=0
	echo ProxySettingsModified=0
	echo ProvisioningAuthority=
	echo AuthTypeOTP=0
	if %version% GEQ 10 (
		echo GREKeyDefined=0
		echo NumPerAppTrafficFilters=0
		echo AlwaysOnCapable=0
		echo PrivateNetwork=0
	)
)

echo.
echo NETCOMPONENTS=
echo ms_msclient=0
echo ms_server=0

if %version% == 5 echo ms_psched=1

echo.
echo MEDIA=rastapi
echo Port=PPPoE%idx%-0

rem :: maybe it's not necessary - test it later
if %version% == 5 (
	echo Device=WAN Miniport ^(PPPoE^)
) else (
	echo Device=WAN Miniport ^(PPPOE^)
)

echo.
echo DEVICE=PPPoE
echo PhoneNumber=%pppCall%
echo AreaCode=
echo CountryCode=0
echo CountryID=0
echo UseDialingRules=0
echo Comment=

rem:: no FriendlyName in WinXP
if not 	%version% == 5 echo FriendlyName=

echo LastSelectedPhone=0
echo PromoteAlternates=0
echo TryNextAlternateOnFail=1
echo.
) >> rasphone.pbk

rem :: -------------------------------------------------------------------
rem :: JS script: create shortcut, set as default connection
rem :: TODO: Shortcut for system object {BA126AD7-2166-11D1-B1D0-00805FC1270E}
rem :: -------------------------------------------------------------------
(
echo var CSIDL_CONNECTIONS = 0x31; // Network Connections
echo var pppName = "%pppName%";
echo var objShell = new ActiveXObject^("Shell.Application"^);
echo var objFolder = objShell.NameSpace^(CSIDL_CONNECTIONS^);
echo.
echo objShell.Open^(CSIDL_CONNECTIONS^); // Open Network Connections folder ^(like press F5^)
echo.
echo WSH.Sleep^(1000^);
echo.
echo /*
echo  close Network Connection folder when job is done?
echo */
echo.
echo if ^(objFolder != null^)
echo {
echo 	var objFolderItems = objFolder.Items^(^); // Network adapters list
echo 	if ^(objFolderItems != null^)
echo 	{
echo 		var nCount = objFolderItems.Count; // Network adapters count
echo 		for ^(i = 0; i ^< nCount; i++^){
echo 			if ^(objFolderItems.Item^(i^).name == pppName^)
echo 			{
echo 				var colVerbs = objFolderItems.Item^(i^).Verbs^(^); // Right click on adapter
echo 				var vCount = colVerbs.Count;
echo 				for ^(j = 0; j ^< vCount; j++^)
echo 				{
echo 					var cVerb = colVerbs.Item^(j^).name;
echo 					//WSH.echo^(cVerb^);
echo 					var shortcut = /^(shortcut^|\u044f\u0440\u043b\u044b\u043a^)^$/i; // "Create shortcut|Создать ярлык"
echo 					var defconn = /^(default^|\u0443\u043c\u043e\u043b\u0447\u0430\u043d\u0438\u044e^)^$/i; // "Make default|Сделать подключением по умолчанию"
echo.
echo 					if ^(shortcut.test^(cVerb^)^)
echo 						colVerbs.Item^(j^).DoIt^(^); // create shortcut
echo.
echo 					if ^(defconn.test^(cVerb^)^)
echo 						colVerbs.Item^(j^).DoIt^(^); // set default
echo 				}
echo 			}
echo 		}
echo 	}
echo }
echo.
echo WSH.Sleep^(5000^); // How long we display alert "Create shortcut"
echo.
) > tmp.js

WScript.exe tmp.js
del tmp.js

rem :: Run connection dialog
start rasphone -d %pppName%