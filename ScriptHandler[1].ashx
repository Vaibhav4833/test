option explicit
'=== Constant to Show or Hide Error messages
Const SHOW_SM_ERROR_MESSAGES                                = "False"
Const SM_SMART_MESSAGING_CLIENT_DB_NAME                     = "SmartMessaging.db"
Const SM_UIPROVIDER_ENUM_OBJECT                             = "SMUIMgrEnum"
Const SM_PROP_UIMGR_NAME                                    = "2"
Const SM_PROP_PROVIDER_APPLICATION_MGR                      = "6"
Const SM_PROP_PROVIDER_SUBSCRIPTION_MGR                     = "5"
Const SM_PROVIDER_NAME                                      = "2"
Const MC_PROP_APPLICATION_ID                                = "1"
Const MC_PROP_APPLICATION_NAME                              = "2"
Const MC_PROP_APPLICATION_VERSION                           = "4"
Const MC_PROP_APPLICATION_INSTALL_LCID						= "5"
Const MC_PROP_SUB_MGR_OBJ                                   = "0"
Const SM_PROP_PROVIDER_ENUM_CLIENTDB_OBJECT                 = "3"
Const SM_PROP_PROVIDER_ENUM_SYSTEM_MGR                      = "4"
Const SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT                 = "5"
Const SM_PROP_PROVIDER_ENUM_FILESYSTEM_OBJECT               = "6"
Const SM_CLIENTDB_OBJECT_CREATION_FAILED    				= "0"
Const SM_CLIENTDB_OBJECT_CREATION_SUCCESS					= "1"
Const SM_CLIENTDB_VERSION_LOCATION					        = "HKLM\Software\McAfee\MSM"  ' Pre-WSS 12.0
Const SM_CLIENTDB_VERSION_KEY_NAME					        = "SMVersion"
Const SM_VERSION_LOCATION       					        = "HKLM\Software\McAfee\MSM"  ' Pre-WSS 12.0
Const SM_REG_LOCATION                                       = "HKLM\Software\McAfee\Platform\MSM" ' WSS12.0+
Const SM_WEBUPDATE_LOCATION                                 = "HKLM\software\McAfee\MSC\Update\Webupdate"
Const SM_VERSIONID_KEY_NAME                                 = "VerID"
Const SM_VERSION_KEY_NAME       					        = "SMVersion"
Const SM_OOBE_LOCATION       					            = "HKLM\Software\McAfee\MSC\OEM\Manage\RGW\ActWiz"
Const SM_OOBE_KEY_NAME       					            = "OOBEModel"
Const SM_OOBE_ACTIVATIONDATE_KEY_NAME       				= "SvrActDate"
Const SM_DEFAULT_VERSION                                    = "1"
Const SM_NO_VERSION                                         = ""
Const SM_APP_VERSION_SPLITTER                               = "."
Const SM_VERSIONMGR_OBJECT_CREATION_SUCCESS                 = "1"
Const SM_VERSIONMGR_OBJECT_CREATION_FAILED                  = "0"
Const SM_CLIENTDB_DEFAULT_VERSION                           = "1"
Const SM_CLIENTDB_NO_VERSION                                = ""
Const SM_CLIENTDB_MISSING_DB                                = "0"
Const SM_CLIENTDB_DB_OPEN                                   = "1"
Const SM_CLIENTDB_QUERY_EXEC_SUCCESSFUL                     = "1"
Const SM_CLIENTDB_QUERY_EXEC_FAIL                           = "0"
Const SM_CLIENTDB_ALL_TABLES_CREATED_SUCCESS                = "1"
Const SM_CLIENTDB_ALL_TABLES_CREATED_FAILED                 = "0"
Const SM_CLIENTDB_VERSION_VALIDATION_SUCCESS                = "1"
Const SM_CLIENTDB_VERSION_VALIDATION_FAILED                 = "0"
Const SM_DEFAULT_SYNC_FLAG_VALUE                            = "0"
Const SM_CLIENTDB_SYNC_MESSAGE_URL                          = "http://us.mcafee.com"
Const SM_MESSAGE_LOG_TABLE_NAME                             = "MessageLog"
Const SM_MESSAGE_LOG_COLUMNS                                = "MessageViewHistoryID, ActionTypeID,Actionvalue, ActionURL,SyncFlag,CreatedDate"
Const SM_MESSAGE_SYNC_LOG_TABLE_NAME                        = "MessageSyncLog"
Const SM_MESSAGE_SYNC_LOG_COLUMNS                           = "SyncURL, ResultTypeID, NumMessage,SyncFlag, CreatedDate"
Const SM_MESSAGE_VIEW_HISTORY_TABLE_NAME                    = "MessageViewHistory"
Const SM_MESSAGE_VIEW_HISTORY_COLUMNS                       = "MessageID,Category,Priority,SyncFlag,AcctID,Affid,LCID,ProductKey,createddate"
Const SM_CLIENTDB_OPERATION_OK                              = 1
Const SM_CLIENTDB_OPERATION_FAILED                          = 2
Const SM_CLIENTDB_ACCESS_DENIED                             = 3
Const SM_CLIENTDB_NO_MORE_DATA                              = 4
Const SM_CLIENTDB_DATA_FOR_CURRENT_ROW_READY                = 5

Const SM_PROVIDER_OBJECT_CREATION_SUCCESS                   = 1
Const SM_PROVIDER_OBJECT_CREATION_FAILED                    = 0
Const SM_GET_INSTALLED_APPS_SUCCESS                         = 1
Const SM_GET_INSTALLED_APPS_FAILED                          = 0
Const SM_APPINFO_OBJECT_CREATION_SUCCESS                    = 1
Const SM_APPINFO_OBJECT_CREATION_FAILED                     = 0
Const SM_SUBMGR_OBJECT_CREATION_SUCCESS                     = 1
Const SM_SUBMGR_OBJECT_CREATION_FAILED                      = 0
Const SM_SYSTEMMGR_PROVIDER                                 = "SysMgrObject"
Const SM_SYSMGR_OBJECT_CREATION_SUCCESS                     = 1
Const SM_SYSMGR_OBJECT_CREATION_FAILED                      = 0
Const SM_UIMGR_NETWORKDRIVES_EXIST                          = "Yes"
Const SM_UIMGR_NETWORKDRIVES_MISSING                        = "No"
Const SM_SYSMGR_ALLOWED_IDLE_TIME_IN_MSEC                    = 60000
Const SM_SYSMGR_SYSTEM_IDLE                                 = 0
Const SM_SYSMGR_SYSTEM_IN_USE                               = 1

Const SM_SYSTEM_IN_SCREEN_SAVER_MODE                        = true
Const SM_SYSTEM_NOT_IN_SCREEN_SAVER_MODE                    = false

Const SM_SYSTEM_IN_FULL_SCREEN_MODE                        = true
Const SM_SYSTEM_NOT_IN_FULL_SCREEN_MODE                    = false
Const SM_OS_IS_WIN8_OR_GREATER                             = True

Const SM_REGISTRY_OBJECT_CREATION_SUCCESS                   = 1
Const SM_REGISTRY_OBJECT_CREATION_FAILED                    = 0

Const SM_TASKSCHEDULER_OBJECT_CREATION_SUCCESS              = 1
Const SM_TASKSCHEDULER_OBJECT_CREATION_FAILED               = 0
Const SM_TASKSCHEDULER_TASK_CREATION_SUCCESS                = 1
Const SM_TASKSCHEDULER_TASK_CREATION_FAILED                 = 0

Const SM_SUBMGR_APPID                                       = "appid"
Const SM_SUBMGR_AFFID                                       = "affid"
Const SM_SUBMGR_APPCODE                                     = "app_code"
Const SM_SUBMGR_ACCTID                                      = "accnt_id"
Const SM_SUBMGR_PRODUCTKEY                                  = "product_key"
Const SM_SUBMGR_EXPIREDATE                                  = "settings"
Const SM_SUBMGR_EXTENDEDEXPIREDATE                          = "extended_expiry_dt"
Const SM_SUBMGR_PERPETUAL                                   = "perpetual"
Const SM_SUBMGR_TRIAL                                       = "trial"
Const SM_SUBMGR_BACKEND                                     = "backend"
Const SM_SUBMGR_SYNCURL                                     = "sync_url"
Const SM_SUBMGR_WEBSITEURL                                  = "website"
Const SM_SUBMGR_LANGCODE                                    = "applang"
Const SM_SUBMGR_PKGID                                       = "package_id"
Const SM_SUBMGR_PKGNAME                                     = "package_name"
Const SM_SUBMGR_RENEWALURL                                  = "renewalurl"
Const SM_SUBMGR_FLAGS                                       = "flags"
Const SM_SUBMGR_EMAIL                                       = "e_mail"
Const SM_SUBMGR_LICENSECOUNT                                = "num_licenses"
Const SM_SUBMGR_INSTALLDT                                   = "installed"

Const SM_TASKSCHEDULER_SHELL                                = "WScript.Shell"
Const SM_TASKSCHEDULER_NAME                                 = "Task1"
Const SM_TASKSCHEDULER_TRIGGER_FREQUENCY                    = "3600"
Const SM_TASKSCHEDULER_TASK_PROCESS_PATH                    = "c:\\My.exe"
Const SM_TASKSCHEDULER_TASK_PROCESS_PARAMS                  = "http://sm.mcafee.com/syncmessage.aspx"
Const SM_TASKSCHEDULER_ACTIVETIMEBIAS_PATH                  = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
Const SM_SUBINFO_TOTAL_KNOWN_APPDATA_PROP_COUNT		        = "1"
Const SM_SHOW_MESSAGE_LIMIT                                 = "1"
Const SM_MAINT_FLAG_ENABLED                                 = False
Const SM_DB_HIT_LIMIT                                       =   1

Const SM_UICONTAINER_DEFAULT_HEIGHT                         = "600"
Const SM_UICONTAINER_DEFAULT_WIDTH                          = "450"
Const SM_UICONTAINER_OBJECT_CREATION_SUCCESS                = 1
Const SM_UICONTAINER_OBJECT_CREATION_FAILED                 = 0
Const SM_UICONTAINER_CENTERED                               = 0
Const SM_UICONTAINER_BOTTOM_RIGHT_CORNER                    = 1
Const SM_UICONTAINER_MANUAL                                 = 2

Const SM_INSTRUMENTATION_OBJECT_CREATION_FAILED             = 0
Const SM_INSTRUMENTATION_OBJECT_CREATION_SUCCESS            = 1

Const SM_INSTRU_STATUS_FAILED				                = False
Const SM_INSTRU_STATUS_SUCCESS				                = True
Const SM_GETUTCDTTIME_FAILED				                = "0"
Const SM_GETUTCDTTIME_SUCCESS				                = "1"
Const SM_GETUTC_DATE_TIME_FN_RESPONSE_HANDLER		        = "GetServerUTCDtTime"
Const SM_GETUTC_DATE_TIME_FN_REQUEST_URL		            = "https://sm.mcafee.com/InstruCommunicator.asmx/GetUTCTime"
Const SM_SEND_INSTRU_XML_FN_RESPONSE_HANDLER		        = "GetInstruResponse"
Const SM_SEND_INSTRU_XML_FN_RESPONSE_TIMEOUT		        = 5000
Const SM_SEND_INSTRU_XML_FN_REQUEST_PARAM_NAME    	        = "InstrumentationXML="
Const SM_SEND_INSTRU_XML_FN_REQUEST_URL		    	        = "https://sm.mcafee.com/InstruCommunicator.asmx/SetReportingInstru"
Const SM_SEND_SMARTINSTRU_XML_FN_REQUEST_URL		    	        = "https://sm.mcafee.com/InstruCommunicator.asmx/SetSmartReportingInstru"

Const SM_MACHINE_CONFIG_OS                                  = "os="
Const SM_MACHINE_CONFIG_OS_VERSION_ID                       = " osv="
Const SM_MACHINE_CONFIG_MEM_SIZE_MB                         = " MemorySizeInMB="
Const SM_MACHINE_CONFIG_PROCESSOR_FMLY                      = " ProcessorFamily="
Const SM_MACHINE_CONFIG_PROCESSOR_VRSN                      = " ProcessorVersion="
Const SM_MACHINE_CONFIG_PROCESSOR_SPEED_GHZ                 = " ProcessorSpeedGHZ="
Const SM_MACHINE_CONFIG_NO_OF_PROCESSORS                    = " NumberOfProcessors="
Const SM_MACHINE_CONFIG_NETWORK_DRIVES = " NetworkDrives="
Const SM_MACHINE_CONFIG_CSP_CLIENT_ID = " CspClientId="
Const SM_MACHINE_CONFIG_HARDWARE_ID                         = " HardwareId="
Const SM_MACHINE_CONFIG_SOFTWARE_ID                         = " SoftwareId="

Const SM_CLIENT_CONFIG_APPID                                = "app_id="
Const SM_CLIENT_CONFIG_APPCODE                              = " app_code="
Const SM_CLIENT_CONFIG_PERPETUAL                            = " perpetual="
Const SM_CLIENT_CONFIG_TRIAL                                = " trial="
Const SM_CLIENT_CONFIG_ACCTID                               = " acct_id="
Const SM_CLIENT_CONFIG_EXP_DT                               = " expire_dt="
Const SM_CLIENT_CONFIG_BACKEND                              = " backend="
Const SM_CLIENT_CONFIG_SYNC_URL                             = " sync_url="
Const SM_CLIENT_CONFIG_WEB_SITE                             = " web_site="
Const SM_CLIENT_CONFIG_AFFID                                = " aff_id="
Const SM_CLIENT_CONFIG_BUILDID                              = " build_id="
Const SM_CLIENT_CONFIG_LANG_CODE                            = " lang_code="
Const SM_CLIENT_CONFIG_PKGID                                = " pkg_id="
Const SM_CLIENT_CONFIG_PKGNM                                = " pkg_name="
Const SM_CLIENT_CONFIG_RNWL_URL                             = " renewal_url="
Const SM_CLIENT_CONFIG_REG_STATUS                           = " reg_status="
Const SM_CLIENT_CONFIG_EMAIL                                = " e_mail="
Const SM_CLIENT_CONFIG_INSTALLED_DT                         = " installed_date="
Const SM_CLIENT_CONFIG_LIC_QTY                              = " lic_qty="
Const SM_CLIENT_CONFIG_PROD_KEY                             = " product_key="
Const SM_CLIENT_CONFIG_INSTALLED_DYS                        = " installed_days="
Const SM_CLIENT_CONFIG_ACTIVATION_DYS                       = " activation_days="
Const SM_CLIENT_CONFIG_ACTIVATION_DATE                       = " activation_date="
Const SM_CLIENT_CONFIG_OOBE                                 = " oobe="
Const SM_CLIENT_CONFIG_IPT                                  = " ipt="
Const SM_CLIENT_CONFIG_VERSIONID                            = " versionid="
Const SM_CLIENT_CONFIG_LOCALEID                             = " locale_id="
Const SM_CLIENT_CONFIG_EULA                                 = " eula="
Const SM_CLIENT_CONFIG_RCODE                                = " rcode="
Const SM_CLIENT_CONFIG_MSC_BUILD                            = " msc_build="
Const SM_CLIENT_OOBE_CODE                                   = "3"
Const SM_CLIENT_CONFIG_FORCEUPGRADE                         = " force_upgrade=" 'Added for force upgrade 
Const SM_CLIENT_CONFIG_DAYSAFTERINSTALL                     = " days_after_install=" 'Added for SM Optimization 
Const SM_CLIENT_CONFIG_ACTIVATIONGRACEPERIOD                = " activation_grace_period=" ' Added for OOBE V3 change
Const SM_CLIENT_CONFIG_UNREGEXPIRED                         = " unreg_expired=" ' Added for OOBE V3 change
Const SM_CLIENT_CONFIG_UNREGEXPIREINDAYS                    = " unreg_expire_in_days=" ' Added for OOBE V3 change
Const SM_CLIENT_CONFIG_EXTENDEDEXPIRYDATE                   = " extended_expiry_dt=" ' Added for Windows 8 change
Const SM_CLIENT_CONFIG_DISABLEDDATE                         = " disabled_in_days=" ' Added for Windows 8 change
Const SM_CLIENT_CONFIG_SAFMESSAGE                           = " sa_message="        ' Site Advisor Freemium Message Applicable
Const SM_CLIENT_CONFIG_COHORTID                           = " cohort_definition_id="  

'=== Sub mgmt II constants
Const SM_PROP_PROVIDER_VERSION                              = "1"
Const SM_SUBMGMT2_OBJ_PROVIDER_VERSION                      = "2.0"
Const SM_PROP_PROVIDER_SUBMGMT2                             = "7"
Const SM_GET_SUBMGMT2_SUCCESS                               = 1
Const SM_GET_SUBMGMT2_FAILED                                = 0
''==== Object Creation
Const SM_OBJECT_CREATION_FAILED    			                = "0"
Const SM_OBJECT_CREATION_SUCCESS				            = "1"

''==== MSC IPT
Const SM_PROP_PROVIDER_MSCIPT                               = "8"
Const PLATFORMDSP                                           = "PlatformDspObj"
Const PRODUCT_KEY                                           = "pk"
Const IPTVALIDTILL                                          = "HKEY_CURRENT_USER\Software\MCAFEE\MSC\IPTValidTill"

'==== OOBE V3
Const SM_OOBE_V3_LOCATION       					        = "HKLM\Software\McAfee\OemInfo\Activation\OneClick\QueryParams"
Const SM_CLIENT_OOBE_V3_CODE                                = "8"
Const SM_CLIENT_OOBE_V3_VERSION                             = "OCVersion"
Const SM_CLIENT_OOBE_V3_ENABLE                              = "OCEnable"
Const SM_CLIENT_OOBE_V3_EULASTATE                           = "EULAState"
Const SM_CLIENT_OOBE_V3V4_EULADATE                          = "EulaDt"
Const SM_CLIENT_OOBE_V3_VERSION_NUMBER                      = "3.0"
Const SM_CLIENT_APPINFO_LOCATION                            = "HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute\QueryParams"
Const SM_CLIENT_OOBE_V3_IPRDONE                             = "IPRdone"

'===== AA: Site Advisor Freemium Project Variables
Const SM_CLIENT_SAFREEMIUM_LOCATION                         = "HKLM\SOFTWARE\McAfee\SiteAdvisor"
Const SM_CLIENT_SAF_VERSION                                 = "FullVersion"
Const SM_CLIENT_SAINSTALL_FOLDER                            = "Install Dir"
Const SM_CLIENT_SAFINSTALLER_FILE                           = "saUi.exe"
Const SM_CLIENT_SAFUICNTARGS                                = " /saFreemium /postinstallpage"
Const SM_CLIENT_SAFSMOFFER_STATUS                           = "IsSAFreemiumOfferAccepted"
'===== AA: Site Advisor Freemium Project Variables end

'Read release code
Const SM_CLIENT_WSS_RELEASECODE  							= "rcode"
Const SM_CLIENT_MSC_LOCATION                                = "HKLM\SOFTWARE\McAfee\MSC"
 

'==== Smart Messaging Optimization
Const SM_CLIENT_UICNTPATH_LOCATION                          = "HKLM\SOFTWARE\McAfee\MSC\OEM\Manage\RGW"
Const SM_CLIENT_UI_CONTAINER_PATH                           = "UICntPath"

''=== Virus Alert
Const REG_PATH_MSC_SECURITYREPORTS = "HKLM\Software\McAfee\MSC\Settings\Securityreports"
Const REG_KEY_THREATMETER = "threatMeter"

'==== Windows 8 Metro
Const REG_KEY_DISABLEDDATE                                  = "disabled_after_uac_date"

Const SM_PROP_METROTOASTYPE                                 = "toasttype"
Const SM_PROP_METROMSGHEADER                                = "msgheader"
Const SM_PROP_METROMSGBODY                                  = "msgbody"
Const SM_PROP_ISMETROVALID                                  = "isvalid"

Const SM_SUBMGR_METRODISPLAYMODE_UNKNOWN = 0
Const SM_SUBMGR_METRODISPLAYMODE_DESKTOP = 1
Const SM_SUBMGR_METRODISPLAYMODE_METRO = 2
Const SM_SUBMGR_METRODISPLAYMODE_STARTUP = 3

Const SM_SUBMGR_METROCLOSEREASON_UNDEFINED = -1
Const SM_SUBMGR_METROCLOSEREASON_CLICKED = 0
Const SM_SUBMGR_METROCLOSEREASON_DISMISSED = 1
Const SM_SUBMGR_METROCLOSEREASON_TIMEOUT = 2
Const SM_SUBMGR_METROCLOSEREASON_ERROR = 3

Const SM_METRO_TEMPLATE_UNDEFINED = -1
Const SM_METRO_TEMPLATE_TOASTTEXT01 = 4
Const SM_METRO_TEMPLATE_TOASTTEXT02 = 5

' A/B Testing before the A/B Framework was implemented.
Const SM_CLIENT_MESSAGEID_LOCATION                         = "HKLM\SOFTWARE\Mcafee\MSC\overrides\B6229FD2-AB40-40ac-B70A-68E0303E0546\Experimentation" 	'
Const SM_CLIENT_MESSAGEID_KEY		                       = "MessageID"
Const SM_CLIENT_INPRODUCT_LOCATION                         = "HKLM\SOFTWARE\McAfee\MSC\Settings\InProductTransaction"
Const SM_CLIENT_IPT_ENABLE_KEY                             = "enable"
Const SM_CLIENT_IPT_LCID                                   = "lcid"
Const SM_CLIENT_APPINFO_LOCATION_SUBSTITUTE                = "HKLM\SOFTWARE\McAfee\MSC\AppInfo\Substitute"
Const SM_CLIENT_IPT_BUILD                                  = "build"
Const SM_CLIENT_INPRODUCT_UPDATE_VERSION                   = "UpdateVersion"
Const SM_CLIENT_INPRODUCT_SUBSTATUS                        = "SubStatus"
Const SM_CLIENT_INPRODUCT_PARENTAFFID                      = "ParentAffId"

' Cohort registry

Const SM_CLIENT_COHORT_SM_LOCATION = "HKLM\SOFTWARE\Mcafee\MSC\overrides\B6229FD2-AB40-40ac-B70A-68E0303E0546\Experimentation\cohort\smid"
Const SM_CLIENT_COHORT_RULE_LOCATION = "HKLM\SOFTWARE\Mcafee\MSC\overrides\B6229FD2-AB40-40ac-B70A-68E0303E0546\Experimentation\cohort\cohortapplied"
Const SM_CLIENT_COHORT_ID = "cohortexecuted"

' ParterUniqueID
Const SM_CLIENT_PARTNER_DATA_INTERNAL_LOCATION = "HKLM\SOFTWARE\McAfee\OEM"
Const SM_CLIENT_PARTNER_DATA_STATUS_FLAG_REG_KEY = "PartnerDataUpdateStatus"
Const SM_CLIENT_PARTNER_DATA_EXTERNAL_LOCATION_01 = "HKLM\SOFTWARE\WOW6432Node\Staples\WO"
Const SM_CLIENT_PARTNER_DATA_EXTERNAL_LOCATION_02 = "HKLM\SOFTWARE\WOW6432Node\Staples\Order"
Const SM_CLIENT_PARTNER_DATA_EXTERNAL_LOCATION_03 = "HKLM\SOFTWARE\Staples\WO"
Const SM_CLIENT_PARTNER_DATA_EXTERNAL_LOCATION_04 = "HKLM\SOFTWARE\Staples\Order"
Const SM_CLIENT_PARTNER_DATA_EXTERNAL_KEY = "String Value"
Const SM_CLIENT_PARTNER_DATA_UPDATE_COMPLETED_FLAG = "3"

Const SM_CLIENT_INSTALLSETTINGS_LOCATION_SUBSTITUTE        = "HKLM\SOFTWARE\McAfee\VirusScan\InstallSettings\Substitute"
Const SM_CLIENT_CONTENT_VERSION                            = "ContentVersion"
Const SM_CLIENT_MSC_LOCATION_UPDATE                        = "HKLM\SOFTWARE\McAfee\MSC\Update"
Const SM_AMINDEX_URL_KEY                                   = "AMIndex_url" 
Const SM_LOG_HANDLER_URL		                           = "https://sm.mcafee.com/LogHandler.ashx?"
Const SM_CLIENT_INTERNET_EXPLORER_LOCATION                 = "HKLM\SOFTWARE\Microsoft\Internet Explorer"
Const SM_CLIENT_INTERNET_EXPLORER_VERSION                  = "VERSION"











    Dim oXMLDoc, oXMLHTTP
    Dim m_UTCDateTime, m_InstruStatus
 '=== This will get the size of the Providerobject array
    Function Length(ByRef vArray)
        Dim intIndex
        Dim intCount
    
        intCount = 0
        For intIndex = 0 To UBound(vArray)
            If Not(IsNull(vArray(intIndex))) And Not(IsEmpty(vArray(intIndex)))Then
                If Not(vArray(intIndex) Is Nothing) Then
                    intCount = intCount + 1
                End If    
            End If
        Next
        Length = intCount
    End Function
    '=== This will be used to show/hide the msgbox.
    '=== it will be used 
    Function ShowErrorMessage(strErrorMessage)
        Dim blnShowMsg
        blnShowMsg  = SHOW_SM_ERROR_MESSAGES
        If CBool(blnShowMsg) Then
            MsgBox(strErrorMessage)
        End If
    End Function
            
    Function HideWindow()
        Dim objExternal,objUICont
	    Set objExternal = window.external
	    Set objUICont = objExternal.GetParam ("McUIContainer")
	    Call objUICont.Hide() 
		Call objUICont.Terminate()
		'window.close
    End Function
    
    Function GetInternetConnectedState(objProviderEnum)
        Dim intResult
        Dim bInternetState
        Dim objSMSystemData
        
        Set objSMSystemData = New CSMSystemData
        bInternetState    = false
        
        If Not(objProviderEnum Is Nothing) Then
            If  Not(objSMSystemData Is Nothing) Then
                intResult   = objSMSystemData.InitializeSystemObject(objProviderEnum)
                If intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS Then
                    bInternetState = objSMSystemData.IsConnectedToInternet
                Else    
                    ShowErrorMessage "GetInternetConnectedState:objSMSystemData initialization failed"
	            End If
	        Else
	            ShowErrorMessage "GetInternetConnectedState:objSMSystemData is  empty"  
            End If
        Else
            ShowErrorMessage "GetInternetConnectedState:objProviderEnum is  empty"  
        End If
        
        Set objSMSystemData = Nothing
        
        GetInternetConnectedState = bInternetState
     End Function
    	
    Function MakeHTTPRequest(RequestURL, ResponseHandlerFunc, strEnvelope)
        On Error Resume Next
		set oXMLHTTP = createobject("msxml2.xmlhttp.3.0")
        set oXMLDoc = createobject("msxml2.domdocument")
        
        ShowErrorMessage("strEnvelope = " & strEnvelope)
        ShowErrorMessage("Calling WebService To Get UTC Time")
        oXMLHTTP.onreadystatechange = getRef(ResponseHandlerFunc)
       
        If Len(Trim(RequestURL)) > 0 Then
            call oXMLHTTP.open("POST",RequestURL,false)
            call oXMLHTTP.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
            If Len(Trim(strEnvelope)) > 0 Then
               setTimeout "AbortRequest(oXMLHTTP)",SM_SEND_INSTRU_XML_FN_RESPONSE_TIMEOUT
            End If
            call oXMLHTTP.send(strEnvelope)
            
        End If
    End Function
    
    Function AbortRequest(oXMLHTTP)
        ShowErrorMessage("Before Aborting")
        oXMLHTTP.abort '=== To abort the httprequest if it takes long time to respond
        m_InstruStatus = SM_INSTRU_STATUS_FAILED
        ShowErrorMessage("After Aborting")
    End Function

    Function SendInstruXML(strInstruXML)
         m_InstruStatus = SM_INSTRU_STATUS_FAILED
        dim strEnvelope : strEnvelope = SM_SEND_INSTRU_XML_FN_REQUEST_PARAM_NAME & strInstruXML
        Call MakeHTTPRequest(SM_SEND_INSTRU_XML_FN_REQUEST_URL, SM_SEND_INSTRU_XML_FN_RESPONSE_HANDLER, strEnvelope)
		SendInstruXML = m_InstruStatus
    End Function

    Function SendSmartInstruXML(strInstruXML)
         m_InstruStatus = SM_INSTRU_STATUS_FAILED
        dim strEnvelope : strEnvelope = SM_SEND_INSTRU_XML_FN_REQUEST_PARAM_NAME & strInstruXML
        Call MakeHTTPRequest(SM_SEND_SMARTINSTRU_XML_FN_REQUEST_URL, SM_SEND_INSTRU_XML_FN_RESPONSE_HANDLER, strEnvelope)
		SendSmartInstruXML = m_InstruStatus
    End Function

    Function GetInstruResponse()
        If(oXMLHTTP.readyState = 4) Then
            dim szResponse: szResponse = oXMLHTTP.responseText
            call oXMLDoc.loadXML(szResponse)
            If(oXMLDoc.parseError.errorCode <> 0) Then
                call ShowErrorMessage(oXMLDoc.parseError.reason)
		        m_InstruStatus = SM_INSTRU_STATUS_FAILED
            else
                m_InstruStatus = oXMLDoc.getElementsByTagName("boolean")(0).childNodes(0).text
                call ShowErrorMessage(oXMLDoc.getElementsByTagName("boolean")(0).childNodes(0).text)
            End If
        End If
	
    End Function
    
    Function GetUTCDateTime()
        Dim RequestURL
        '===RequestURL = "https://sm.mcafee.com/smartservice/smartmessaginginstrumentation.asmx/GetUTCTime"
        Call MakeHTTPRequest(SM_GETUTC_DATE_TIME_FN_REQUEST_URL, SM_GETUTC_DATE_TIME_FN_RESPONSE_HANDLER, "")
        GetUTCDateTime = m_UTCDateTime
    End Function

    Sub GetServerUTCDtTime()
        If(oXMLHTTP.readyState = 4) Then
            dim szResponse: szResponse = oXMLHTTP.responseText
            call oXMLDoc.loadXML(szResponse)
            If(oXMLDoc.parseError.errorCode <> 0) Then
                call ShowErrorMessage(oXMLDoc.parseError.reason)
		m_UTCDateTime = SM_GETUTCDTTIME_FAILED
            else
                m_UTCDateTime = oXMLDoc.getElementsByTagName("string")(0).childNodes(0).text
                call ShowErrorMessage(oXMLDoc.getElementsByTagName("string")(0).childNodes(0).text)
            End If
        End If
    End Sub

   Private Function GetQueryStringValueFromURL(strActionURL,qrystr)
        Dim intCIDPos, qryVal
        Dim intEqualPos, intAmpPos
        Dim intUrlLen
        qryVal = ""
        ShowErrorMessage "getcid"
        '===temporary
       '' Dim testurl: testurl = "http://home.mcafee.com/root/campaign.aspx?cid=45856&abc=1"
        If Len(Trim(strActionURL)) > 0 Then
            intCIDPos = Instr(1,strActionURL,qrystr) '=== get the cid position
            If CInt(intCIDPos) > 0 Then
                intEqualPos = Instr(intCIDPos,strActionURL,"=") '=== get the position of equal to
                If intEqualPos > 0 Then
                    If Instr(intEqualPos+1, strActionURL, "&", 1) > 0 Then 
                        intAmpPos = Instr(intEqualPos, strActionURL, "&", 1)    
                        qryVal = Mid(strActionURL, intEqualPos+1, (intAmpPos - (intEqualPos+1))) ' retrieve the variable value
                    Else
                        intUrlLen = Len(strActionURL)
                        qryVal = Mid(strActionURL, intEqualPos+1, (intUrlLen - (intEqualPos))) ' retrieve the variable value                            
                    End If    
                End If
            End If
        End If
        
        GetQueryStringValueFromURL = qryVal
    End Function
    
    Function GetCurrentDate(strServerUTCTime)
        Dim strFinalCurrentDate
        Dim strDay, strMonth, strYr
        
        If Len(Trim(strServerUTCTime)) > 0 Then
            strFinalCurrentDate = strServerUTCTime
        Else
            strFinalCurrentDate = Date
        End If
        
        
        strDay = FormatTwoDigit(Day(strFinalCurrentDate))
        strMonth = FormatTwoDigit(Month(strFinalCurrentDate))
        strYr = Year(strFinalCurrentDate)
        
    
        strFinalCurrentDate = strMonth & "/" & strDay & "/" &  strYr
    
        GetCurrentDate = strFinalCurrentDate
    End Function
    
    Function FormatTwoDigit(strValue)
        If Len(Trim(strValue)) = 1 Then
            strValue = "0" & strValue
        End If    
        FormatTwoDigit = strValue   
    End Function

    Function GetBrowserVersion()
        Dim regEx, userAgent, CurrentMatches
		Set regEx = new RegExp
		userAgent = Navigator.UserAgent 
		userAgent = Mid(userAgent, InStr(userAgent, "MSIE ") + 5)
		regEx.Pattern = "(\d+\.\d+)"
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.MultiLine = True
		Set CurrentMatches = regEx.Execute(userAgent)
		If CurrentMatches.Count > 0 Then
			If IsNumeric(CurrentMatches(0)) Then
				GetBrowserVersion = CDbl(CurrentMatches(0))
				Exit Function
			End If
		End If
		GetBrowserVersion = 6
    End Function

    Function GetIptValidTill()
        Dim l_ObjShell
        Dim l_IptValidTill
        set l_ObjShell = CreateObject(SM_TASKSCHEDULER_SHELL) 
        On Error Resume Next
        l_IptValidTill = l_ObjShell.RegRead(IPTVALIDTILL) 
        If Err.Number <> 0 Then
            Err.Clear
            GetIptValidTill = ""
        Else
            GetIptValidTill = l_IptValidTill
        End If
     End Function


Class CSMProviderEnum

Private m_ProvArray()
Private m_objExternal
Private m_objProvEnum

    Private Sub Class_Initialize
        On Error Resume Next 
    
	    Dim objProv
	    Dim ProviderArrayLen
	    
	    Set m_objExternal   = Nothing
	    Set m_objProvEnum   = Nothing
	    Set objProv         = Nothing
	    
	    ReDim m_ProvArray(0)
	    
	    ProviderArrayLen    = 0
	    
	    If m_objExternal Is Nothing Then
	     	If IsObject( window.external ) And Not(window.external Is Nothing) Then
	            Set m_objExternal = window.external
	        Else
	            ShowErrorMessage("SMProviderEnum:Unable to get the window.external")  
	        End if
	    End If 
	    
	    If IsObject(m_objExternal) And Not m_objExternal Is Nothing Then
            Set m_objProvEnum = m_objExternal.GetParam("SMProviderEnum")
	    End If
	    
	    If IsObject(m_objProvEnum) And Not m_objProvEnum Is Nothing Then
	        m_objProvEnum.Reset
	        
	        Do
                Set objProv = Nothing
                Set objProv = m_objProvEnum.Next()
                
                If Not (objProv Is Nothing) Then
                    ProviderArrayLen = Length(m_ProvArray)
                    If ProviderArrayLen > 0 Then
	                    ReDim Preserve m_ProvArray(ProviderArrayLen) 
	                End if
	                Set m_ProvArray(ProviderArrayLen) = objProv  
	            End If
            Loop Until objProv Is Nothing
	        
            If Err.number > 0 Then
               If Length(m_ProvArray) = 0 Then
                ShowErrorMessage("SMProviderEnum Initialize: " & Err.Description) 
               End If
            End If


''''	        For Each objProv In m_objProvEnum
''''	            If objProv.IsEnabled()Then
''''	                ProviderArrayLen = Length(m_ProvArray)
''''	                If ProviderArrayLen > 0 Then
''''	                    ReDim Preserve m_ProvArray(ProviderArrayLen) 
''''	                End if
''''	                 m_ProvArray(ProviderArrayLen) = objProv       
''''	                
''''	            End If
''''	        Next
        Else
            ShowErrorMessage("SMProviderEnum:Unable to get the Provider Enum object")
	    End If
	End Sub

	Public Property Get ProviderSize()
		ProviderSize = Length(m_ProvArray)
	End Property
	
	Public Property Get ProviderEnumObject()
	    Set ProviderEnumObject = m_objProvEnum    
	End property
	
	Public Function GetProviderObject(ByRef intRetVal, intObjNumber)
		    On Error Resume Next
		    
		    Dim objNxt
		    Set objNxt = Nothing
		    
		    If Length(m_ProvArray) > 0 Then
		        Set objNxt      = m_ProvArray(intObjNumber)
		        intRetVal       = SM_PROVIDER_OBJECT_CREATION_SUCCESS
		    Else
		        ShowErrorMessage("SMProviderEnum:Unable to get the Provider object")
		        intRetVal = SM_PROVIDER_OBJECT_CREATION_FAILED    
		    End If
		    
		    Set GetProviderObject = objNxt
	End Function
	
   Private Sub Class_Terminate
	
    	If Not IsEmpty(m_objProvEnum) Then
    	    If Not m_objProvEnum Is Nothing Then
	            set m_objProvEnum = nothing
	        End If
	    End If
	
	    If Not IsEmpty(m_objExternal) Then
	        If Not m_objExternal Is Nothing Then
	            set m_objExternal = nothing
	        End If
	    End If
	
	End Sub


End Class


Class CSMSystemData
    
    Private m_objExternal
    Private m_objSystemInfo
    Private m_arrProcessorInfo
    Private m_NetworkDrivesExist
    Private m_LastInputInfoInSec
    Private m_intNoOfProcessors
    Private m_strProcessorFamily    '=== eg: intel (vendor)
    Private m_strProcessorVersion   '=== eg: CPU Name "Centrino"
    Private m_strProcessorSpeed
    Private m_blnNetworkDrives
    
    
    Private Sub Class_Initialize
        Set m_objExternal     = window.external
        Set m_objSystemInfo   = Nothing
       
    End Sub
    
    Public Property Get IsConnectedToInternet()
        Dim bInternetState: bInternetState = false
        
        bInternetState    = m_objSystemInfo.InternetConnectedState()
        IsConnectedToInternet   = bInternetState
     End Property	 

     Public Property Get MachineOsId()
        On Error Resume Next

        Dim strMachineOSNumber : strMachineOSNumber = m_objSystemInfo.GetOSVersion()
        If Len(Trim(strMachineOSNumber)) > 0 And IsNumeric(strMachineOSNumber) Then
            MachineOsId = CInt(strMachineOSNumber)
        Else
            MachineOsId = 0
        End If
     End Property
    
     Public Property Get MachineOS()
        on error resume next
        Dim strMachineOSNumber, strMachineOS
        strMachineOSNumber    = m_objSystemInfo.GetOSVersion()
        If Len(Trim(strMachineOSNumber)) > 0 Then
            Select Case strMachineOSNumber
                Case "1"
                    strMachineOS = "Windows 95"
                Case "2"
                    strMachineOS = "Windows 98"
                Case "3"
                    strMachineOS = "Windows NT"
                Case "4"
                    strMachineOS = "Windows 2000"
                Case "5"
                    strMachineOS = "Windows ME"
                Case "6"
                    strMachineOS = "Windows XP 32bit"
                Case "7"
                    strMachineOS = "Windows Vista 32bit"
                Case "8"
                    strMachineOS = "Windows XP 64bit"
                Case "9"
                    strMachineOS = "Windows Vista 64bit"
                Case "10"
                    strMachineOS = "Windows 7 32bit"
                Case "11"
                    strMachineOS = "Windows 7 64bit"
                Case "12"
                    strMachineOS = "Windows8_32bit"
                Case "13"
                    strMachineOS = "Windows8_64bit"
                Case "14"
                    strMachineOS = "Windows8_32bit" ' Actually Windows8.1_32bit
                Case "15"
                    strMachineOS = "Windows8_64bit" ' Actually Windows8.1_64bit
                Case "16"
                    strMachineOS = "Windows10_32bit"
                Case "17"
                    strMachineOS = "Windows10_64bit"
                Case else
                    strMachineOS = ""
            End Select
        
            MachineOS   = strMachineOS
        Else
            MachineOS = ""
        End If
     End Property
     
     Public Property Get MachineHardwareID()
        Dim strHardwareID
        strHardwareID    = m_objSystemInfo.GetMachineID()	    
        'strHardwareID    = m_objSystemInfo.GetProperty(MC_PROP_SYSTEMINFO_HARDWAREID)
        If Len(Trim(strHardwareID)) > 0 Then
            MachineHardwareID   = strHardwareID
        Else
            MachineHardwareID = ""
        End If
     End Property

     Public Property Get MachineSoftwareID()
        Dim strSoftwareID
        strSoftwareID    = m_objSystemInfo.GetWindowsID()
        'strSoftwareID    = m_objSystemInfo.GetProperty(MC_PROP_SYSTEMINFO_SOFTWAREID)
        If Len(Trim(strSoftwareID)) > 0 Then
            MachineSoftwareID   = strSoftwareID
        Else
            MachineSoftwareID = ""
        End If
     End Property	 

     Public Property Get MachineMemorySize()  '=== Check whether we get that in MB, as API needs that in MB
        on error resume next
        Dim strMachineMemory
        strMachineMemory    = m_objSystemInfo.GetSystemMemorySize()
        If Len(Trim(strMachineMemory)) > 0 Then
            MachineMemorySize   = GetMemorySizeInMB(strMachineMemory)
        Else
            MachineMemorySize = ""
        End If
     End Property     
     
     Public Property Get NumberOfProcessors()
        If m_intNoOfProcessors > 0 Then
            NumberOfProcessors   = m_intNoOfProcessors
        Else
            NumberOfProcessors = 0
        End If
     End Property	 
     
     Public Property Get ProcessorFamily()
        If Len(Trim(m_strProcessorFamily)) > 0 Then
            ProcessorFamily   = m_strProcessorFamily
        Else
            ProcessorFamily = ""
        End If
     End Property		 

     Public Property Get ProcessorVersion()
        If Len(Trim(m_strProcessorVersion)) > 0 Then
            ProcessorVersion   = m_strProcessorVersion
        Else
            ProcessorVersion = ""
        End If
     End Property	
     
     Public Property Get ProcessorSpeedGHZ()
        Dim strProcSpeedGHZ
        strProcSpeedGHZ = ""
        If Len(Trim(m_strProcessorSpeed)) > 0 Then
            strProcSpeedGHZ = GetSpeedInGHZ(m_strProcessorSpeed)
            ProcessorSpeedGHZ   = strProcSpeedGHZ
        Else
            ProcessorSpeedGHZ = 0
        End If
     End Property	
     
     Public Property Get NetworkDrivesExist()
        NetworkDrivesExist  = m_NetworkDrivesExist
     End Property

     Public Property Get IsSystemInUse()
        Dim intLastInputInfoInMSec
        Dim intSystemInUse, intResult
        Dim objSMUIContainer
         
        Set objSMUIContainer    = New CSMUIContainer
        intLastInputInfoInMSec   = 0
        If Not objSMUIContainer is Nothing Then
            intResult = objSMUIContainer.InitializeUIContainerObject()
            If intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS Then
                intLastInputInfoInMSec = objSMUIContainer.GetLastInputInfo  
            End If
        End If    
        intSystemInUse          = SM_SYSMGR_SYSTEM_IDLE
        ShowErrorMessage "intLastInputInfoInMSec is " & intLastInputInfoInMSec
        If intLastInputInfoInMSec > 0 Then
            ShowErrorMessage "intLastInputInfoInMSec is " & intLastInputInfoInMSec
            If intLastInputInfoInMSec < SM_SYSMGR_ALLOWED_IDLE_TIME_IN_MSEC Then
                intSystemInUse  = SM_SYSMGR_SYSTEM_IN_USE
            Else
                intSystemInUse  = SM_SYSMGR_SYSTEM_IDLE
            End If
        Else
            intSystemInUse  = SM_SYSMGR_SYSTEM_IN_USE
        End If
        ShowErrorMessage "intSystemInUse is " & intSystemInUse
        IsSystemInUse = intSystemInUse
     End Property    

     Public Function InitializeSystemObject(objProviderEnum)
      
        On Error Resume Next
        
        Dim ProviderEnumObject 
        Dim intResult, retSubInfo
        Dim testArray
        intResult = 0
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objSystemInfo = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_SYSTEM_MGR)
            
            If Not (m_objSystemInfo Is Nothing) Then
               
                testArray = m_objSystemInfo.GetProcessorInfo()
                m_arrProcessorInfo  = m_objSystemInfo.GetProcessorInfo()
                
                '=== Newly added code below
                Dim objSMUIContainer
                Set objSMUIContainer    = New CSMUIContainer
                
                If Not objSMUIContainer is Nothing Then
                    intResult = objSMUIContainer.InitializeUIContainerObject()
                    
                    If intResult = SM_UICONTAINER_OBJECT_CREATION_SUCCESS Then
                        m_LastInputInfoInSec    = objSMUIContainer.GetLastInputInfo  
                        m_NetworkDrivesExist    = objSMUIContainer.NetworkDrivesExist    
                    End If
                    
                End If    
                '=== Newly added code above
                If Length(m_arrProcessorInfo) = 0 Then
                    ShowErrorMessage("SMSystemData:Unable to get the Processor Info")
                    intResult = SM_SYSMGR_OBJECT_CREATION_FAILEDs
                Else
                    SetProcessorInfo()
                End If
                intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS
            End If
        Else
            intResult = SM_SYSMGR_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMAppData:Unable to get the Provider object")
        End If
        
        InitializeSystemObject = intResult 
    End Function
    
    Private Sub SetProcessorInfo()
        on Error Resume next
        
        Dim intNoOfProcessors
        Dim objProcessorInfo
        
        Set objProcessorInfo = Nothing
        intNoOfProcessors = Length(m_arrProcessorInfo)
        
        If intNoOfProcessors > 0 Then
            Set objProcessorInfo = m_arrProcessorInfo(0)
            m_intNoOfProcessors = intNoOfProcessors
            If Not (objProcessorInfo Is Nothing) Then
                m_strProcessorFamily    = objProcessorInfo.GetCPUFamily()
                m_strProcessorVersion   = objProcessorInfo.GetCPUModel()  
                m_strProcessorSpeed     = objProcessorInfo.GetCPUSpeed()  
            End If
        End If
    End Sub
    
    Public Function GetMemorySizeInMB(strMemorySizeInKB)
        Dim lMemorySizeMB
        Dim intMBEquiv
        lMemorySizeMB = 0
        intMBEquiv = 1024
        If Len(Trim(strMemorySizeInKB)) > 0 Then
            lMemorySizeMB = Round(CDBL(strMemorySizeInKB) / intMBEquiv)
        End If
        
        GetMemorySizeInMB = lMemorySizeMB
    End Function

    ' Added for Smart Messaging Optimization 
    Public Function LaunchApplication(strExePath, strCmdLine)
        Dim intResult
        
        intResult = window.external.LaunchApplication(cstr(strExePath), cstr(strCmdLine))

        LaunchApplication = intResult
    End Function

    Public Function LaunchSAApplication(objProviderEnum, strExePath, strCmdLine)
        On Error Resume Next
        
        Dim ProviderEnumObject 
        Dim intResult, retSubInfo
        intResult = 0
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objSystemInfo = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_SYSTEM_MGR)
            If Not (m_objSystemInfo Is Nothing) Then
                intResult  = m_objSystemInfo.RunProgram(strExePath, strCmdLine)
                If intResult <> 0 Then
                    ShowErrorMessage("SMSystemData:Unable to launch SA")
                    intResult = SM_SYSMGR_OBJECT_CREATION_FAILEDs
                End If
                intResult = SM_SYSMGR_OBJECT_CREATION_SUCCESS
            End If
        Else
            intResult = SM_SYSMGR_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMAppData:Unable to get the Provider object")
        End If
        
        LaunchSAApp = intResult
    End Function
    ' End of Smart Messaging Optimization 
    
    Private Function GetSpeedInGHZ(SpeedInMB)
        Dim lSpeedGHZ
        Dim lGHZEquiv
        lSpeedGHZ = 0
        lGHZEquiv = 1000
        
        lSpeedGHZ = Round(CDBL(SpeedInMB) / lGHZEquiv)
        GetSpeedInGHZ = lSpeedGHZ
    End Function
    
     Private Sub Class_Terminate
        If Not IsEmpty(m_objSystemInfo) Then
            If Not m_objSystemInfo Is Nothing Then
                Set m_objSystemInfo = Nothing
            End If
        End If
        If Not IsEmpty(m_objExternal) Then
            If Not m_objExternal Is Nothing Then
                Set m_objExternal = Nothing
            End If
        End If
    End Sub


End Class


Class CSMRegistry
    
    Private m_objRegistryMgr
    
    Private Sub Class_Initialize
        Set m_objRegistryMgr    = Nothing
    End Sub
    
    Public Function InitializeRegistryObject(objProviderEnum)
    On Error Resume Next
        ShowErrorMessage "Initialize Function Called"
        Dim ProviderEnumObject 
        Dim intResult
        
        intResult = 0
        Set ProviderEnumObject = objProviderEnum.ProviderEnumObject
        
        If Not(ProviderEnumObject Is Nothing) Then   
            Set m_objRegistryMgr = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)
            
            If Not (m_objRegistryMgr Is Nothing) Then
                intResult = SM_REGISTRY_OBJECT_CREATION_SUCCESS
                ShowErrorMessage("SMRegistry:Registry object created")
            Else
                intResult = SM_REGISTRY_OBJECT_CREATION_FAILED
                ShowErrorMessage("SMRegistry:Unable to get the Registry object")
            End if
        Else
            intResult = SM_REGISTRY_OBJECT_CREATION_FAILED
            ShowErrorMessage("SMRegistry:Unable to get the Provider object")
        End If
    
        InitializeRegistryObject = intResult
    End Function
    
    Private Property Get SMRegistryPath
        Dim bIsRegPathFound : bIsRegPathFound = False
        
        If Not m_objRegistryMgr Is Nothing Then
            bIsRegPathFound = m_objRegistryMgr.RegKeyPresent(SM_REG_LOCATION)
            If CBool(bIsRegPathFound) Then
                ' SM Registry location used by WSS 12.0 onward
                SMRegistryPath = SM_REG_LOCATION
            Else
                ' SM Registry location used previous to WSS 12.0
                SMRegistryPath = SM_CLIENTDB_VERSION_LOCATION
            End If            
        Else
            ShowErrorMessage("Unable to create Registry object") 
        End If
    End Property    
    
    Public Function GetClientSMVersion()
       Dim bResult, strSMVersion
       
       bResult = False
       strSMVersion = ""
            
            If Not m_objRegistryMgr Is Nothing Then
              bResult = m_objRegistryMgr.RegKeyPresent(SMRegistryPath)
               
               If CBool(bResult) Then
                   '=== TODO: Check whether this will get the Version value correctly.
                   '=== RegQueryValue(BSTR bstrKeyPath,BSTR bstrValue,VARIANT *pvValue) 
                    strSMVersion = m_objRegistryMgr.RegQueryValue(SMRegistryPath, SM_VERSION_KEY_NAME)
               Else
                    strSMVersion = SM_NO_VERSION
               End If
            Else   
               ShowErrorMessage("Unable to create Registry object")
            End If
        
        GetClientSMVersion = strSMVersion
    End Function

    Public Sub SetClientSMVersion(strSMVersionValue)
      
        Dim bCreateKey, bSetValue, bResult
        bCreateKey  = False
        bResult     = False
        bSetValue   = False
       
        If Not m_objRegistryMgr Is Nothing Then
            bResult = m_objRegistryMgr.RegKeyPresent(SMRegistryPath)
            If bResult Then
                '=== TODO: Check whether this parameter is passed correctly.
                '=== RegSetValue(BSTR bstrKeyPath,BSTR bstrKeyValueName,VARIANT vValue,BOOL *pbResult) 
                bSetValue = m_objRegistryMgr.RegSetValue(SMRegistryPath, SM_VERSION_KEY_NAME, Trim(strSMVersionValue))
                If CBool(bSetValue) Then
                    ShowErrorMessage("ResgistryValue for SMVersion is created")  
                Else
                    ShowErrorMessage("Unable to set Client SMVersion in Registry")
                End If
            Else
                ShowErrorMessage("MSM is missing in registry")
            End If
        Else
            ShowErrorMessage("Unable to create Registry object")       
        End If
    End Sub
    
    'ARPU Fix Start
    
    Public Sub SetRegKeyValueCstr(strKeyName,strValueName,strValue,bObfuscate)
        Dim bOldObfuscate,bSetValue
        bSetValue = False
        If Not m_objRegistryMgr Is Nothing Then
           bOldObfuscate = m_objRegistryMgr.Obfuscate
           m_objRegistryMgr.Obfuscate = bObfuscate
           bSetValue = m_objRegistryMgr.RegSetValue(strKeyName,strValueName, Trim(strValue))
          If CBool(bSetValue) Then
               ShowErrorMessage("Registry Entry " & strKeyName & " " & strValueName & " " & " is set to " & m_objRegistryMgr.RegQueryValue(strKeyName, strValueName) )
           Else
               ShowErrorMessage("Unable to set Registry entry " & strKeyName & " " & strValueName)
           End If            
           m_objRegistryMgr.Obfuscate = bOldObfuscate
        End If
    End Sub
    
    Public Function GetRegKeyValueName(strKeyName,strValueName,bObfuscate)
'        ShowErrorMessage("Registry Entry " & strKeyName & " " & strValueName)
        Dim bOldObfuscate,strValue,bResult
        strValue = ""
        bResult  = False
        If Not m_objRegistryMgr Is Nothing Then
            bOldObfuscate = m_objRegistryMgr.Obfuscate             
            m_objRegistryMgr.Obfuscate = bObfuscate
            strValue = m_objRegistryMgr.RegQueryValue(strKeyName, strValueName)
            ShowErrorMessage("Aff Id " & strValue)
            m_objRegistryMgr.Obfuscate = bOldObfuscate
'        Else
'           ShowErrorMessage("Registry Mgr is nothing")
        End If
'        ShowErrorMessage("Aff Id " & strValue)
        GetRegKeyValueName = strValue 
    End Function   	
    
    'ARPU Fix End  
    Public Function GetVersionId()
        Dim bResult, versionId
        bResult = False
        versionId = ""
                   
        If Not m_objRegistryMgr Is Nothing Then
            bResult = m_objRegistryMgr.RegKeyPresent(SM_WEBUPDATE_LOCATION)
               
            If CBool(bResult) Then
                versionId = m_objRegistryMgr.RegQueryValue(SM_WEBUPDATE_LOCATION,SM_VERSIONID_KEY_NAME)
            Else
                versionId = ""
            End If
        Else   
            ShowErrorMessage("Unable to create Registry object")
        End If

        GetVersionId = versionId

    End Function
    
     Private Sub Class_Terminate        
        If Not(m_objRegistryMgr is Nothing) Then
            Set m_objRegistryMgr    = Nothing
        End if
    End Sub   	

End Class


Dim m_objProviderEnum
Dim m_mscVersion
Dim m_vsoEngineVersion
Dim m_IsAppInfoSubstituteExists
Dim m_IsInstallSettingsSubstituteExists
Dim m_IsIELocationExists
Dim m_isMSCUpdateExists
Dim m_amIndexURLKeyValue  
Dim m_amIndexBetaUrl
Dim m_OSId
Dim int_IEVersion
Dim bOldObfuscate
Set m_objProviderEnum   = New CSMProviderEnum

Function PatchAMIndexUrl()
    On Error Resume Next
	
    Dim objSMSystemData, ieVersion  
    Set objSMSystemData = New CSMSystemData
    If Not(m_objProviderEnum Is Nothing) Then        
        
        intResult = objSMSystemData.InitializeSystemObject(m_objProviderEnum)
        m_OSId = objSMSystemData.MachineOsId            

        Dim m_objRegistryMgr, ProviderEnumObject
        Set ProviderEnumObject = m_objProviderEnum.ProviderEnumObject
        Set m_objRegistryMgr = ProviderEnumObject.GetProperty(SM_PROP_PROVIDER_ENUM_REGISTRY_OBJECT)
	    m_objRegistryMgr.Obfuscate = False

	    m_IsAppInfoSubstituteExists = FALSE
        m_IsAppInfoSubstituteExists = m_objRegistryMgr.RegKeyPresent(SM_CLIENT_APPINFO_LOCATION_SUBSTITUTE)
        If m_IsAppInfoSubstituteExists THEN
        m_mscVersion=m_objRegistryMgr.RegQueryValue(SM_CLIENT_APPINFO_LOCATION_SUBSTITUTE, SM_CLIENT_IPT_BUILD)
        End If

        m_IsInstallSettingsSubstituteExists = FALSE
        m_IsInstallSettingsSubstituteExists = m_objRegistryMgr.RegKeyPresent(SM_CLIENT_INSTALLSETTINGS_LOCATION_SUBSTITUTE)
        If m_IsInstallSettingsSubstituteExists THEN
        m_vsoEngineVersion=m_objRegistryMgr.RegQueryValue(SM_CLIENT_INSTALLSETTINGS_LOCATION_SUBSTITUTE, SM_CLIENT_CONTENT_VERSION)
        End If

        m_IsIELocationExists = FALSE
        m_IsIELocationExists=m_objRegistryMgr.RegKeyPresent(SM_CLIENT_INTERNET_EXPLORER_LOCATION)
        If m_IsIELocationExists THEN
        ieVersion=m_objRegistryMgr.RegQueryValue(SM_CLIENT_INTERNET_EXPLORER_LOCATION, SM_CLIENT_INTERNET_EXPLORER_VERSION)
        End If
        
        Dim mscVersion15, mscVersion1281005, mscVersion1281003, vsoEngineVersion3772, l_blockingStatus 
        Dim t_IEVersion
        t_IEVersion = Split(ieVersion, ".", -1)	
        If UBound(t_IEVersion)=0 Then
            If Len(Trim(ieVersion))>0 Then
				    int_IEVersion = CInt(Trim(ieVersion))
		    End If        
        ElseIf UBound(t_IEVersion)>0 Then
            If Len(Trim(t_IEVersion(0)))>0 Then
				    int_IEVersion = CInt(Trim(t_IEVersion(0)))
		    End If
        End If
        
        If Len(Trim(m_mscVersion)) > 0 And Len(Trim(m_vsoEngineVersion)) > 0 And Len(Trim(m_OSId)) > 0 And Len(Trim(int_IEVersion)) > 0 Then
            mscVersion15 = CompareDotDelimitedIntegerString(CStr(m_mscVersion), "15.0")
            mscVersion1281005 = CompareDotDelimitedIntegerString(CStr(m_mscVersion), "12.8.1005")
            mscVersion1281003 = CompareDotDelimitedIntegerString(CStr(m_mscVersion), "12.8.1003")
            vsoEngineVersion3772 = CompareDotDelimitedIntegerString(CStr(m_vsoEngineVersion), "3772.*")
            l_blockingStatus = GetBlockingStatus(int_IEVersion, m_OSId)

            If (mscVersion15 = 1 Or mscVersion15 = 0) Then
                'Do Nothing
            ElseIf (mscVersion1281005 = -1  Or mscVersion1281005  = 0) Then
                If (mscVersion1281003 = -1 And vsoEngineVersion3772 = -1) Then
                    UpdateToBeta(m_objRegistryMgr)
                ElseIf (l_blockingStatus = 1 Or l_blockingStatus = 0) Then
                    UpdateToBeta(m_objRegistryMgr)
                Else
                    'Do Nothing
                End If
            Else
                If (l_blockingStatus = 0 And (vsoEngineVersion3772 = -1 Or vsoEngineVersion3772 = 0)) Then
                    UpdateToBeta(m_objRegistryMgr)
                Else
                    'Do Nothing
                End If
            End If
        End If
    End If

    HideWindow()
    
End Function

Private Function UpdateToBeta(objRegistryMgr)    
    m_isMSCUpdateExists = FALSE
    m_isMSCUpdateExists = objRegistryMgr.RegKeyPresent(SM_CLIENT_MSC_LOCATION_UPDATE)
            
    If m_isMSCUpdateExists Then
	    bOldObfuscate = objRegistryMgr.Obfuscate	
	    objRegistryMgr.Obfuscate = True 
	    m_amIndexURLKeyValue= objRegistryMgr.RegQueryValue(SM_CLIENT_MSC_LOCATION_UPDATE, SM_AMINDEX_URL_KEY)
	    objRegistryMgr.Obfuscate = bOldObfuscate		
        If Len(Trim(m_amIndexURLKeyValue))>0 AND InStr(1, lcase(Trim(m_amIndexURLKeyValue)), lcase("http://download.mcafee.com/molbin/iss-loc/amcore/amindex/"))=1 AND StrComp(Right(Trim(m_amIndexURLKeyValue), 11),  "amindex.cab") = 0 THEN
			
            m_amIndexBetaUrl = Replace(m_amIndexURLKeyValue, "download.mcafee.com", "download.beta.mcafee.com") 
				
	        bOldObfuscate = objRegistryMgr.Obfuscate	
	        objRegistryMgr.Obfuscate = True 				
	        Call objRegistryMgr.RegSetValue(SM_CLIENT_MSC_LOCATION_UPDATE, SM_AMINDEX_URL_KEY, CStr(m_amIndexBetaUrl))
	        objRegistryMgr.Obfuscate = bOldObfuscate

	        Call LogHandler(m_mscVersion, m_vsoEngineVersion, m_amIndexURLKeyValue, m_OSId, int_IEVersion)				

        End If	
    End If	
End Function


Function MakeHTTPGETRequest(requestURL)
    On Error Resume Next
	set oXMLHTTP = createobject("msxml2.xmlhttp.3.0")
    set oXMLDoc = createobject("msxml2.domdocument")              
    call oXMLHTTP.open("GET", requestURL, false)
    call oXMLHTTP.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")	
    call oXMLHTTP.send() 
End Function

Function LogHandler(mscVersion, vsoEngineVersion, amIndexURLKeyValue, osId, ieVersion)
    Dim requestURL        
    Dim objSMSystemData
	Dim intResult
	intResult   = 0
    Set objSMSystemData = Nothing
    Set objSMSystemData = New CSMSystemData
	
	If  Not(objSMSystemData Is Nothing) Then
        intResult   = objSMSystemData.InitializeSystemObject(m_objProviderEnum)
	End If

    requestURL= SM_LOG_HANDLER_URL & "hid=" & objSMSystemData.MachineHardwareID() & "&sid=" & objSMSystemData.MachineSoftwareID() & "&build=" & mscVersion & "&cversion=" & vsoEngineVersion & "&osid=" & osId & "&ieversion=" & ieVersion & "&amiurl=" & amIndexURLKeyValue
	
    Call MakeHTTPGETRequest(requestURL)        
End Function

Function GetBlockingStatus(IEVersion, OSId)
    GetBlockingStatus = -1 'Error 
    
    Dim osBlockingStatus, ieBlockingStatus
    osBlockingStatus = 0
    ieBlockingStatus = 0
       
    ieBlockingStatus = GetIEBlockingStatus(IEVersion)
    osBlockingStatus = GetOSBlockingStatus(OSId)
    
    GetBlockingStatus = Max(ieBlockingStatus, osBlockingStatus)
End Function

Function Max(x, y)
    If x > y Then Max = x Else Max = y
End Function

Private Function GetIEBlockingStatus(IEVersion)
    'Blocked to legacy > Blocked to wss14.0r2 > No blocking
    '2 > 1 > 0
    GetIEBlockingStatus = 0 
    If (IEVersion < 8) Then
        GetIEBlockingStatus = 2
    ElseIf (IEVersion = 8) Then
        GetIEBlockingStatus = 1
    Else
        GetIEBlockingStatus = 0
    End If
End Function

Private Function GetOSBlockingStatus(OSId)
    'Blocked to legacy > Blocked to wss14.0r2 > No blocking
    '2 > 1 > 0
    GetOSBlockingStatus = 0 
    If IsLowerThanVista(OSId) Then
        GetOSBlockingStatus = 2
    ElseIf IsVista(OSId) Then
        GetOSBlockingStatus = 1
    Else
        GetOSBlockingStatus = 0
    End If
End Function

Private Function IsVista(OSId)
    IsVista = False
    If (OSId = 7 And OSId = 9) Then
        IsVista= True
    End If    
End Function

Private Function IsLowerThanVista(OSId)
    IsLowerThanVista = False
    If (OSId <= 6 Or OSId=8) Then
        IsLowerThanVista= True
    End If    
End Function


Function GetDotDelimitedIntegerStringAsArray(ByVal DotDelimitedIntegerString)
  Dim FullArray(), DelimitedStringArray, N, trimmedString  
  DelimitedStringArray = Split(DotDelimitedIntegerString, ".")
  Redim FullArray(UBound(DelimitedStringArray))
  For N = 0 To UBound(DelimitedStringArray)
	trimmedString = Trim(DelimitedStringArray(N))
  
    If Len(trimmedString)>0 Then
		If IsNumeric(trimmedString) Then
			FullArray(N) = trimmedString
		ElseIf trimmedString = "*" Then
			FullArray(N) = "*"
		Else
			FullArray(N) = ""
		End If
	Else
		FullArray(N) = ""
	End If
    
  Next 
  GetDotDelimitedIntegerStringAsArray = FullArray
End Function

' Compares two dot delimited strings "1.2.3.4". 
' If String1 < String2, returns -1. 
' If String1 = String2, returns  0.
' If String1 > String2, returns  1.
Function CompareDotDelimitedIntegerString(ByVal String1, ByVal String2)
  Dim String1Array, String2Array, Result, String1ArrayUBound, String2ArrayUBound, N
  String1Array = GetDotDelimitedIntegerStringAsArray(String1)
  String2Array = GetDotDelimitedIntegerStringAsArray(String2)
  
  String1ArrayUBound=UBound(String1Array)
  String2ArrayUBound=UBound(String2Array)
  
   For N = 0 To String2ArrayUBound
	If String1ArrayUBound >=N  Then	
		If IsNumeric(String1Array(N)) And IsNumeric(String2Array(N)) Then
			If(CInt(String1Array(N)) > CInt(String2Array(N))) Then
				Result = 1
                Exit For
			ElseIf(CInt(String1Array(N)) < CInt(String2Array(N))) Then
				Result = -1
				Exit For
			Else
				Result = 0
			End If			
		ElseIf (String2Array(N) = "*") Then
				Result = 0	
		ElseIf (String1Array(N) = "*" Or Not(Len(String1Array(N))>0)) Then
				Result = -1
				Exit For
		End If
	ElseIf IsNumeric(String2Array(N)) Then
		If CInt(String2Array(N))>0 Then		
			Result = 1				
		ElseIf (Not(Len(String2Array(N))>0)) Then
			Result = 1		
		End If
		Exit For
	End If    
	
   Next  
   
	If Result = 0 And String1ArrayUBound > String2ArrayUBound Then	
		For N = String2ArrayUBound To String1ArrayUBound
		If IsNumeric(String1Array(N)) Then
			If(CInt(String1Array(N)) > 0) Then
				Result = 1
				Exit For
			End If			
		ElseIf (String1Array(N) = "*") Then
				Result = 0	
		ElseIf (Not(Len(String1Array(N))>0)) Then
				Result = -1
				Exit For
		End If   
		
		Next
		
	End If
    
  CompareDotDelimitedIntegerString = Result
End Function


