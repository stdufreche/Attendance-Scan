;###################### BarcodeAttendance.ahk ###########################

Global version := "0.28.1"
#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

;########################################################################
;###### DEBUG POPUP WINDOW ##############################################
;########################################################################
DebugOutput(Level, Var1:="", Var2:="", Var3:="", Var4:="", Var5:="", Var6:="", Var7:="", Var8:="", Var9:="", Var10:="") {
    IF (DebugLevel>=Level) {
        MsgBox %Var1%%Var2%%Var3%%Var4%%Var5%%Var6%%Var7%%Var8%%Var9%%Var10%
    }
    ;Level 1 = Subroutines and Major Timers
    ;Level 2 = Variable Definitions
    ;Level 3 = Comm Port Debug & Frequent Timers
}

;########################################################################
;###### VERSION AND ENVIRONMENT CONTROL #################################
;########################################################################
{
UpdateURL:=False

oHttp := ComObjCreate("WinHttp.Winhttprequest.5.1")
httpsend=https://raw.githubusercontent.com/stdufreche/Attendance-Scan/refs/heads/main/version
oHttp.open("GET",httpsend)
oHttp.send()

currentVersion := StrSplit(oHttp.responseText, ".")
localVersion := StrSplit(version, ".")

;Check Major/Minor/Patch version numbers
Loop, 3
    {
        IF (currentVersion[A_INDEX]>localVersion[A_INDEX]) {
            UpdateURL=https://github.com/stdufreche/Attendance-Scan/releases
            Break
        } 
    }

;Check for existing configuration file and generate new if not found
cfgFile := FileOpen("LHScfg.ini", "r")
if !IsObject(cfgFile)
	CfgFileCreate()
    
global LogFile := FileOpen("Attendance.log", "a")

;Read configuration variables from file
IniRead, WebAppURL,     LHScfg.ini,    GoogleAPI,          DeploymentURL
IniRead, DefaultTrip,   LHScfg.ini,    HallPassAPI,        DefaultTrip
IniRead, DebugLevel,    LHScfg.ini,    GeneralSettings,    DebugLevel
IniRead, DisplayMonitor,LHScfg.ini,    GUIconfig,          DisplayMonitor
IniRead, RS232_Port,    LHScfg.ini,    CommSetting,        RS232_Port
IniRead, RS232_Baud,    LHScfg.ini,    CommSetting,        RS232_Baud
IniRead, RS232_Parity,  LHScfg.ini,    CommSetting,        RS232_Parity
IniRead, RS232_Data,    LHScfg.ini,    CommSetting,        RS232_Data
IniRead, RS232_Stop,    LHScfg.ini,    CommSetting,        RS232_Stop
IniRead, scanSuffix,    LHScfg.ini,    CommSetting,        scanSuffix
IniRead, BTenable,      LHScfg.ini,    CommSetting,        BTenable
IniRead, BTdeviceName,  LHScfg.ini,    CommSetting,        BTdeviceName
}

;########################################################################
;########################################################################
;######################## BEGIN MAIN PROGRAM ############################
;########################################################################
;########################################################################

Global Wrapper := A_Args.Length() ? True : False

global DebugLevel=DebugLevel
global WebAppURL:=WebAppURL
global ScanTimeout:=5000
global scanID:=
global EventLog:=
global scanData:=
global tripType=%DefaultTrip%
global BTenable:=%BTenable%
global BTdeviceName:="BarCode Scanner spp"
global queueCount:=0
global ScanLine:=[]
global BufferLength:=0
global QueueButton := "0 Scans Queued"




    
Switch Wrapper
{
    Case True:
        ;Construct main GUI
        Gui, 1:New,-MaximizeBox -Caption ;-SysMenu +ToolWindow 
        Gui Font, s10 Bold
        If (UpdateURL!=False)
            Gui Add, Link, h30 +0x1 +Center, <a href="%UpdateURL%">Download New Version</a>
        Else
            Gui Add, Text, w156 h30 +Center vScanStatus cGreen, ____________________
        ;Gui Add, Text, x20 y5 h33 +0x1, Com Port Interface
        ;Gui Font
        
        ;Gui Font, s9, Segoe UI
        ;Gui Add, StatusBar,,
        ;SB_SetText("v" version)
        ;Gui Show, w163 h115, CommWrapper
        Gui Show, w156 h30, CommWrapper

    Case False:
        Gui, 1:New,-MaximizeBox ;-SysMenu
        Gui Font, s9, Segoe UI
        Gui Font
        Gui Font, s10 Bold
        ;Gui Add, Text, x6 y5 w222 h33 +0x1, Scan ID# to log.
        Gui Font, s12 Bold
        Gui Add, Text, +Center vScanStatus cGreen, ____________________
        Gui Font, s9, Segoe UI
        Gui Add, StatusBar,,
        Gui Font
        Gui Font, s11 Bold
        Gui Add, Button, x49 y40 w136 h23 gManualSend vQueueButton, No Scans Queued
        If (UpdateURL!=False)
            Gui Add, Link, x36 y70 w200 h23 +0x1 +Center, <a href="%UpdateURL%">Download New Version</a>
        Gui Font
        Gui Show, w231 h115, ID Scan
        SB_SetText("v" version)
    
    default:
}

;########################################################################
;###### OSD for MultiMonitor Displays ###################################
;########################################################################
If (DisplayMonitor != 0)
{
    global CustomColor := "B7E1CD"
    SysGet, Mon, Monitor, %DisplayMonitor% ; 1 is the primairy monitor, 2 the secondairy etc. 
    bckgrdX := "W" Abs(MonRight - MonLeft)
    bckgrdY := "H" Abs(MonTop - MonBottom)
    Xpos := "X" MonLeft
    Ypos := "Y" MonTop
    textYpos := "Y" (Abs(MonTop - MonBottom) - 100)
    ;MsgBox, Left: %MonLeft% -- Top: %MonTop% -- Right: %MonRight% -- Bottom %MonBottom% -- bckgrdX: %bckgrdX% bckgrdY: %bckgrdY% textYpos: %textYpos%.

    ;GuiOSD()
    ;Construct main GUI
    ; Can be any RGB color (it will be made transparent below).
    Gui, OSD:New, +LastFound +AlwaysOnTop +ToolWindow -Caption -MinimizeBox -SysMenu ;+Resize  ; +ToolWindow avoids a taskbar button and an alt-tab menu item.
    Gui, Color, cGreen ;%CustomColor%
    Gui, Font, s32  ; Set a large font size (32-point).
    Gui, Add, Text, +Center %bckgrdX% %LargeGUIOffset% cLime
    ;Gui, Add, Text, +Center %bckgrdX% vClock cGreen,%A_Hour%:%A_Min%:%A_Sec%
    Gui, Add, Text, +Center %bckgrdX% X0 %textYpos% vVar1 cLime,_________________________________________
    ;Gui, Add, Text, +Center %bckgrdX% vVar2 cGreen,_________________________________________
    ;Gui, Add, Text, +Center %bckgrdX% vVar3 cGreen,_________________________________________
    ;Gui, Add, Text, +Center %bckgrdX% vVar4 cGreen,_________________________________________
    ; Make all pixels of this color transparent and make the text itself translucent (150):
    WinSet, TransColor, cGreen 250 ;%CustomColor% 250
    ;WinSet, Transparent, 150

    ;SetTimer, UpdateOSD, 200
    ;Gosub, UpdateOSD  ; Make the first update immediate rather than waiting for the timer.
    ;If LargeGUIOffset=="y0"
    Gui, Show, NoActivate %bckgrdX% %bckgrdY% %Xpos% %Ypos% ; NoActivate avoids deactivating the currently active window.
    ;Else
        ;Gui, Show, NoActivate xCenter %LargeGUIOffset% %bckgrdX% %bckgrdY% ; NoActivate avoids deactivating the currently active window.
    
    GuiControl, OSD:, MoveDraw
}

;########################################################################
;###### CHECK SERIAL PORT FOR DATA ######################################
;########################################################################
GuiControl,1:,ScanStatus,Scanner Startup
If (scanSuffix=="CR" || scanSuffix=="CRLF")
    scanSuffix:=Chr(13)
scanState = 0
scanData =
SetTimer, connectionTimer, 1000

return

Loop 
{
    ;scanLength:=ScanLine.Length()
    ;If DebugLevel>1
    ;    MsgBox ScanLine: %scanLength%
    ;IF (scanLength!=0)
    ;{
    ;ProcessScan(ScanLine[1], DefaultTrip1)
    ;ScanLine.RemoveAt(1)
    ;SB_SetText("Scan Queue: " ScanLine.Length())
    ;}
    Sleep 1000
}


Return
;########################################################################
;########################################################################
;########################## END MAIN PROGRAM ############################
;########################################################################
;########################################################################

;########################################################################
;###### ON GUI CLOSE ####################################################
;########################################################################
{
GuiClose:
RS232_Close(RS232_FileHandle)
LogFile.Close()

ExitApp
Exit
Return
}

;########################################################################
;###### SERIAL READ TIMER ###############################################
;########################################################################
{
SerialTimer:
If DebugLevel>2
    MsgBox SerialTimer Expired
;Prevent interruption during execution of this timed thread.
;Critical, On
;0xFF in the line below sets the size of the read buffer.
Read_Data := RS232_Read(RS232_FileHandle,"0xFF",RS232_Bytes_Received)
;Break the timer loop if serial port is closed
IF (Read_Data == False)
    Return
;Process the data, if there is any.
If (RS232_Bytes_Received > 0) {
    If DebugLevel>1
        SB_SetText("scanData:" %scanData1%)
    ;Begin Data to ASCII conversion
    ASCII =
    Read_Data_Num_Bytes := StrLen(Read_Data) / 2 ;RS232_Read() returns 2 characters for each byte

    Loop %Read_Data_Num_Bytes%
    {
        StringLeft, Byte, Read_Data, 2
        StringTrimLeft, Read_Data, Read_Data, 2
        Byte = 0x%Byte%
        Byte := Byte + 0 ;Convert to Decimal       
        ASCII_Chr := Chr(Byte)

        ;Send accumulated characters if suffix character detected
        IF (ASCII_Chr == scanSuffix) {
            string:=scanData
            
            If (Wrapper)
            {
                Send_WM_COPYDATA(scanData, A_Args[1])
            }
            
            If (string == "MultiScan" && !Wrapper)
            {
                JSONSend(ScanLine)
                ;ScanLine := ""
            }
                
            assess:=RegExReplace(string,"[^a-zA-z]",0)
            
            If (assess==0 && !Wrapper) {
                ;MsgBox String is numeric: %scanData% - %assess%
                SplashTextOff
                FormatTime, NowDateTime, A_Now, MM/dd/yyyy hh:mm:ss tt
                ScanLine := ScanLine "&" NowDateTime "=" scanData
                BufferLength++
                ;SB_SetText("Scan Queue: " BufferLength)
                If (BufferLength==1)
                {
                    QueueButton := "Scan Queued"
                    GuiControl,OSD:,Var1, Scan Queued
                }
                Else
                {
                QueueButton := "Scans Queued"
                GuiControl,OSD:,Var1, Scans Queued
                }
                GuiControl,1:,QueueButton,%QueueButton%
                LogFile.Write(NowDateTime "," scanData "`n")
                }
            ;Clear out barcode data and reset scanState to 0 (ignore)	
            scanData =
            }
            Else
            {
                scanData = %scanData%%ASCII_Chr%
            }

        ;Add new character to existing character string
        

    }

    }
;Critical, Off	
Return
}

;########################################################################
;###### CONNECTION TIMER ################################################
;########################################################################
{
connectionTimer:
If DebugLevel>2
    MsgBox connectionTimer Expired

IF (BTenable==True) {

    IF CheckBTConnection(BTdeviceName) {
        Return
    } 
    SetTimer, connectionTimer, Off
    RS232_Close(RS232_FileHandle)
    RS232_Settings = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=Off
    RS232_FileHandle:=RS232_Initialize(RS232_Settings)
    
    While (RS232_FileHandle == False)
    {
        Loop, 5
        {
            loopTimer:=6-A_Index
            SB_SetText("Trying again in " loopTimer " seconds")
            Sleep 1000
        }
        SB_SetText("Trying to open serial port")
        RS232_FileHandle:=RS232_Initialize(RS232_Settings)
    
        If DebugLevel>1
            MsgBox FileHandle %RS232_FileHandle%
    }
    SetTimer, connectionTimer, %ScanTimeout%
    SetTimer, SerialTimer, 100
    GuiControl,1:,ScanStatus,Scanner Connected
    SB_SetText("v" version)
    Return
}

SetTimer, connectionTimer, Off  ; i.e. the timer turns itself off here.
RS232_Settings = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=Off
RS232_FileHandle:=RS232_Initialize(RS232_Settings)
IF (!RS232_FileHandle) {
    MsgBox Comm Port %RS232Port% Not Open. Exiting.
    ExitApp
    Exit
}
GuiControl,1:,ScanStatus,Scanner Connected
SB_SetText("v" version)
;SetTimer, connectionTimer, %ScanTimeout%
SetTimer, SerialTimer, 100
return
}

;########################################################################
;###### TRIP TIMER $$####################################################
;########################################################################
{
tripTimer:
SetTimer, tripTimer, Off  ; i.e. the timer turns itself off here.
tripType:= %DefaultTrip%
;SplashTextOff
GuiControl,,Var,
return
}


;########################################################################
;###### SEND WINDOWS MESSAGE ############################################
;########################################################################
Send_WM_COPYDATA(ByRef StringToSend, ByRef TargetScriptTitle)  ; ByRef saves a little memory in this case.
; This function sends the specified string to the specified window and returns the reply.
; The reply is 1 if the target window processed the message, or 0 if it ignored it.
{
    VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)  ; Set up the structure's memory area.
    ; First set the structure's cbData member to the size of the string, including its zero terminator:
    SizeInBytes := (StrLen(StringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)  ; OS requires that this be done.
    NumPut(&StringToSend, CopyDataStruct, 2*A_PtrSize)  ; Set lpData to point to the string itself.
    Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
    TimeOutTime := 4000  ; Optional. Milliseconds to wait for response from receiver.ahk. Default is 5000
    ; Must use SendMessage not PostMessage.
    SendMessage, 0x004A, 0, &CopyDataStruct,, %TargetScriptTitle%,,,, %TimeOutTime% ; 0x004A is WM_COPYDATA.
    DetectHiddenWindows %Prev_DetectHiddenWindows%  ; Restore original setting for the caller.
    SetTitleMatchMode %Prev_TitleMatchMode%         ; Same.
    return ErrorLevel  ; Return SendMessage's reply back to our caller.
}

Send_WM_DDE_ADVISE()
{
WM_DDE_ADVISE := 0x03E2

}

;########################################################################
;###### RECEIVE WINDOWS MESSAGE #########################################
;########################################################################
OnMessage(0x004A, "Receive_WM_COPYDATA")  ; 0x004A is WM_COPYDATA

Receive_WM_COPYDATA(wParam, lParam)
{
    StringAddress := NumGet(lParam + 2*A_PtrSize)  ; Retrieves the CopyDataStruct's lpData member.
    CopyOfData := StrGet(StringAddress)  ; Copy the string out of the structure.
    ; Show it with ToolTip vs. MsgBox so we can return in a timely fashion:
    MsgBox %A_ScriptName%`nReceived the following string:`n%CopyOfData%
    return true  ; Returning 1 (true) is the traditional way to acknowledge this message.
}

;########################################################################
;###### SEND JSON POST TO GOOGLE API ####################################
;########################################################################
JSONSend(ScannedArray) {
    try{ ; only way to properly protect from an error here
        LogFile.Close()
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", WebAppURL, false)
        whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        http1=tripType=MultiScan
        httpSend=%http1%%ScannedArray%
        ;whr.SetRequestHeader("Content-Type", "application/json")
        ;JSONarray := JSON.Load("{""firstName"": ""John"", ""lastName"": ""Doe""}")
        ;JSONarray := JSON.Dump(JSONarray)
        ;If DebugLevel>1
        ;MsgBox HTTP %httpSend%
        whr.Send(httpSend)
        SystemSound("ScanBeep")
        ScannedArray := ""   
        ScanLine := ""
        BufferLength:=0
        QueueButton := "No Scans Queued"
        GuiControl,OSD:,Var1,
        GuiControl,1:,QueueButton,%QueueButton% 
        LogFile := FileOpen("Attendance.log", "a")
        }
        catch e 
        {
            SystemSound("FailBeep")
            Loop, 5
            {
                loopTimer:=6-A_Index
                SB_SetText("Queue Failure. Trying again in " loopTimer " seconds")
                Sleep 1000
            }
            JSONSend(ScannedArray)
            ;Return False
        }
    SB_SetText("v" version)
    Return ; % whr.responseText
    }

ManualSend:
JSONSend(ScanLine)
Return

;########################################################################
;###### CHECK BLUETOOTH CONNECTION ######################################
;########################################################################
CheckBTConnection(BTdeviceName) {
    DllCall("LoadLibrary", "str", "Bthprops.cpl", "ptr")
    VarSetCapacity(BLUETOOTH_DEVICE_SEARCH_PARAMS, 24+A_PtrSize*2, 0)
    NumPut(24+A_PtrSize*2, BLUETOOTH_DEVICE_SEARCH_PARAMS, 0, "uint")
    NumPut(1, BLUETOOTH_DEVICE_SEARCH_PARAMS, 16, "uint")   ; fReturnConnected
    VarSetCapacity(BLUETOOTH_DEVICE_INFO, 560, 0)
    NumPut(560, BLUETOOTH_DEVICE_INFO, 0, "uint")
    Loop ;Looping through connected BT devices and check for deviceName
    {
        GuiControl,1:,ScanStatus,Checking Connection
        If (A_Index = 1)
        {
            foundedDevice := DllCall("Bthprops.cpl\BluetoothFindFirstDevice", "ptr", &BLUETOOTH_DEVICE_SEARCH_PARAMS, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr")
            if !foundedDevice
                GuiControl,1:,ScanStatus,No Devices
        }
        else
        {
        if !DllCall("Bthprops.cpl\BluetoothFindNextDevice", "ptr", foundedDevice, "ptr", &BLUETOOTH_DEVICE_INFO)
            GuiControl,1:,ScanStatus,Not Connected
            Break
        }

        If (StrGet(&BLUETOOTH_DEVICE_INFO+64) = BTdeviceName)
        {
            GuiControl,1:,ScanStatus,Device Connected
            DllCall("Bthprops.cpl\BluetoothFindDeviceClose", "ptr", foundedDevice)
            Return True ;Send true if deviceName is connected
        }
    }
    
DllCall("Bthprops.cpl\BluetoothFindDeviceClose", "ptr", foundedDevice)
Return False ;Send false if BTdeviceName is not connected
}

;########################################################################
;###### AUDIO FILE PLAY #################################################
;########################################################################
SystemSound(soundFile) {

    Switch soundFile
    {
    Case "FailBeep":
        try{
        SoundPlay FailBeep.wav
        } catch e {
            SoundPlay *16
            }
    Case "ScanBeep":
        try{
        SoundPlay ScanBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "DeviceBeep":
        try{
        SoundPlay DeviceBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "TripBeep":
        try{
        SoundPlay TripBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "CheckOutBeep":
        try{
        SoundPlay CheckOutBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "CheckInBeep":
        try{
        SoundPlay CheckInBeep.wav
        } catch e {
            SoundPlay *64
            }            
    default:
        Return False
    }
Return True
}

;########################################################################
;###### CONFIG FILE CREATION ############################################
;########################################################################
CfgFileCreate() {
	MsgBox Can't open "LHScfg.ini". A new default config file has been created. Close the program and change default values before continuing.
    cfgFile.Close()
	
	IniWrite, Sheets WebApp URL (Scan Log), LHScfg.ini, GoogleAPI,        DeploymentURL
    IniWrite, Attendance,                   LHScfg.ini, HallPassAPI,      DefaultTrip
	IniWrite, 0,                            LHScfg.ini, GeneralSettings,  DebugLevel    
	IniWrite, 0,                            LHScfg.ini, GUIconfig,        DisplayMonitor
	IniWrite, COM3,                         LHScfg.ini, CommSetting,      RS232_Port
	IniWrite, 9600,                         LHScfg.ini, CommSetting,      RS232_Baud
	IniWrite, N,                            LHScfg.ini, CommSetting,      RS232_Parity
	IniWrite, 8,                            LHScfg.ini, CommSetting,      RS232_Data
	IniWrite, 1,                            LHScfg.ini, CommSetting,      RS232_Stop
	IniWrite, CR,                           LHScfg.ini, CommSetting,      scanSuffix
  	IniWrite, True,                         LHScfg.ini, CommSetting,      BTenable
    IniWrite, BarCode Scanner spp,          LHScfg.ini, CommSetting,      BTdeviceName
	ExitApp 
	Exit Script
}    

;########################################################################
;###### Initialize RS232 COM Subroutine #################################
;########################################################################
RS232_Initialize(RS232_Settings) {
  ;###### Extract/Format the RS232 COM Port Number ######
  ;7/23/08 Thanks krisky68 for finding/solving the bug in which RS232 COM Ports greater than 9 didn't work.
  StringSplit, RS232_Temp, RS232_Settings, `:
  RS232_Temp1_Len := StrLen(RS232_Temp1)  ;For COM Ports > 9 \\.\ needs to prepended to the COM Port name.
  If (RS232_Temp1_Len > 4)                   ;So the valid names are
    RS232_COM = \\.\%RS232_Temp1%             ; ... COM8  COM9   \\.\COM10  \\.\COM11  \\.\COM12 and so on...
  Else                                          ;
    RS232_COM = %RS232_Temp1%

  ;8/10/09 A BIG Thanks to trenton_xavier for figuring out how to make COM Ports greater than 9 work for USB-Serial Dongles.
  StringTrimLeft, RS232_Settings, RS232_Settings, RS232_Temp1_Len+1 ;Remove the COM number (+1 for the semicolon) for BuildCommDCB.
  ;MsgBox, RS232_COM=%RS232_COM% `nRS232_Settings=%RS232_Settings%

  ;###### Build RS232 COM DCB ######
  ;Creates the structure that contains the RS232 COM Port number, baud rate,...
VarSetCapacity(DCB, 28)
  BCD_Result := DllCall("BuildCommDCB"
       ,"str" , RS232_Settings ;lpDef
       ,"UInt", &DCB)        ;lpDCB
  If (BCD_Result <> 1)
  {
    MsgBox, There is a problem with Serial Port communication. `nFailed Dll BuildCommDCB, BCD_Result=%BCD_Result% `nThe Script Will Now Exit.
    Exit
  }

  ;###### Create RS232 COM File ######
  ;Creates the RS232 COM Port File Handle
  RS232_FileHandle := DllCall("CreateFile"
       ,"Str" , RS232_COM     ;File Name         
       ,"UInt", 0xC0000000   ;Desired Access
       ,"UInt", 3            ;Safe Mode
       ,"UInt", 0            ;Security Attributes
       ,"UInt", 3            ;Creation Disposition
       ,"UInt", 0            ;Flags And Attributes
       ,"UInt", 0            ;Template File
       ,"Cdecl Int")

  If (RS232_FileHandle < 1)
  {
    ;MsgBox, There is a problem with Serial Port communication. `nFailed Dll CreateFile, RS232_FileHandle=%RS232_FileHandle% `nThe Script Will Now Exit.
    ;Exit
    If DebugLevel>1
        MsgBox CreateFile
    Return False
  }

  ;###### Set COM State ######
  ;Sets the RS232 COM Port number, baud rate,...
  SCS_Result := DllCall("SetCommState"
       ,"UInt", RS232_FileHandle ;File Handle
       ,"UInt", &DCB)          ;Pointer to DCB structure
  If (SCS_Result <> 1)
  {
    ;MsgBox, There is a problem with Serial Port communication. `nFailed Dll SetCommState, SCS_Result=%SCS_Result% `nThe Script Will Now Exit.
    RS232_Close(RS232_FileHandle)
    ;Exit
    If DebugLevel>1
        MsgBox SetCommState
    Return False
  }

  ;###### Create the SetCommTimeouts Structure ######
  ReadIntervalTimeout        = 0xffffffff
  ReadTotalTimeoutMultiplier = 0x00000000
  ReadTotalTimeoutConstant   = 0x00000000
  WriteTotalTimeoutMultiplier= 0x00000000
  WriteTotalTimeoutConstant  = 0x00000000

  VarSetCapacity(Data, 20, 0) ; 5 * sizeof(DWORD)
  NumPut(ReadIntervalTimeout,         Data,  0, "UInt")
  NumPut(ReadTotalTimeoutMultiplier,  Data,  4, "UInt")
  NumPut(ReadTotalTimeoutConstant,    Data,  8, "UInt")
  NumPut(WriteTotalTimeoutMultiplier, Data, 12, "UInt")
  NumPut(WriteTotalTimeoutConstant,   Data, 16, "UInt")

  ;###### Set the RS232 COM Timeouts ######
  SCT_result := DllCall("SetCommTimeouts"
     ,"UInt", RS232_FileHandle ;File Handle
     ,"UInt", &Data)         ;Pointer to the data structure
  If (SCT_result <> 1)
  {
    ;MsgBox, There is a problem with Serial Port communication. `nFailed Dll SetCommState, SCT_result=%SCT_result% `nThe Script Will Now Exit.
    RS232_Close(RS232_FileHandle)
    ;Exit
    If DebugLevel>1
        MsgBox SetCommState
    Return False
  }
  
  Return %RS232_FileHandle%
}

;########################################################################
;###### Close RS23 COM Subroutine #######################################
;########################################################################
RS232_Close(RS232_FileHandle) {
  ;###### Close the COM File ######
  CH_result := DllCall("CloseHandle", "UInt", RS232_FileHandle)
  If (CH_result <> 1)
    ;MsgBox, Failed Dll CloseHandle CH_result=%CH_result%

  Return
}

;########################################################################
;###### Read from RS232 COM Subroutines #################################
;########################################################################
RS232_Read(RS232_FileHandle,Num_Bytes,ByRef RS232_Bytes_Received) {
  SetFormat, Integer, HEX

  ;Set the Data buffer size, prefill with 0x55 = ASCII character "U"
  ;VarSetCapacity won't assign anything less than 3 bytes. Meaning: If you
  ;  tell it you want 1 or 2 byte size variable it will give you 3.
  Data_Length  := VarSetCapacity(Data, Num_Bytes, 0x55)
  ;MsgBox, Data_Length=%Data_Length%

  ;###### Read the data from the RS232 COM Port ######
  ;MsgBox, RS232_FileHandle=%RS232_FileHandle% `nNum_Bytes=%Num_Bytes%
  Read_Result := DllCall("ReadFile"
       ,"UInt" , RS232_FileHandle   ; hFile
       ,"Str"  , Data             ; lpBuffer
       ,"Int"  , Num_Bytes        ; nNumberOfBytesToRead
       ,"UInt*", RS232_Bytes_Received   ; lpNumberOfBytesReceived
       ,"Int"  , 0)               ; lpOverlapped

  ;MsgBox, RS232_FileHandle=%RS232_FileHandle% `nRead_Result=%Read_Result% `nBR=%RS232_Bytes_Received% ,`nData=%Data%
  If (Read_Result <> 1)
  {
    ;MsgBox, There is a problem with Serial Port communication. `nFailed Dll ReadFile on RS232 COM, result=%Read_Result% - The Script Will Now Exit.
    RS232_Close(RS232_FileHandle)
    Return False
  }

  ;###### Format the received data ######
  ;This loop is necessary because AHK doesn't handle NULL (0x00) characters very nicely.
  ;Quote from AHK documentation under DllCall:
  ;     "Any binary zero stored in a variable by a function will hide all data to the right
  ;     of the zero; that is, such data cannot be accessed or changed by most commands and
  ;     functions. However, such data can be manipulated by the address and dereference operators
  ;     (& and *), as well as DllCall itself."
  i = 0
  Data_HEX =
  Loop %RS232_Bytes_Received%
  {
    ;First byte into the Rx FIFO ends up at position 0

    Data_HEX_Temp := NumGet(Data, i, "UChar") ;Convert to HEX byte-by-byte
    StringTrimLeft, Data_HEX_Temp, Data_HEX_Temp, 2 ;Remove the 0x (added by the above line) from the front

    ;If there is only 1 character then add the leading "0'
    Length := StrLen(Data_HEX_Temp)
    If (Length =1)
      Data_HEX_Temp = 0%Data_HEX_Temp%

    i++

    ;Put it all together
    Data_HEX := Data_HEX . Data_HEX_Temp
  }
  ;MsgBox, Read_Result=%Read_Result% `nRS232_Bytes_Received=%RS232_Bytes_Received% ,`nData_HEX=%Data_HEX%

  SetFormat, Integer, DEC
  Data := Data_HEX

  Return Data
}