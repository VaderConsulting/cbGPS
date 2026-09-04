VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GPSProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As String, ByVal lpFileName As String) As Long

Private globTime As String
Private longPos As Double
Private latPos As Double
Private velocity As Double
Private direction As Double
Private globDate As String

Private speedType As String
Private tempPrevString As String

Public comPort As Integer

Private Function loadGPSConfig()
    
    Dim varINI As Long
    Dim tempIn As String
    tempIn$ = String(100, ";")
    
    '// Load the gps settings
    varINI = GetPrivateProfileString("gpsconfig", "comPort", "5", tempIn$, 100, App.Path & "\gpsConfig.ini")
    comPort = Val(Left$(tempIn, varINI))
    
    '// Load the gps settings
    varINI = GetPrivateProfileString("gpsconfig", "velocity", "kph", tempIn$, 100, App.Path & "\gpsConfig.ini")
    speedType = Left$(tempIn, varINI)
    
End Function

Public Function loadGPS(ByRef gpsComPort As MSComm)

    '// this module is the one most likely to get an error
    '// this is because the user may already have the comm port in use
    '// or the user may have the wrong port number in the ini file
    On Error GoTo portError

    Call loadGPSConfig

    '// Set the input to ascii (text), and setup the port and buffer max length
    gpsComPort.InputMode = comInputModeText
    gpsComPort.InputLen = 1024 '// about 1k of ram assigned to comm buffer
    gpsComPort.CommPort = comPort
    
    '// Open the serial port
    gpsComPort.PortOpen = True
    
    Exit Function
    
portError:
    
    '// handle the error
    Dim i As Integer
    i = MsgBox("Error: Comm port settings in gpsConfig.ini may be incorrect or serial port may already be in use.", vbCritical, "Comm Error")

End Function

Public Function processGPS(ByVal strdata As String)

    Dim splitBuf() As String
    Dim i As Integer
    
    '// if characters need to overlap put them at the front of the string
    tempPrevString = tempPrevString & strdata
    strdata = tempPrevString
    
    '// clear the temporary buffer
    tempPrevString = ""
    
    '// divide the strings by the delimiter
    splitBuf = Split(strdata, "$")
    
    '// Loop through each line from the gps
    For i = LBound(splitBuf) To UBound(splitBuf)
        
        '// if its the long and lat information line
        If InStr(splitBuf(i), "GPRMC") Then
            
            '// Solve / process each line
            Call solveGPSPos(splitBuf(i))
            
        End If
    Next i
    
    '// Output the data to a text box [debugging]
    'frmGPS.Text1.Text = strdata
    
End Function


Private Function solveGPSPos(ByVal strdata As String)

    Dim splitStr() As String
    Dim i As Integer
    
    splitStr = Split(strdata, ",")
    
    '// ensure complete string exists, otherwise trial it across to the next count
    If UBound(splitStr) <> 11 Then
        tempPrevString = strdata
        Exit Function
    End If
    
    '// Loop through each item and add it to the data sheets [debugging]
    'For i = LBound(splitStr) To UBound(splitStr)
        'frmGPS.lstGPSIn.AddItem i & ": " & splitStr(i)
    'Next i
    
    '// get the stats from the split up string
    globTime = splitStr(1)
    latPos = Val(splitStr(3))
    longPos = Val(splitStr(5))
    velocity = Val(splitStr(7))
    direction = Val(splitStr(8))
    globDate = splitStr(9)

End Function

Public Function getGPSTime() As String
    
    '// build a time string to send back formatted
    getGPSTime = Left$(globTime, 2) & ":" & Mid$(globTime, 3, 2) & ":" & Mid$(globTime, 5, 2)

End Function

Public Function getLatPos() As Double
    
    getLatPos = latPos
    
End Function

Public Function getLongPos() As Double
    
    getLongPos = longPos
    
End Function

Public Function getVelocity() As Double
    
    '// Depending on everyones speed the return value is different
    If speedType = "kph" Then
        getVelocity = velocity * 2
    ElseIf speedType = "mph" Then
        getVelocity = (velocity * 1.25)
    Else
        getVelocity = velocity / 1.8
    End If
    
End Function

Public Function getDirection() As Double
    
    getDirection = direction
    
End Function

Public Function getGPSDate() As String

    getGPSDate = Left$(globDate, 2) & "/" & Mid$(globDate, 3, 2) & "/" & Mid$(globDate, 5, 2)

End Function

