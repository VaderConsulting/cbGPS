Attribute VB_Name = "mdlGPSProcess"
Option Explicit

Private globTime As String
Private longPos As Double
Private latPos As Double
Private velocity As Double
Private direction As Double
Private globDate As String

Private tempPrevString As String

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
    frmGPS.Text1.Text = strdata
    
End Function


Private Function solveGPSPos(ByVal strdata As String)

    Dim splitStr() As String
    Dim i As Integer
    
    splitStr = Split(strdata, ",")
    
    '// ensure complete string exists
    If UBound(splitStr) <> 11 Then
        tempPrevString = strdata
        Exit Function
    End If
    
    '// Loop through each item and add it to the data sheets
    For i = LBound(splitStr) To UBound(splitStr)
        frmGPS.lstGPSIn.AddItem i & ": " & splitStr(i)
    Next i
    
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
    
    getVelocity = velocity
    
End Function

Public Function getDirection() As Double
    
    getDirection = direction
    
End Function

Public Function getGPSDate() As String

    getGPSDate = Left$(globDate, 2) & "/" & Mid$(globDate, 3, 2) & "/" & Mid$(globDate, 5, 2)

End Function
