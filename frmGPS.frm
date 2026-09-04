VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmGPS 
   Caption         =   "GPS Map"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "-"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdZoomIn 
      Caption         =   "+"
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   5880
      Width           =   375
   End
   Begin VB.PictureBox picRendMap 
      AutoRedraw      =   -1  'True
      Height          =   5775
      Left            =   0
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   13
      Top             =   0
      Width           =   7215
      Begin VB.Line lneDirection 
         BorderColor     =   &H000000FF&
         X1              =   236
         X2              =   236
         Y1              =   191
         Y2              =   180
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   135
         Left            =   3480
         Shape           =   3  'Circle
         Top             =   2805
         Width           =   135
      End
   End
   Begin VB.Frame freProccess 
      Caption         =   "Processed"
      Height          =   5775
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox txtVel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox txtLong 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtLat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox txtTime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblVel 
         Caption         =   "Velocity:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label lblDir 
         Caption         =   "Direction:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblLat 
         Caption         =   "Latitude:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblLong 
         Caption         =   "Longitute:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer tmeClear 
      Interval        =   1000
      Left            =   4320
      Top             =   120
   End
   Begin MSCommLib.MSComm comPort 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      BaudRate        =   4800
   End
End
Attribute VB_Name = "frmGPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// declare the gps processing
Private gpsSet As GPSProcess
Private gpsRend As GPSMap

Private buildGPSString As String

Private Sub cmdZoomIn_Click()

    Call gpsRend.setScale(2)

End Sub

Private Sub cmdZoomOut_Click()

    Call gpsRend.setScale(0.5)

End Sub

Private Sub Form_Load()

    '// set the GPS classes
    Set gpsSet = New GPSProcess
    Set gpsRend = New GPSMap
    
    '// setup the map with its device and directory
    Call gpsRend.createMap(picRendMap, App.Path & "\map\1mapWI.ini")
    Call gpsSet.loadGPS(comPort)

End Sub

Private Sub tmeClear_Timer()

    '// make sure data exists in the com port buffer
    If comPort.CommEvent Then
        
        '// send data to the process gps function
        Call gpsSet.processGPS(comPort.Input)
        
    End If
    
    '// get latest stats from the gps module
    txtTime.Text = gpsSet.getGPSTime()
    txtDate.Text = gpsSet.getGPSDate()
    txtLat.Text = gpsSet.getLatPos()
    txtLong.Text = gpsSet.getLongPos()
    txtVel.Text = gpsSet.getVelocity() & " Km\h"
    txtDir.Text = gpsSet.getDirection()
    
    '// Set the arrow to point in the right direction
    '// Currently not complete or optimised, or tested properly for that matter
    'Dim tempX As Double
    'Dim tempDirection As Double
    'Dim tempNegX As Integer
    'Dim tempNegY As Integer
    
    'tempDirection = gpsSet.getDirection
    'tempDirection = 350
    
    'If tempDirection > 270 Then
    '    tempNegX = 1
    '    tempNegY = -1
    'ElseIf tempDirection > 180 Then
    '    tempNegY = -1
    '    tempNegX = -1
    'ElseIf tempDirection > 90 Then
    '    tempNegX = -1
    '    tempNegY = 1
    'ElseIf tempDirection > -1 Then
    '    tempNegX = 1
    '    tempNegY = -1
    'End If
    
    'tempX = 121 / (1 + (Tan(tempDirection)) ^ 2)
    'lneDirection.X2 = lneDirection.X1 + tempNegX * (tempX ^ 0.5)
    'lneDirection.Y2 = lneDirection.Y1 + tempNegY * (121 - tempX) ^ 0.5
    
    '// update the position on the gps
    Call gpsRend.setGPSPos(gpsSet.getLongPos(), gpsSet.getLatPos())
    
End Sub
