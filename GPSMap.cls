VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GPSMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As String, ByVal lpFileName As String) As Long

Private rendPict As PictureBox
Private loadedMap As StdPicture

Private mapPosX As Single
Private mapPosY As Single
Private mapScale As Single

Private refPoint1Lat As Double
Private refPoint1Long As Double
Private refPixil1X As Double
Private refPixil1Y As Double

Private refPoint2Lat As Double
Private refPoint2Long As Double
Private refPixil2X As Double
Private refPixil2Y As Double

Private oldGPSPosX As Double
Private oldGPSPosY As Double

Private scaleChange As Boolean

Private mapWidth As Integer
Private mapHeight As Integer

Private Function loadMapIni(ByVal mapIni As String) As String
    
    Dim varINI As Long
    Dim tempIn As String
    tempIn$ = String(100, ";")

    '// Load the map image
    varINI = GetPrivateProfileString("mapsettings", "mapImage", "map.gif", tempIn$, 100, mapIni)
    loadMapIni = App.Path & Left$(tempIn, varINI)
    
    '// Load the point1Lat
    varINI = GetPrivateProfileString("mapsettings", "point1Lat", "map.gif", tempIn$, 100, mapIni)
    refPoint1Lat = Val(Left$(tempIn, varINI))
    
    '// Load the point1Long
    varINI = GetPrivateProfileString("mapsettings", "point1Long", "", tempIn$, 100, mapIni)
    refPoint1Long = Val(Left$(tempIn, varINI))
    
    '// Load the pixil 1 Y
    varINI = GetPrivateProfileString("mapsettings", "pixil1Lat", "", tempIn$, 100, mapIni)
    refPixil1Y = Val(Left$(tempIn, varINI))
    
    '// Load the pixil 1 X
    varINI = GetPrivateProfileString("mapsettings", "pixil1Long", "", tempIn$, 100, mapIni)
    refPixil1X = Val(Left$(tempIn, varINI))
    
    '// Load the point2Lat
    varINI = GetPrivateProfileString("mapsettings", "point2Lat", "", tempIn$, 100, mapIni)
    refPoint2Lat = Val(Left$(tempIn, varINI))
    
    '// Load the point2Long
    varINI = GetPrivateProfileString("mapsettings", "point2Long", "", tempIn$, 100, mapIni)
    refPoint2Long = Val(Left$(tempIn, varINI))
    
    '// Load the pixil 2 Y
    varINI = GetPrivateProfileString("mapsettings", "pixil2Lat", "", tempIn$, 100, mapIni)
    refPixil2Y = Val(Left$(tempIn, varINI))
    
    '// Load the pixil 2 X
    varINI = GetPrivateProfileString("mapsettings", "pixil2Long", "", tempIn$, 100, mapIni)
    refPixil2X = Val(Left$(tempIn, varINI))


    '// Load the width
    varINI = GetPrivateProfileString("mapsettings", "mapWidth", "", tempIn$, 100, mapIni)
    mapWidth = Val(Left$(tempIn, varINI))
    
    '// Load the height
    varINI = GetPrivateProfileString("mapsettings", "mapHeight", "", tempIn$, 100, mapIni)
    mapHeight = Val(Left$(tempIn, varINI))

End Function

Public Function createMap(ByRef rendBox As PictureBox, ByVal mapFile As String)
    
    Set rendPict = rendBox
    
    '// Load default values
    mapPosX = 0
    mapPosY = 0
    mapScale = 1
    
    '// Load the map file in (may take several seconds to minutes depending on the comp)
    Set loadedMap = LoadPicture(loadMapIni(mapFile))
    Call updateMap
    
End Function

Private Function updateMap()
    
    '// Apply changes according to our zoom, map size etc
    rendPict.PaintPicture loadedMap, mapPosX, mapPosY, mapWidth * mapScale, mapHeight * mapScale
    
End Function

Public Function setGPSPos(ByVal posLongX As Double, ByVal posLatY As Double)
    
    Dim posX As Double
    Dim posY As Double
    Dim pixilValueX As Double
    Dim pixilValueY As Double
    
    '// testing positions - debugging purposes
    'posLongX = 15038.095
    'posLatY = 3345.17105
    
    '// find the difference between one of our reference points and
    '// the point where we are
    posX = refPoint1Long - posLongX
    posY = refPoint1Lat - posLatY
    
    '// now we need to work out how much each pixil is worth to map point
    pixilValueX = (refPixil1X - refPixil2X) / (refPoint1Long - refPoint2Long)
    pixilValueY = (refPixil1Y - refPixil2Y) / (refPoint1Lat - refPoint2Lat)
    
    '// each pixel is a specific value of the long and lat scale
    posX = (refPixil1X - (posX * pixilValueX))
    posY = (refPixil1Y - (posY * pixilValueY))


    '// finally set location
    mapPosX = 0 - (posX - ((rendPict.ScaleWidth / 2) / mapScale)) * mapScale
    mapPosY = 0 - (posY - ((rendPict.ScaleHeight / 2) / mapScale)) * mapScale
    
        '// first we check to see if we have moved since last time.
    '// if we havnt, dont update the map, it just wastes resources
    If ((oldGPSPosX = mapPosX) And (oldGPSPosY = mapPosY)) And scaleChange = False Then Exit Function
    
    '// set old values
    oldGPSPosX = mapPosX
    oldGPSPosY = mapPosY
    
    Call updateMap
    
    scaleChange = False
    
End Function

Public Function setScale(ByVal scaleAmount As Single)
    
    '// set the scale depending on the scale passed
    mapScale = mapScale * scaleAmount
    scaleChange = True
    
End Function

