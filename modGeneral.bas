Attribute VB_Name = "modGeneral"
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const GRAVITY        As Long = 9.8
Public Const PI             As Long = 3.14159265358979
Public Const LoopRate       As Long = 10
Public MissilesGrounded     As Boolean
Public lngLastTick          As Long
Public lngNowTick           As Long
Public cMissile             As New clsProjectile
Public cMissileArray        As New colMissiles
Public iMissilesAirborne    As Integer

' This sub outputs text onto the DataSheet
Public Sub WriteText(x As Integer, y As Integer, str As String)
    TextOut frmMissiles.picData.hDC, x, y, str, Len(str)
End Sub

' Converts Degrees to Radians
Public Function Radians(Degrees As Double) As Double
    Radians = Degrees / 180 * 3.14159265358979
End Function

' Load a missile list.
Public Sub LoadMissileList(Path As String)
    Static fileNumber As Integer
    
    fileNumber = FreeFile
    
    ' Loop through each line, and break the data into velocity and angle.
    ' Then, add a new missile based on that information.
    Open Path For Input As fileNumber
        Do Until EOF(fileNumber)
            Dim dblVelocity As Double, dblAngle As Double, strBuffer As String
            Input #fileNumber, strBuffer
            
            dblVelocity = CDbl(Split(strBuffer, "#")(0))
            dblAngle = CDbl(Split(strBuffer, "#")(1))
            
            cMissileArray.Add dblVelocity, Radians(dblAngle), 0, 0, 0, 0, 0
    Loop
    
    Close fileNumber
End Sub

