VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colMissiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsProjectile"
Attribute VB_Ext_KEY = "Member0" ,"clsProjectile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mCol As Collection

Public Sub RefreshAll()
    For Each Missile In MissileArray
        Missile.GetPos: Missile.Draw
    Next Missile
End Sub

Public Function Add(Velocity As Double, Angle As Double, x As Double, y As Double, Time As Double, OldX As Double, OldY As Double, Optional sKey As String) As clsProjectile
    Dim objNewMember As clsProjectile
    Set objNewMember = New clsProjectile

    objNewMember.Velocity = Velocity
    objNewMember.Angle = Angle
    objNewMember.x = x
    objNewMember.y = y
    objNewMember.Time = Time
    objNewMember.OldX = OldX
    objNewMember.OldY = OldY
    
      If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsProjectile
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub LoadMissiles(Path As String)
fileMissile = FreeFile
    Open Path For Input As fileMissile
        Do Until EOF(1)
            Input #fileMissile, strBuffer
                
                dblElement(1) = Split(strBuffer, "#")(0)
                dblElement(2) = Split(strBuffer, "#")(1)
                
                Me.Add dblElement(1), Radians(dblElement(2)), 0, 0, 0, 0, 0, True
        Loop
    Close #fileMissile
End Sub

