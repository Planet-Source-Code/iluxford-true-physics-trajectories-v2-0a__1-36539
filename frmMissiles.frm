VERSION 5.00
Begin VB.Form frmMissiles 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Missile Simulation"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   65.25
      ScaleMode       =   2  'Point
      ScaleWidth      =   401.25
      TabIndex        =   7
      Top             =   6360
      Width           =   8055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Field"
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtV 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "&Fire"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picField 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H008080FF&
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5145
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   840
      Width           =   8055
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Angle"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   405
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Velocity"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmMissiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' True-Physics Trajectories
' by Isaac Luxford
' Version 2.0a

' Included File Information:
'   frmMissiles.frm   - this form. Contains main-loop code and objects.
'   modGeneral.bas    - contains the declarations and subs.
'   clsProjectile.cls - individual projectile class. Contains calculations
'                       and drawing subs.
'   colMissiles.cls   - the missile array. Contains individual clsProjectile
'                       instances.

' Changes / Additions since v1.0
' 1) Multiple Missiles can now be in flight simultaneously. This is thanks
'    to a new Collection Class. cMissileArray represents a missile array, to
'    which individual missiles can be added. Then, in the main loop, the program
'    updates each missile in that array.

' 2) Added support for loading missile lists, using LoadMissileList(path)
'    the list needs to be in the same format as the included burst.txt, that is
'    "velocity#angle" e.g. for a projectile to be launched at 45 degrees with
'    a velocity of 100, you the file should read, on one line, "100#45".
'    Each missile should be included in the file on a separate line.

Private Sub cmdClear_Click()
    ' Clear the picture box where the trajectory is displayed.
    picField.Cls
End Sub

Private Sub cmdFire_Click()
    
    MissilesGrounded = False
    
    ' Add a new missile to the array IF both text-boxes are filled with
    ' numeric data.
    If IsNumeric(txtV.Text) And txtV.Text <> "" Then
        If IsNumeric(txtAngle.Text) And txtAngle.Text <> "" Then
            cMissileArray.Add CDbl(txtV.Text), Radians(CDbl(txtAngle.Text)), 0, 0, 0, 0, 0
        End If
    End If
  
    ' This will load the included missile list. Uncomment this line to see
    ' what I mean.
    ' LoadMissileList App.Path & "\burst.txt"
    
    ' The [While] condition for this loop is used so that the program will
    ' refresh the missiles' position ONLY when at least one is above airborne.
    ' If the condition is removed, the program will render the missiles'
    ' subterranean position :)
    Do While MissilesGrounded = False
        lngNowTick = GetTickCount
        
        ' The If-EndIf is used to control the speed of the program. The rate
        ' can be altered to suit specific purposes
        If (lngNowTick - lngLastTick) > LoopRate Then
            iMissilesAirborne = 0
            MissilesGrounded = True
            
            ' Loop through each missile and perform a few operations on them.
            For Each cMissile In cMissileArray
                ' If it's airborne then increment the counter.
                If cMissile.y > -1 Then
                    MissilesGrounded = False
                    iMissilesAirborne = iMissilesAirborne + 1
                End If
                
                ' Get the Missile object to calculate its new position
                ' and redraw the Missile.
                cMissile.Update
            Next
                    
            ' Output all the relevant data
            picData.Cls
            SetTextColor picData.hDC, vbBlack
            WriteText 2, 2, "Missile Simulation"
            WriteText 2, 25, "Missiles airborne: "
            WriteText 180, 25, str(iMissilesAirborne)
            picData.Refresh
                    
            lngLastTick = lngNowTick
            DoEvents
        End If
    Loop


End Sub

Private Sub Form_Load()
    ' Set up the scale of the picture box, in the format of:
    '   (X1,Y1)-(X2,Y2)
    ' where X1 and Y1 are the top left corner, while X2 and Y2 are the bottom
    ' right corner. Extend to the boundaries to see the full flight of
    ' projectiles with limits outside the current boundary.
    picField.Scale (0, 1000)-(1000, 0)
    picField.ForeColor = vbWhite

End Sub
