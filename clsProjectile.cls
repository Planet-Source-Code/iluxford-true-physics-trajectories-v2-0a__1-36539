VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Declare the variables needed
Private mvarVelocity        As Double
Private mvarAngle           As Double
Private mvarX               As Double
Private mvarY               As Double
Private mvarTime            As Double
Private mvarOldX            As Double
Private mvarOldY            As Double
'local variable(s) to hold property value(s)
Private mvarID As Integer 'local copy
Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property




Public Function Apex() As Double
    ' Calculate the apex (maximum height over entire flight) of the projectile.
    ' The formula for this calculation is:
    Apex = Me.Velocity * ((Me.Velocity * Sin(Me.Angle) / GRAVITY)) * Sin(Me.Angle) - 0.5 * GRAVITY * ((Me.Velocity * Sin(Me.Angle)) / GRAVITY) ^ 2
End Function

Public Sub Initialise()
    ' Reset all the variables.
    With Me
        .Angle = 0
        .OldX = 0
        .OldY = 0
        .Time = 0
        .Velocity = 0
        .x = 0
        .y = 0
    End With
End Sub

Public Sub Update()

    'Calculate the X and Y co-ordinates of the projectile.
    Me.y = Me.Velocity * Me.Time * Sin(Me.Angle) - (0.5 * GRAVITY * Me.Time ^ 2)
    Me.x = (Me.Velocity * Cos(Me.Angle)) * Me.Time
    
    ' Two methods can be used to draw the projectile:
    ' 1) A constant line tracing the trajectory. This method is slower than (2)
        'frmMissiles.picField.Line (Me.OldX, Me.OldY)-(Me.X, Me.Y)
    ' 2) A dotted line representing the trajectory. This is the fastest method.
        frmMissiles.picField.PSet (Me.x, Me.y), vbYellow
        
    ' The 'Old' co-ordinates are used for Method(1) to connect the line to
    ' its previous point.
    Me.OldX = Me.x
    Me.OldY = Me.y
    
    ' Increment the time. A smaller increment will result in a slower but
    ' more accurate flight representation.
    Me.Time = Me.Time + 0.1
End Sub

Public Property Let OldY(ByVal vData As Double)
    mvarOldY = vData
End Property

Public Property Get OldY() As Double
    OldY = mvarOldY
End Property

Public Property Let OldX(ByVal vData As Double)
    mvarOldX = vData
End Property

Public Property Get OldX() As Double
    OldX = mvarOldX
End Property

Public Property Let Time(ByVal vData As Double)
    mvarTime = vData
End Property

Public Property Get Time() As Double
    Time = mvarTime
End Property

Public Property Let y(ByVal vData As Double)
    mvarY = vData
End Property

Public Property Get y() As Double
    y = mvarY
End Property

Public Property Let x(ByVal vData As Double)
    mvarX = vData
End Property

Public Property Get x() As Double
    x = mvarX
End Property

Public Property Let Angle(ByVal vData As Double)
    mvarAngle = vData
End Property

Public Property Get Angle() As Double
    Angle = mvarAngle
End Property

Public Property Let Velocity(ByVal vData As Double)
    mvarVelocity = vData
End Property

Public Property Get Velocity() As Double
    Velocity = mvarVelocity
End Property
