VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Radians"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "- Jonas Ask"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sx = Me.ScaleWidth / 2
    Sy = Me.ScaleHeight / 2
    
    FindAng X, Y, Sx, Sy, Ang, Check1.Value
    
    Me.Cls
    Me.Print Ang
    Me.Line (Sx, Sy)-Step(2, 2), , BF
    Me.Line (Sx, Sy)-(X, Y)
End Sub
Public Function FindAng(X, Y, Sx, Sy, Ang, RAD) 'convert as set of coordinates to the angle between them (given standard VB scales)
Const Pi = 3.14159265358979
    If Y - Sy = 0 Then
        Ang = IIf(X >= Sx, 0, Pi)
    Else
        Ang = Atn((X - Sx) / (Y - Sy))
        Ang = Ang + (Pi / 2)
        If Y >= Sy Then
            Ang = Pi + Ang
        End If
    End If
    If RAD = 0 Then Ang = Ang / Pi * 180
End Function
