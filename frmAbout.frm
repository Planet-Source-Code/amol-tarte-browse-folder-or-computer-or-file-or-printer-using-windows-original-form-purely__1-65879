VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About WinXpC Engine."
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   6
      Left            =   4080
      Top             =   1800
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "O&k"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   765
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About Author:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2250
      TabIndex        =   0
      Top             =   60
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   120
      Picture         =   "frmAbout.frx":08CA
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Single
Dim j As Byte
Dim done As Boolean

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
          lblCredits(0).Caption = "Amol Tarte."
          lblCredits(1).Caption = "Pune, Maharashtra"
          lblCredits(2).Caption = "India."
          lblCredits(3).Caption = "email: amoltarte@gmail.com"
          lblCredits(4).Caption = "Cell: (91) 9822220318"
For i = 0 To 4
    lblCredits(i).Left = Label1.Left
Next i
DoG Me
End Sub

Public Sub DoG(FormName As Object)
    On Error Resume Next
    Dim i As Integer, Y As Integer
    FormName.AutoRedraw = True
    FormName.DrawStyle = 6
    FormName.DrawMode = 13
    FormName.DrawWidth = 13
    FormName.ScaleMode = 3
    FormName.ScaleHeight = 256
    For i = 255 To 0 Step -1
        FormName.Line (0, Y)-(FormName.Width, Y + 1), RGB(i, i, 255), BF
        Y = Y + 1
    Next i
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()

If done = False Then
    If temp < 255 Then
        For j = 0 To 4
            lblCredits(j).ForeColor = RGB(temp, temp, temp)
        Next j
        temp = temp + 1
    Else
            done = True
    End If
Else
    If temp > 0 Then
        For j = 0 To 4
            lblCredits(j).ForeColor = RGB(temp, temp, temp)
        Next j
        temp = temp - 1
    Else
        done = False
    End If
End If
End Sub
