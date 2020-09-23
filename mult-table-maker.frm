VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiplication Table Maker"
   ClientHeight    =   2130
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3585
   Icon            =   "mult-table-maker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox rowstart 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox colstart 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox filen 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "c:\windows\desktop\mult.html"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox col 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "100"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox row 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "100"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Rows"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Columns"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu Systray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "Show"
      End
      Begin VB.Menu Hide 
         Caption         =   "Hide"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
ProgressBar1.Visible = True


Open filen.Text For Output As #1
'21 wide, 41 long (including spaces) (on portait)
Y1 = rowstart.Text
X1 = colstart.Text
X = col.Text
Y = row.Text

Print #1, "<table border=1>"
Print #1, "<tr><th>X";
For st = X1 To X
Print #1, "<th>"; st;
Next st
Print #1,

For rows = Y1 To Y
Print #1, "<tr><th>"; rows;
ProgressBar1.Value = ((rows - Y1) / (Y - Y1)) * 100
For cols = X1 To X
Print #1, "<td>"; rows * cols;
Next cols
Print #1,
Next rows

Print #1, "</table>"
Close #1

ProgressBar1.Visible = False

End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub filen_Change()
'if Mid(filen.Text, Len(filen.Text) - 4, 5)
End Sub

Private Sub filen_LostFocus()
If Mid(filen.Text, Len(filen.Text) - 4, 5) <> ".html" Then filen.Text = filen.Text + ".html"



End Sub

Private Sub Form_Load()
AddToTray Me, Me.Caption, Me.Icon
ProgressBar1.Visible = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_RBUTTONUP
            'Something useful I just found out:
            ' You need to verify the height, otherwise
            ' it'll pop up the menu mid-form, if the
            ' form is big enough
            temp = GetY
            If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                PopupMenu Systray
            End If
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
        RemoveFromTray
End Sub

Private Sub Hide_Click()
Me.Hide

End Sub

Private Sub Show_Click()
Me.Show
End Sub
