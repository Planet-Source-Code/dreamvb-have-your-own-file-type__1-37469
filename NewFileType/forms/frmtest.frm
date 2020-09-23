VERSION 5.00
Begin VB.Form frmtest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "the test program dmpad"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   3420
      Width           =   1470
   End
   Begin VB.TextBox txtpad 
      Height          =   2340
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   900
      Width           =   6390
   End
   Begin VB.Label Label1 
      Caption         =   "with any look your new file type has opened this program and you should see the text below"
      Height          =   645
      Left            =   255
      TabIndex        =   0
      Top             =   105
      Width           =   4065
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    End
    
End Sub

Private Sub Form_Load()
Dim cmd As String, StrBuff As String
Dim tFile As Long

    cmd = Command$
    
    If Len(Trim(cmd)) <= 0 Then
        Exit Sub
    Else
        tFile = FreeFile
        Open cmd For Binary Access Read As #tFile
            StrBuff = Space(LOF(tFile))
            Get #tFile, , StrBuff
        Close #tFile
        '
        txtpad.Text = StrBuff
        StrBuff = ""
        cmd = ""
    End If
    
End Sub
