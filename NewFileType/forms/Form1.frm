VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Welcome"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   420
      Left            =   2340
      TabIndex        =   4
      Top             =   2295
      Width           =   735
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   2340
      TabIndex        =   3
      Top             =   1770
      Width           =   735
   End
   Begin VB.CommandButton cmdremovekeys 
      Caption         =   "Remove The New File Type"
      Height          =   450
      Left            =   90
      TabIndex        =   2
      Top             =   2295
      Width           =   2145
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install New File Type"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   1770
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      Height          =   1515
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   4920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function fixpath(lzpath As String)
    If Right$(lzpath, 1) = "\" Then fixpath = lzpath Else fixpath = lzpath & "\"
    
End Function

Private Sub cmdabout_Click()
    MsgBox "New file type for your Vb programs" _
    & vbCrLf & vbCrLf & " By Ben Jones", vbInformation, Form1.Caption
    
End Sub

Private Sub cmdexit_Click()
    Unload Form1
    
End Sub

Private Sub cmdInstall_Click()
Dim ans
Dim IconPath As String, ProgPath As String


    ans = MsgBox("This will now install the needed keys for your new file type " _
    & vbNewLine & "Do you want to carry on", vbYesNo Or vbQuestion)
    If ans = vbNo Then: Unload Form1: Exit Sub
    
    
    IconPath = fixpath(App.Path) & "appIcon.ico" ' location of icon
    ProgPath = "C:\dmpad.exe"   ' location of the program to load
    
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".dmtxt"  ' your new file type
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".dmtxt\DefaultIcon"  ' your new file types icon root
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".dmtxt\shell"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".dmtxt\shell\open"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".dmtxt\shell\open\command"
    
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, ".dmtxt\DefaultIcon", "", IconPath ' your new filetype icon to use
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, ".dmtxt\shell\open\command", "", Chr(34) & ProgPath & Chr(34) & " %1"
    
    MsgBox "The new keys have now been added in the registery for your new file type", vbInformation
    
    
End Sub

Private Sub cmdremovekeys_Click()
Dim ans

    ans = MsgBox("This will now remove all the keys for your new file type you just created " _
    & vbNewLine & "Do you want to carry on", vbYesNo Or vbQuestion)
    If ans = vbNo Then: Unload Form1: Exit Sub
    
    Reg32Mod.DeleteValue HKEY_CLASSES_ROOT, ".dmtxt\DefaultIcon", IconPath   ' your new filetype icon to use
    Reg32Mod.DeleteValue HKEY_CLASSES_ROOT, ".dmtxt\shell\open\command", Chr(34) & ProgPath & Chr(34) & "%"
    Reg32Mod.DeleteKey HKEY_CLASSES_ROOT, ".dmtxt\DefaultIcon"  ' your new file types icon root
    Reg32Mod.DeleteKey HKEY_CLASSES_ROOT, ".dmtxt\shell\open\command"
    Reg32Mod.DeleteKey HKEY_CLASSES_ROOT, ".dmtxt\shell\open"
    Reg32Mod.DeleteKey HKEY_CLASSES_ROOT, ".dmtxt\shell"
    Reg32Mod.DeleteKey HKEY_CLASSES_ROOT, ".dmtxt"   ' your new file type
    '
    MsgBox "All the keys have now been removed", vbInformation
    
    
    
End Sub
