VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   3930
   ClientTop       =   1950
   ClientWidth     =   8400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   8400
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About"
      Height          =   975
      Left            =   5880
      TabIndex        =   16
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ListBox MainKeyList 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   5520
      TabIndex        =   14
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton CmdDeleteValue 
      Caption         =   "DeleteValue"
      Height          =   735
      Left            =   2400
      TabIndex        =   13
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton CmdDeleteKey 
      Caption         =   "Delete Key"
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton CmdWriteKey 
      Caption         =   "Save Key"
      Height          =   735
      Left            =   2280
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton CmdGetKey 
      Caption         =   "Get Key"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Text            =   "James"
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Text            =   "Test"
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Text            =   "Software\TestApp"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Choose a key to use"
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.itechecom.com/begware"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   7680
      Width           =   7215
   End
   Begin VB.Label Label3 
      Caption         =   "Begware Software              Visit us Online"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   7200
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "Select From List"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Result"
      Height          =   375
      Index           =   4
      Left            =   4200
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Key"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Area"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Main Key Area"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'Begware Software ( a division of Independent Technical Services )
'http://www.itechecom.com
'itech@itechecom.com
'
'use this code for anything that you want
'but if you do please give credit to those that made this code
'Kevin Mackey for the routines to access the registry
'James Blanchette for the return codes and this example programSectionram
'If you can improve  upon this code please share it with others
' and send us a copy
' to
'itech@itechecom.com
'***************************************************************************
'
Dim TheHkey As Long ' you can not pass the key as a string it must be
' a long value so we use this var to pass to the module as well as
'the list box changes it from a string value to a long value ( see MainKeyList_click)


    Dim programSection As String ' Value as a string for the section of the registery that you want
    
    Dim skey As String ' value as string for the Key in the registry
    
    
    Dim kvalue As String ' Value as string for the Value of the Key ion the Registry
    
Private Sub CmdAbout_Click()
' just calls the about box
Load Form2
Form2.Show

End Sub

Private Sub CmdDeleteKey_Click()
If Text1(0).Text = "Choose a key to use" Then
Text1(4).Text = "Choose a Key to use first From the Main Key List"

Exit Sub
End If

'
    'var= DeleteKey(HKEY_CURRENT_USER, "Software\VBW")
    
    
    
    
    programSection = Text1(1).Text
    skey = Text1(2).Text
    kvalue = Text1(3).Text
    
    Text1(4).Text = DeleteKey(TheHkey, programSection)
    
End Sub

Private Sub CmdDeleteValue_Click()
'text1.text= DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    If Text1(0).Text = "Choose a key to use" Then
Text1(4).Text = "Choose a Key to use first From the Main Key List"

Exit Sub
End If
   
    
    programSection = Text1(1).Text
    skey = Text1(2).Text
    kvalue = Text1(3).Text
    Text1(4).Text = DeleteValue(TheHkey, programSection, skey)
End Sub

Private Sub CmdGetKey_Click()
If Text1(0).Text = "Choose a key to use" Then
Text1(4).Text = "Choose a Key to use first From the Main Key List"

Exit Sub
End If
    
    programSection = Text1(1).Text
    skey = Text1(2).Text
    kvalue = Text1(3).Text
    
    
    'text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
    Text1(4) = getstring(TheHkey, programSection, skey)
    
    
End Sub

Private Sub CmdWriteKey_Click()
'var= savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", "Value")
   If Text1(0).Text = "Choose a key to use" Then
Text1(4).Text = "Choose a Key to use first From the Main Key List"

Exit Sub
End If
    'Note all / should be \ when you use them
    ' you can add sections just by putting them there
    ' example
    ' you start with Software\TestApp
    'but you want other sections so
    'Software\TestApp\Startup
    'Software\TestApp\Registration
    ' etc etc
    'I'm not sure how far you can nest but it should be sufficent for most purposes
    ' Software\TestApp\Startup\OriginalSettings
    'Software\TestApp\Startup\UserSettings
    'Software\TestApp\Startup\UserSettings\window\LastRan\Colors
    '
    hkey = Text1(0).Text
    programSection = Text1(1).Text
    skey = Text1(2).Text
    kvalue = Text1(3).Text
    
    Text1(4).Text = savestring(TheHkey, programSection, skey, kvalue)
    
    
End Sub

Private Sub Form_Load()
'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_USERS = &H80000003
'Public Const HKEY_PERFORMANCE_DATA = &H80000004
'Public Const ERROR_SUCCESS = 0&
' add items to List as Strings
MainKeyList.AddItem ("HKEY_CLASSES_ROOT"), 0
MainKeyList.AddItem ("HKEY_CURRENT_USER"), 1
MainKeyList.AddItem ("HKEY_LOCAL_MACHINE"), 2
MainKeyList.AddItem ("HKEY_USERS"), 3
MainKeyList.AddItem ("HKEY_PERFORMANCE_DATA"), 4

End Sub



Private Sub MainKeyList_Click()
Dim hkeyTemp As String
' 0 Public Const HKEY_CLASSES_ROOT = &H80000000
' 1 Public Const HKEY_CURRENT_USER = &H80000001
' 2 Public Const HKEY_LOCAL_MACHINE = &H80000002
' 3 Public Const HKEY_USERS = &H80000003
' 4 Public Const HKEY_PERFORMANCE_DATA = &H80000004


'Set theHkey by getting the string from the list

Select Case MainKeyList.List(MainKeyList.ListIndex)
Case "HKEY_CLASSES_ROOT"
TheHkey = HKEY_CLASSES_ROOT
Case "HKEY_CURRENT_USER"
TheHkey = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
TheHkey = HKEY_LOCAL_MACHINE
Case "HKEY_USERS"
TheHkey = HKEY_USERS
Case "HKEY_PERFORMANCE_DATA"
TheHkey = HKEY_PERFORMANCE_DATA

End Select
' show what we got in the first text box

Text1(0).Text = MainKeyList.List(MainKeyList.ListIndex)


End Sub
