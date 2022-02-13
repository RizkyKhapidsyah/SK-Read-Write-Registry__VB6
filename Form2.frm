VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5010
   ClientLeft      =   4905
   ClientTop       =   2640
   ClientWidth     =   6600
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6600
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "This is Just a little front End to show you how to Use the Registry"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload.me
End Sub
