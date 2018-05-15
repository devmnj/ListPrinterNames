VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Command1_Click()
GetPrinterList List1
End Sub
Public Sub GetPrinterList(lstPrinter As ListBox)
Dim PrintData As Printer
Dim defprinterpos%

For Each PrintData In Printers

 'Add printer name and port to list

List1.AddItem PrintData.DeviceName & " at: " & PrintData.Port

'Check for default printer

 

Next

 

End Sub
 
