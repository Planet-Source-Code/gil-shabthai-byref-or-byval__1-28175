VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "ByVal & ByRef"
   ClientHeight    =   3876
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8436
   LinkTopic       =   "Form1"
   ScaleHeight     =   3876
   ScaleWidth      =   8436
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "By Val"
      Height          =   492
      Left            =   5520
      TabIndex        =   3
      Top             =   3240
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "By Ref"
      Height          =   492
      Left            =   1440
      TabIndex        =   2
      Top             =   3240
      Width           =   972
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   3012
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   5313
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3012
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   5313
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0099
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x As Integer
    x = 3
    DoSomthing x
    MsgBox "ByRef x Change To " & x, vbInformation, ""
End Sub

Sub DoSomthing(ByRef y As Integer)
    y = 5
End Sub

Private Sub Command2_Click()
    Dim x As Integer
    x = 3
    DoSomthingElse x
    MsgBox "ByVal x Without eny change is: " & x, vbInformation, ""
End Sub
Sub DoSomthingElse(ByVal y As Integer)
    y = 5
End Sub
Private Sub Form_Load()

RichTextBox1.Text = "Private Sub Command1_Click()" & vbNewLine & _
                       vbTab & "Dim x As Integer" & vbNewLine & _
                       vbTab & "x = 3" & vbNewLine & _
                       vbTab & "DoSomthing x" & vbNewLine & _
                       vbTab & "MsgBox x" & vbNewLine & _
                       "End Sub" & vbNewLine & _
                       vbNewLine & "Sub DoSomthing(ByRef y As Integer)" & vbNewLine & _
                       vbTab & "y = 5" & vbNewLine & _
                       "End Sub"

RichTextBox2.Text = "Private Sub Command1_Click()" & vbNewLine & _
                       vbTab & "Dim x As Integer" & vbNewLine & _
                       vbTab & "x = 3" & vbNewLine & _
                       vbTab & "DoSomthingElse x" & vbNewLine & _
                       vbTab & "MsgBox x" & vbNewLine & _
                       "End Sub" & vbNewLine & _
                       vbNewLine & "Sub DoSomthingElse(ByVal y As Integer)" & vbNewLine & _
                       vbTab & "y = 5" & vbNewLine & _
                       "End Sub"

End Sub





