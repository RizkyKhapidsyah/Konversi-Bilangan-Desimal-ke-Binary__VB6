VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mengkonversi Bilangan Desimal ke Binary"
   ClientHeight    =   1995
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function dec2bin(mynum As Variant) As String
Dim loopcounter As Integer
  If mynum >= 2 ^ 31 Then
     dec2bin = "Bilangan terlalu besar!"
     Exit Function
  End If
  Do
    If (mynum And 2 ^ loopcounter) = 2 ^ loopcounter Then
       dec2bin = "1" & dec2bin
    Else
       dec2bin = "0" & dec2bin
    End If
    loopcounter = loopcounter + 1
  Loop Until 2 ^ loopcounter > mynum
End Function

'Masukkan bilangan ke dalam Text1.
'Lihat hasilnya di Label1...
Private Sub Text1_Change()
  Label1.Caption = dec2bin(Text1.Text)
End Sub

