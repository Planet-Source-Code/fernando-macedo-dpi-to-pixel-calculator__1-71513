VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DPI to Pixel Calculator"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "DPI to Pixel Calculator.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1320
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   855
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   300
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Top             =   300
      Width           =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Pixels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1560
      TabIndex        =   7
      Top             =   630
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Size In:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Desired DPI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   4
      Top             =   90
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Desired Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Dim m_Combo_Ini(0 To 2) As Variant
Dim m_Combo_End(0 To 2) As Variant
Dim m_Value As Integer

Private Sub Combo1_Click()
m_Value = m_Combo_End(Combo1.ListIndex)
On Error GoTo Error
If Text1.Text = "" Then
Text3.Text = ""
End If
If Text1.Text = "" And Text2.Text = "" Then
Text3.Text = ""
Else
Select Case m_Value
Case 1
Text3.Text = Round(Text1.Text / 2.54 * Text2.Text, 0)
Case 2
Text3.Text = Round(Text1.Text / 2.54 * Text2.Text / 10, 0)
Case 3
Text3.Text = Round(Text1.Text * Text2.Text, 0)
End Select
End If
Error:
End Sub


Private Sub Form_Load()

Dim m_On_Top
m_On_Top = SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Dim m_Count
m_Combo_Ini(0) = "Centimeter"
m_Combo_Ini(1) = "Milimeter"
m_Combo_Ini(2) = "Inches"
m_Combo_End(0) = "1"
m_Combo_End(1) = "2"
m_Combo_End(2) = "3"
For m_Count = 0 To 2
Combo1.AddItem m_Combo_Ini(m_Count)
Next m_Count
Combo1.ListIndex = 1

End Sub


Private Sub Text1_Change()
On Error GoTo Error
If Text1.Text = "" Then
Text3.Text = ""
End If
If Text1.Text = "" And Text2.Text = "" Then
Text3.Text = ""
Else
Select Case m_Value
Case 1
Text3.Text = Round(Text1.Text / 2.54 * Text2.Text, 0)
Case 2
Text3.Text = Round(Text1.Text / 2.54 * Text2.Text / 10, 0)
Case 3
Text3.Text = Round(Text1.Text * Text2.Text, 0)
End Select
End If
Error:
End Sub


Private Sub Text2_Change()
On Error GoTo Error
If Text2.Text = "" Then
Text3.Text = ""
End If
If Text1.Text = "" And Text2.Text = "" Then
Text3.Text = ""
Else
Select Case m_Value
Case 1
Text3.Text = Round(Text1.Text / 2.54 * Text2.Text, 0)
Case 2
Text3.Text = Round(Text1.Text / 2.54 * Text2.Text / 10, 0)
Case 3
Text3.Text = Round(Text1.Text * Text2.Text, 0)
End Select
End If
Error:
End Sub


