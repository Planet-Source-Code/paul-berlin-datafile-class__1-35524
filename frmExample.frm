VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Example Form - By Paul Berlin 2002"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtIncome 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Text            =   "100000"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cmbSex 
      Height          =   315
      ItemData        =   "frmExample.frx":0000
      Left            =   120
      List            =   "frmExample.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      Height          =   315
      Left            =   120
      MaxLength       =   255
      TabIndex        =   3
      Text            =   "West Ave. 12"
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   120
      MaxLength       =   255
      TabIndex        =   1
      Text            =   "John Doe"
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Income:"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Sex:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
  Dim File As New clsDatafile
  
  'Setup filename to read from
  File.Filename = App.Path & "\temp.dat"
  File.OpenFile
  
  'Read Name as string
  txtName = File.ReadStr
  'Read Address as string
  txtAddress = File.ReadStr
  'Read Sex as an byte
  cmbSex.ListIndex = File.ReadByte
  'Read Income as long
  txtIncome = CStr(File.ReadLong)
  
  'Use this if you want to close the file before ending the sub
  Set File = Nothing
  
  'Remove the file
  Kill App.Path & "\temp.dat"
  
  'Enable & Disable controls
  cmdSave.Enabled = True
  cmdLoad.Enabled = False
  
End Sub

Private Sub cmdSave_Click()
  Dim File As New clsDatafile
  
  'Setup filename to write to
  File.Filename = App.Path & "\temp.dat"
  File.OpenFile
  
  'Write Name as string
  File.WriteStr Trim(txtName)
  'Write Address as string
  File.WriteStr Trim(txtAddress)
  'Write Sex as an byte
  File.WriteByte CByte(cmbSex.ListIndex)
  'Write Income as long
  File.WriteLong CLng(val(txtIncome))
  
  'Clear controls
  txtName = ""
  txtAddress = ""
  txtIncome = ""
  cmbSex.ListIndex = -1
  
  'Enable & Disable controls
  cmdSave.Enabled = False
  cmdLoad.Enabled = True
  
End Sub

Private Sub cmdInfo_Click()
  MsgBox "This is an example of how to use the Datafile class. The Name & Address fields are saved using WriteStr, Sex using WriteByte and Income using WriteLong." & vbNewLine & "Press Save to write the info to temp.dat and Load to read it again." & vbNewLine & vbNewLine & "Read the source for more info.", vbInformation, "Info"
End Sub

Private Sub Form_Load()
  'Setup cmbSex
  cmbSex.ListIndex = 0
End Sub
