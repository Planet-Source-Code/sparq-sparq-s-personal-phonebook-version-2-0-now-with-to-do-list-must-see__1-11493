VERSION 5.00
Begin VB.Form frmContact 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6450
   Begin VB.CommandButton Command2 
      Caption         =   "Print Record"
      Height          =   375
      Left            =   180
      TabIndex        =   34
      Top             =   5340
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Record"
      Height          =   375
      Left            =   4680
      TabIndex        =   33
      Top             =   5340
      Width           =   1635
   End
   Begin VB.TextBox txtNotes 
      Height          =   1515
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   3780
      Width           =   6135
   End
   Begin VB.TextBox txtFName 
      Height          =   315
      Left            =   180
      TabIndex        =   16
      Top             =   300
      Width           =   1635
   End
   Begin VB.TextBox txtLName 
      Height          =   315
      Left            =   1920
      TabIndex        =   15
      Top             =   300
      Width           =   1635
   End
   Begin VB.ComboBox cmbCat 
      Height          =   315
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   300
      Width           =   2535
   End
   Begin VB.TextBox txtAdd1 
      Height          =   315
      Left            =   180
      TabIndex        =   13
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtAdd2 
      Height          =   315
      Left            =   180
      TabIndex        =   12
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtCity 
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtState 
      Height          =   315
      Left            =   1860
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtZip 
      Height          =   315
      Left            =   2340
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPhone1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3780
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPhone2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5100
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtFax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3780
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtCell 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5100
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   3120
      Width           =   6135
   End
   Begin VB.ComboBox cmbBDayM 
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2100
      Width           =   1815
   End
   Begin VB.ComboBox cmbBDayD 
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2460
      Width           =   675
   End
   Begin VB.ComboBox cmbBDayY 
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label lblDays 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4500
      TabIndex        =   32
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   195
      Left            =   180
      TabIndex        =   31
      Top             =   3540
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   1920
      TabIndex        =   28
      Top             =   60
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      Height          =   195
      Left            =   3780
      TabIndex        =   27
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   195
      Left            =   180
      TabIndex        =   25
      Top             =   1680
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   195
      Left            =   1800
      TabIndex        =   24
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zip"
      Height          =   195
      Left            =   2340
      TabIndex        =   23
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone(s):      (Double - Click to dial)"
      Height          =   195
      Left            =   3780
      TabIndex        =   22
      Top             =   720
      Width           =   2490
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      Height          =   195
      Left            =   3780
      TabIndex        =   21
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cell / Pager:"
      Height          =   195
      Left            =   5100
      TabIndex        =   20
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address:           (Double - Click to E-Mail)"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   2280
      Width           =   3330
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Address:        (Double - Click to visit URL)"
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   2880
      Width           =   3315
   End
   Begin VB.Line Line1 
      X1              =   3780
      X2              =   6300
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3780
      X2              =   6300
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday:"
      Height          =   195
      Left            =   3780
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   3660
      X2              =   3660
      Y1              =   60
      Y2              =   3060
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3675
      X2              =   3660
      Y1              =   60
      Y2              =   3060
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Public Changes As Boolean

Private Sub cmbBDayD_Click()
    Changes = True
    UpdateDays
End Sub

Private Sub cmbBDayM_Click()
    Changes = True
    UpdateDays
End Sub

Private Sub cmbBDayY_Click()
    Changes = True
End Sub

Private Sub cmbCat_Click()
    Changes = True
End Sub

Private Sub Command1_Click()
    If Changes = True Then UpdateMe: MsgBox "Updated!", vbExclamation, "Updated"
End Sub

Private Sub Command2_Click()
    'PrintRecord (txtLName & ", " & txtFName)
    MsgBox "PrintRecord isnt working yet - I am having problems testing it" & vbCrLf & _
           "because my printer is a P.O.S.  (if you dont know what POS means," & vbCrLf & _
           "It is not good :)"
End Sub

Private Sub Form_Load()
    FillDates
End Sub

Sub UpdateDays()
    If cmbBDayM.ListCount < 1 Then Exit Sub
    If cmbBDayD.ListCount < 1 Then Exit Sub
    If cmbBDayM.ListIndex < 0 Then Exit Sub
    If cmbBDayD.ListIndex < 0 Then Exit Sub
    
    Dim TDate As Date
    TDate = Left(cmbBDayM.Text, 2) & "/" & cmbBDayD.Text & "/00"
    lblDays = "Days until BDay: " & GetDays(TDate)
End Sub

Sub FillDates()
  Dim X As Integer
  Dim TempDate As Date
    For X = 1 To 12
        TempDate = Format(X, "00") & "/01/00"
        cmbBDayM.AddItem Format(X, "00") & "- " & Format(TempDate, "mmm")
    Next X
    For X = 1 To 31
        cmbBDayD.AddItem Format(X, "00")
    Next X
    
    For X = 1930 To Year(Date)
        cmbBDayY.AddItem X
    Next X
    cmbCat.AddItem "Friend"
    cmbCat.AddItem "Family"
    cmbCat.AddItem "Co-Worker"
    cmbCat.AddItem "General"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Changes Then
      Dim Answer As Integer
        Answer = MsgBox("Do you want to update this record?", vbYesNoCancel + vbQuestion, "Update")
        If Answer = vbYes Then
            UpdateMe
            Unload Me
        ElseIf Answer = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Sub UpdateMe()
    With frmMain.ContactTable
        .MoveFirst
        Do While Not .EOF
            If !LName & ", " & !Fname = Tag Then
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        On Error Resume Next
        .Edit
        If txtFName <> "" Then !Fname = txtFName
        If txtLName <> "" Then !LName = txtLName
        If txtPhone1 <> "" Then !Phone1 = txtPhone1
        If txtPhone2 <> "" Then !Phone2 = txtPhone2
        If txtCell <> "" Then !Cell = txtCell
        If txtFax <> "" Then !Fax = txtFax
        If txtAdd1 <> "" Then !Address1 = txtAdd1
        If txtAdd2 <> "" Then !Address2 = txtAdd2
        If txtCity <> "" Then !City = txtCity
        If txtState <> "" Then !State = UCase(txtState)
        If txtZip <> "" Then !zip = txtZip
        If txtNotes <> "" Then !Notes = txtNotes
        If txtEmail <> "" Then !EMail = txtEmail
        If txtURL <> "" Then !URL = txtURL
        If cmbCat.ListIndex > -1 Then !cat = cmbCat.ListIndex
        If cmbBDayM.ListIndex > -1 Then !BDayM = Left(cmbBDayM.Text, 2)
        If cmbBDayD.ListIndex > -1 Then !BDayD = cmbBDayD.ListIndex + 1
        If cmbBDayY.ListIndex > -1 Then !BDayY = cmbBDayY.Text
        .Update
    End With
    frmContList.LoadContacts
    Changes = False
End Sub

Private Sub txtAdd1_Change()
    Changes = True
End Sub

Private Sub txtAdd2_Change()
    Changes = True
End Sub

Private Sub txtCell_Change()
    Changes = True
End Sub

Private Sub txtCell_DblClick()
    Dial Me, FormatNumber(txtCell)
End Sub

Private Sub txtCity_Change()
    Changes = True
End Sub

Private Sub txtEmail_Change()
    Changes = True
End Sub

Private Sub txtEmail_DblClick()
  Dim AtSpot As Integer
  Dim DotSpot As Integer
    AtSpot = InStr(0, txtEmail, "@")
    AtSpot = InStr(0, txtEmail, ".")
    If AtSpot = 0 Or DotSpot = 0 Then Exit Sub
    ShellExecute hwnd, "open", "mailto:" & txtEmail, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub txtFax_Change()
    Changes = True
End Sub

Private Sub txtFax_DblClick()
    Dial Me, FormatNumber(txtFax)
End Sub

Private Sub txtFName_Change()
    Changes = True
End Sub

Private Sub txtLName_Change()
    Changes = True
End Sub

Private Sub txtNotes_Change()
    Changes = True
End Sub

Private Sub txtPhone1_Change()
    Changes = True
End Sub

Private Sub txtPhone1_DblClick()
    Dial Me, FormatNumber(txtPhone1)
End Sub

Private Sub txtPhone2_Change()
    Changes = True
End Sub

Private Sub txtPhone2_DblClick()
    Dial Me, FormatNumber(txtPhone2)
End Sub

Private Sub txtState_Change()
    Changes = True
End Sub

Private Sub txtURL_Change()
    Changes = True
End Sub

Private Sub txtURL_DblClick()
    If txtURL = "" Then Exit Sub
    ShellExecute hwnd, "open", txtURL, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub txtZip_Change()
    Changes = True
End Sub
