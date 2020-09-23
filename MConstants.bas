Attribute VB_Name = "MConstants"
Option Explicit

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&

Public Sub Dial(Frm As Form, Num As String)
  Dim buff As String
  Dim nResult As Long
    nResult = tapiRequestMakeCall&(Trim$(Num), CStr(Frm.Caption), Frm.txtLName & ", " & Frm.txtFName, "")
    If nResult <> 0 Then
        buff = "Error dialing number : "
        Select Case nResult
               Case TAPIERR_NOREQUESTRECIPIENT
                    buff = buff & "No Windows Telephony dialing application is running and none could be started."
               Case TAPIERR_REQUESTQUEUEFULL
                    buff = buff & "The queue of pending Windows Telephony dialing requests is full."
               Case TAPIERR_INVALDESTADDRESS
                    buff = buff & "The phone number is Not valid."
               Case Else
                    buff = buff & "Unknown error."
               End Select
    End If
End Sub


Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Function FormatNumber(Text As String) As String
  Dim X As Integer
  Dim TempNum As String
  Dim CurLet As String
    For X = 1 To Len(Text)
        CurLet = Mid(Text, X, 1)
        If IsNumeric(CurLet) Then TempNum = TempNum & CurLet
    Next X
    FormatNumber = TempNum
End Function


Public Sub OpenContact(Name As String)
  Dim X As Integer
  Dim Another As New frmContact
  Dim YearDiff As Integer
  
    'Check if record is already open by searching the captions of all loaded forms.
    For X = 0 To Forms.Count - 1
        'If so, Exit sub
        If Forms(X).Caption = "Contacts - " & Name Then Forms(X).SetFocus: Exit Sub
    Next X
    
    
    With frmMain.ContactTable
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !LName & ", " & !Fname = Name Then
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        
        Dim BDate As Date
        On Error Resume Next
        Another.Visible = False
        Another.Width = 6570
        Another.Height = 6270
        Another.Caption = "Contacts - " & Name
        Another.txtFName = !Fname
        Another.txtLName = !LName
        Another.txtPhone1 = !Phone1
        Another.txtPhone2 = !Phone2
        Another.txtCell = !Cell
        Another.txtFax = !Fax
        Another.txtAdd1 = !Address1
        Another.txtAdd2 = !Address2
        Another.txtCity = !City
        Another.txtState = UCase(!State)
        Another.txtZip = !zip
        Another.txtNotes = !Notes
        Another.txtNotes.TabIndex = 0
        Another.txtEmail = !EMail
        Another.txtURL = !URL
        Another.cmbBDayM.ListIndex = Val(!BDayM) - 1
        Another.cmbBDayD.ListIndex = Val(!BDayD) - 1
        YearDiff = Year(Date) - !BDayY
        Another.cmbBDayY.ListIndex = Another.cmbBDayY.ListCount - (YearDiff + 1)
        Another.Tag = Name
        Another.cmbCat.ListIndex = !cat
        BDate = !BDayM & "/" & !BDayD & "/" & Year(Date)
        Another.lblDays = "Days until BDay: " & GetDays(BDate)
        Load Another
        Another.Visible = True
        Another.Changes = False
        Another.Show
    End With
    Exit Sub
End Sub

Public Function GetDays(BDate As Date) As Integer
    If DateDiff("d", Date, BDate) < 1 Then
        BDate = Month(BDate) & "/" & Day(BDate) & "/" & Year(BDate) + 1
    End If
    GetDays = DateDiff("d", Date, BDate)
End Function


Public Sub PrintRecord(Name As String)
  Dim Found As Boolean
    Found = False
    With frmMain.ContactTable
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !LName & ", " & !Fname = Name Then
                Found = True
                Exit Do
            Else
                .MoveNext
                Found = False
            End If
        Loop
        
        If Not (Found) Then MsgBox "Record not found", vbExclamation, "Error": Exit Sub
        
        Printer.ScaleMode = vbInches
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.FontSize = 18
        Printer.Print Name
        Printer.CurrentY = Printer.CurrentY + Printer.TextHeight(Name) + 0.3
        Printer.FontSize = 8
        Printer.CurrentX = 0
        Printer.Print !Phone1
        Printer.EndDoc
    End With
End Sub
