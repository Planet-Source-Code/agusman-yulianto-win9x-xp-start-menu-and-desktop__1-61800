Attribute VB_Name = "mdlGeneral"

' Global database connection
Global g_ConnectionType As String
Global g_objConn           As ADODB.Connection
Global g_strUserName       As String
Global g_strPassword       As String
Global g_strDBName       As String
Global strConnString As String
' Global constants
Global Const INICONN_FN    As String = "RASK.INI"
Global g_UnitDinas As String
Global g_NamaUnitDinas As String
Global g_NamaKepalaDinas As String
Global g_NIPKepalaDinas As String
Global g_Jabatan As String
Global g_Bidang As String
Global g_PemegangKas As String
Global g_JabatanPemegangKas As String
Global g_nipPemegangKas As String
    
Global m_FileGabung As String
Global NIPAnggaran(1 To 9) As String
Global NMAnggaran(1 To 9) As String
Global JBTAnggaran(1 To 9) As String

Global Kegiatan As String
Global Tahun As String

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Enum CommandType
    ctNew
    ctEdit
    ctDelete
    ctSave
    ctPrint
    ctPrevious
    ctNext
    ctBrowse
    ctFind
End Enum

Sub OpenConnection(ByVal FileName As String)
    strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & FileName & ";" & _
                    "Persist Security Info=True"
    Set g_objConn = New ADODB.Connection
    With g_objConn
       .ConnectionString = strConnString
       .ConnectionTimeout = 5
       .Open
    End With
    Screen.MousePointer = 0
    g_strDBName = FileName
End Sub
Function CloseConnection() As Boolean
   g_objConn.Close
   Set g_objConn = Nothing
End Function
Sub ShowMsg(ByVal strInfo As String, _
            Optional ByVal blnSuccess As Boolean, _
            Optional ByVal strTitle As String)
   
    Dim strWarn As String
   
    If blnSuccess Then
        strWarn = IIf(strTitle = "", "Information", strTitle)
        MsgBox strInfo, vbInformation, strWarn
    Else
        strWarn = IIf(strTitle = "", "Warning", strTitle)
        MsgBox strInfo, vbExclamation, strWarn
    End If
End Sub

Sub ShowError(Optional ByVal strMsg As String, _
              Optional ByVal strTitle As String)
   Dim sTmp As String
   Dim sTitle As String
   
   Screen.MousePointer = vbDefault
   
   If strMsg <> "" Then
      sTmp = "The following Error occurred:" & vbCrLf & vbCrLf
      sTmp = sTmp & Err.Number & ": " & Err.Description & vbCrLf
   Else
      sTmp = strMsg
   End If
          
   Err.Clear
   Beep
   If strTitle = "" Then
      sTitle = "Error Message"
   Else
      sTitle = strTitle
   End If
   MsgBox sTmp, vbOKOnly + vbExclamation, sTitle
End Sub
Function vVal(ByRef vntFieldVal As Variant) As Variant
    If IsNull(vntFieldVal) Then
        vVal = vbNullString
    Else
        vVal = CStr(vntFieldVal)
    End If
End Function
Function AddPath(ByVal strFileName As String) As String
    Dim strWindir As String
   strWindir = WinDir
   AddPath = IIf(Right(strWindir, 1) <> "\", strWindir & "\", strWindir) & strFileName
End Function

Function IsDateDDMMYYYY(ByVal DateString As String, ByVal DateSeparator As String) As Boolean
    Dim Kabisat As Boolean
    
    DateString = Trim(DateString)
    DateSeparator = Trim(DateSeparator)
    IsDateDDMMYYYY = False
    If IsDate(DateString) Then
        DateString = Replace(DateString, DateSeparator, "", , , vbTextCompare)
        If CInt(Mid(DateString, 3, 2)) = 2 Then
            If (Right(DateString, 4) Mod 4) = 0 Then
                If CInt(Left$(DateString, 2)) <= 29 Then IsDateDDMMYYYY = True
            Else
                If CInt(Left$(DateString, 2)) <= 28 Then IsDateDDMMYYYY = True
            End If
        Else
            Select Case CInt(Mid(DateString, 3, 2))
                Case 1, 3, 5, 7, 8, 10, 12
                    If CInt(Left$(DateString, 2)) <= 31 Then IsDateDDMMYYYY = True
                Case 4, 6, 9, 11
                    If CInt(Left$(DateString, 2)) <= 30 Then IsDateDDMMYYYY = True
            End Select
        End If
    End If
End Function

Function NumToText(dblValue As Double) As String
    Static ones(0 To 9) As String
    Static teens(0 To 9) As String
    Static tens(0 To 9) As String
    Static thousands(0 To 4) As String
    Dim i As Integer, nPosition As Integer
    Dim nDigit As Integer, bAllZeros As Integer
    Dim strResult As String, strTemp As String
    Dim tmpBuff As String

    ones(0) = "kosong"
    ones(1) = "se"
    ones(2) = "dua"
    ones(3) = "tiga"
    ones(4) = "empat"
    ones(5) = "lima"
    ones(6) = "enam"
    ones(7) = "tujuh"
    ones(8) = "delapan"
    ones(9) = "sembilan"

    teens(0) = "sepuluh"
    teens(1) = "sebelas"
    teens(2) = "dua belas"
    teens(3) = "tiga belas"
    teens(4) = "empat belas"
    teens(5) = "lima belas"
    teens(6) = "enam belas"
    teens(7) = "tujuh belas"
    teens(8) = "delapan belas"
    teens(9) = "sembilan belas"

    tens(0) = ""
    tens(1) = "sepuluh"
    tens(2) = "dua puluh"
    tens(3) = "tiga puluh"
    tens(4) = "empat puluh"
    tens(5) = "lima puluh"
    tens(6) = "enam puluh"
    tens(7) = "tujuh puluh"
    tens(8) = "delapan puluh"
    tens(9) = "sembilan puluh"

    thousands(0) = ""
    thousands(1) = "ribu"
    thousands(2) = "juta"
    thousands(3) = "milyar"
    thousands(4) = "trilyun"

    On Error GoTo NumToTextError
    
    strResult = Format((dblValue - Int(dblValue)) * 100, "00")
    If strResult <> "00" Then
        strResult = strResult & " sen"
    Else
        strResult = ""
    End If
    
    strTemp = CStr(Int(dblValue))
    For i = Len(strTemp) To 1 Step -1
        nDigit = Val(Mid$(strTemp, i, 1))
        nPosition = (Len(strTemp) - i) + 1
        Select Case (nPosition Mod 3)
            Case 1
                bAllZeros = False
                If i = 1 Then
                    If nPosition > 6 Then
                        If nDigit = 1 Then
                            tmpBuff = "satu "
                        Else
                            tmpBuff = ones(nDigit) & " "
                        End If
                    Else
                        tmpBuff = ones(nDigit) & IIf(nDigit = 1, "", " ")
                    End If
                ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
                    tmpBuff = teens(nDigit) & " "
                    i = i - 1
                ElseIf nDigit > 0 Then
                    If Len(strTemp) > 4 Then
                        If nDigit = 1 Then
                            tmpBuff = "satu " & IIf(nDigit = 1, "", " ")
                        Else
                            tmpBuff = ones(nDigit) & IIf(nDigit = 1, "", " ")
                        End If
                    Else
                        tmpBuff = ones(nDigit) & IIf(nDigit = 1, "", " ")
                    End If
                Else
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    tmpBuff = tmpBuff & thousands(nPosition / 3) & " "
                End If
                strResult = tmpBuff & strResult
            Case 2
                If nDigit > 0 Then
                    strResult = tens(nDigit) & " " & strResult
                End If
            Case 0
                If nDigit > 0 Then
                    strResult = ones(nDigit) & IIf(nDigit = 1, "", " ") & "ratus " & strResult
                End If
        End Select
    Next i
    If Len(strResult) > 0 Then
        strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
    End If
    
    NumToText = strResult
    GoTo EndNumToText

NumToTextError:
    MsgBox Err.Description

EndNumToText:

End Function

Public Function Convert(strField As String) As String
    Dim i As Integer
    
    For i = 215 To 218
        Convert = Replace(strField, "'", Chr(215))
    Next
End Function

Public Function DeConvert(strField As String) As String
    DeConvert = Replace(strField, Chr(215), "'")
End Function

Public Function NZ(varValue As Variant) As Variant
    If IsNull(varValue) Then
        NZ = 0
    Else
        NZ = varValue
    End If
End Function

Sub SelectText(ByVal X As TextBox)
    X.SelStart = 0
    X.SelLength = Len(X.Text)
End Sub

Sub RefreshData()
    If InStr(1, g_strDBName, ".MDB", vbTextCompare) <> 0 Then 'If Ms-Access File
        g_objConn.Close
        g_objConn.Open strConnString
    End If
End Sub
Function TypeName(iType As Integer) As String
    Select Case iType
    Case adBigInt
        TypeName = "Big Integer"
    Case adBinary
        TypeName = "Binary"
    Case adBoolean
        TypeName = "Boolean"
    Case adSmallInt
        TypeName = "Byte"
    Case adChar
        TypeName = "Char"
    Case adCurrency
        TypeName = "Currency"
    Case adDate
        TypeName = "Date"
    Case adDecimal, adDouble, adNumeric
        TypeName = "Double"
    Case adGUID
        TypeName = "GUID"
    Case adInteger
        TypeName = "Integer"
    Case adLongVarBinary
        TypeName = "Long Binary"
    Case adLongVarChar
        TypeName = "Memo"
    Case adSingle
        TypeName = "Single"
    Case adVarChar
        TypeName = "Text"
    Case adDBTime
        TypeName = "Time"
    Case adDBTimeStamp
        TypeName = "Time Stamp"
    Case Else
        TypeName = ""
End Select
End Function

Sub CreateTableFromQuery(ByVal TableName As String, ByVal strQry As String)
Dim rs As New ADODB.Recordset
    rs.Open strQry, g_objConn, adOpenDynamic, adLockOptimistic
    s = "("
    For i = 0 To rs.Fields.Count - 1
        s = s & rs.Fields(i).Name & " " & TypeName(rs.Fields(i).Type) & " " & rs.Fields(i).ActualSize
    Next i
    s = s + ")"
    MsgBox "CREATE TABLE " & TableName & s
    'g_objConn.Execute "Create Table " & TableName
End Sub
Function Enkrip(pwd As String) As String
Dim s As String
    For i = 1 To Len(pwd)
        s = s + Chr(Asc(Mid(pwd, i, 1)) + 32)
    Next i
    Enkrip = s
End Function
Function Dekrip(pwd As String) As String
Dim s As String
    For i = 1 To Len(pwd)
        s = s + Chr(Asc(Mid(pwd, i, 1)) - 32)
    Next i
    Dekrip = s
End Function

Function CheckPrivileges(ByVal FormCaption As String, Cmd As CommandType) As Boolean
Dim rs As New ADODB.Recordset
    rs.Open "Select * From USER_PRIVILEGES where FORM_CAPTION='" & FormCaption & "' and USER_NAME='" & g_strUserName & "'", g_objConn, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        Select Case Cmd
        Case ctAddNew
            CheckPrivileges = rs!New
        Case ctSave
            CheckPrivileges = rs!Save
        Case ctDelete
            CheckPrivileges = rs!Delete
        Case ctEdit
            CheckPrivileges = rs!Edit
        Case ctFind
            CheckPrivileges = rs!Find
        Case ctBrowse
            CheckPrivileges = rs!Browse
        Case ctPrint
            CheckPrivileges = rs!Print
        Case ctPrevious
            CheckPrivileges = rs!MovePrevious
        Case ctNext
            CheckPrivileges = rs!MoveNext
        End Select
    Else
        CheckPrivileges = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Sub SetupMenu(User As String)
On Error Resume Next
Dim Y As MyMenu
Dim KosongSemua As Boolean
Dim rs As New ADODB.Recordset
    KosongSemua = True
    For Each X In Desktop.Controls
        If TypeOf X Is MyMenu Then
            For i = 1 To X.Count
                X.SetMenuItemByIndex i, False
            Next i
        End If
    Next X
    KosongSemua = True
    
    rs.Open "Select * From USER_MENU where USER_ID='" & User & "'", g_objConn, adOpenDynamic
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        While Not rs.EOF
            s = rs("MenuName")
            For Each X In Desktop.Controls
                If TypeOf X Is MyMenu Then
                  'For i = 1 To X.Count
                  
                    If Trim(Split(s, "|")(0)) = X.Name Then
                        X.SetMenuItemByCaption Trim(Split(s, "|")(1)), True
                        KosongSemua = False
                    End If
                  'Next i
                End If
            Next X
            rs.MoveNext
        Wend
    End If
    If KosongSemua Then
        Desktop.Start.SetMenuItemByCaption "Admin", True
        Desktop.Start.SetMenuItemByCaption "Logoff", True
        Desktop.Start.SetMenuItemByCaption "Exit", True
    End If
End Sub

Sub FillToListView(ByVal SQL As String, List As ListView, ParamArray Lebar())
Dim rs As New ADODB.Recordset
    List.View = lvwReport
    List.ColumnHeaders.Clear
    rs.Open SQL, g_objConn, adOpenDynamic, adLockReadOnly
    For i = 0 To rs.Fields.Count - 1
        List.ColumnHeaders.Add 1
    Next i
    For i = 0 To rs.Fields.Count - 1
        List.ColumnHeaders(i + 1).Text = rs(i).Name
    Next i
    i = 0
    For Each X In Lebar
        i = i + 1
        If i <= rs.Fields.Count Then List.ColumnHeaders(i).Width = X
    Next
    List.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        While Not rs.EOF
            List.ListItems.Add 1, , IIf(IsNull(rs(0).Value), "", rs(0).Value)
            For i = 1 To rs.Fields.Count - 1
                List.ListItems.Item(1).SubItems(i) = rs(i) & ""
            Next i
            rs.MoveNext
        Wend
    End If
    If List.ListItems.Count > 0 Then
        List.ListItems(1).EnsureVisible
        List.ListItems(1).Selected = True
    End If
    rs.Close
    Set rs = Nothing
End Sub
Sub CreateView(ByVal ViewName As String, ByVal SQL As String)
    g_objConn.Execute "CREATE VIEW " & ViewName & " AS (" & SQL & ")"
End Sub
Sub FillToComboBox(ByVal cb As ComboBox, ByVal Table As String, ByVal Field As String, Optional Criteria As String)
Dim rs As New ADODB.Recordset
    If Criteria = "" Then
        rs.Open "Select " & Field & " as x From " & Table, g_objConn, adOpenDynamic, adLockOptimistic
    Else
        rs.Open "Select " & Field & " as x From " & Table & " Where " & Criteria, g_objConn, adOpenDynamic, adLockOptimistic
    End If
    cb.Clear
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        While Not rs.EOF
            cb.AddItem RTrim(rs("x") & "")
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Sub
Sub FillToListBox(ByVal lb As ListBox, ByVal Table As String, ByVal Field As String, Optional Criteria As String)
Dim rs As New ADODB.Recordset
    If Criteria = "" Then
        rs.Open "Select " & Field & " as x From " & Table, g_objConn, adOpenDynamic, adLockOptimistic
    Else
        rs.Open "Select " & Field & " as x From " & Table & " Where " & Criteria, g_objConn, adOpenDynamic, adLockOptimistic
    End If
    lb.Clear
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        While Not rs.EOF
            lb.AddItem rs("X") & ""
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Sub

Function Search(ByVal Table As String, ByVal Field As String, Optional Criteria As String, Optional FieldTitle As String, Optional Caption As String) As String
Dim f As frmSearch
    Set f = New frmSearch
    f.Table = Table
    f.Field = Field
    f.Criteria = Criteria
    f.FieldTitle = FieldTitle
    f.Title = Caption
    f.Center
    f.Show 1
    Search = Trim(f.Selected)
End Function

Function Find(ByVal Table As String, ByVal Criteria As String, Optional Field As String) As Boolean
'On Error Resume Next
Dim rs As New ADODB.Recordset
Dim SQL As String
    Find = True
    If Field = "" Then
        SQL = "Select * From " & Table & " Where " & Criteria
    Else
        SQL = "Select " & Field & " From " & Table & " Where " & Criteria
    End If
    rs.Open SQL, g_objConn, adOpenDynamic, adLockOptimistic
    Find = Not (rs.EOF And rs.BOF)
    rs.Close
    Set rs = Nothing
End Function

Sub GetFieldValue(ByVal Table As String, ByVal Fields As String, ByVal Criteria As String, ParamArray OutputValues())
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim SQL As String
    If Fields = "" Then
        SQL = "Select * From " & Table & " Where " & Criteria
    Else
        SQL = "Select " & Fields & " From " & Table & " Where " & Criteria
    End If
    rs.Open SQL, g_objConn
    If Not (rs.EOF And rs.BOF) Then
        For i = 0 To rs.Fields.Count - 1
            OutputValues(i) = rs(i).Value
        Next i
    Else
        For i = 0 To UBound(OutputValues)
            OutputValues(i) = ""
        Next
    End If
    rs.Close
    Set rs = Nothing
End Sub
Sub DiEnter(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Sub CenterForm(frm As Form)
    frm.Left = (Screen.Width - frm.Width) \ 2
    frm.Top = (Screen.Height - frm.Height) \ 2
End Sub
Sub ClearText(f As Form)
    For Each X In f
        If (TypeOf X Is TextBox) Or (TypeOf X Is ComboBox) Then
            X.Text = ""
        End If
    Next
End Sub
Sub ClearCheck(f As Form)
    For Each X In f
        If TypeOf X Is CheckBox Then
            X.Value = 0
        End If
    Next
End Sub
Sub EnableTextControl(f As Form, boleh As Boolean, Optional ByVal InActiveColor As Long)
On Error Resume Next
    For Each X In f.Controls
        If (TypeOf X Is ComboBox) Or (TypeOf X Is TextBox) Then
            X.Enabled = boleh
            If boleh Then
                X.BackColor = &HFFFFFF
            Else
                X.BackColor = &HE0E0E0
            End If
        End If
        If (TypeOf X Is OptionButton) Or (TypeOf X Is CheckBox) Then
            X.Enabled = boleh
        End If
    Next
End Sub
Sub ClearAllTextBox(f As Form, boleh As Boolean)
    For Each X In f.Controls
        If (TypeOf X Is TextBox) Or (TypeOf X Is ComboBox) Then
            X.Text = ""
        End If
    Next
End Sub

Sub EnableCmdbutton(f As Form, boleh As Boolean)
    For Each X In f.Controls
        If TypeOf X Is CommandButton Then
            X.Enabled = boleh
        End If
        IsiData.CmdPrev.Enabled = True
        IsiData.CmdNext.Enabled = True
    Next
End Sub
Sub EnableComboControl(f As Form, boleh As Boolean)
    For Each X In f.Controls
        If TypeOf X Is ComboBox Then
            X.Enabled = boleh
            If boleh Then
                X.BackColor = &HFFFFFF
            Else
                X.BackColor = &H8000000F
            End If
        End If
    Next
End Sub
Sub EnableOptionAndCheckBox(f As Form, boleh As Boolean)
    For Each X In f.Controls
        If (TypeOf X Is OptionButton) Or (TypeOf X Is CheckBox) Then
            X.Enabled = boleh
        End If
    Next
End Sub
Sub ClearAllText(f As Form)
    For Each X In f.Controls
        If (TypeOf X Is TextBox) Or (TypeOf X Is ComboBox) Then
            X.Text = ""
        End If
    Next
End Sub
Function UserExist(ByVal UserName As String) As Boolean
Dim rs As New ADODB.Recordset
    rs.Open "Select * From USER_MASTER Where User_Name='" & UserName & "'", g_objConn, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then UserExist = True Else UserExist = False
    rs.Close
    Set rs = Nothing
End Function

Function BuatPanjangString(ByVal s As String, ByVal N As Integer) As String
    While Len(s) < N
        s = s + " "
    Wend
    BuatPanjangString = s
End Function

Function FormatCrptDate(ByVal d As Date)
    FormatCrptDate = "Date(" & Year(d) & "," & Month(d) & "," & Day(d) & ")"
End Function

Function DateBetweenCrpt(ByVal Field As String, ByVal d1 As Date, ByVal d2 As Date)
    DateBetweenCrpt = "{" & Field & "}>=" & FormatCrptDate(d1) & " and {" & Field & "}<=" & FormatCrptDate(d2)
End Function

Function Hari(ByVal Tgl As Date) As String
    h = Choose(Weekday(Tgl, vbMonday), "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu")
    B = Choose(Month(Tgl), "Januari", "Februari", "Maret", "Arpil", "Mei", "Juni", "Juli", "Agustus", "Septembar", "Oktober", "Nopember", "Desember")
    Hari = h & ", " & Day(Tgl) & " " & B & " " & Year(Tgl)
End Function

Sub EnableButton(ByVal f As Form, Baru As Boolean, Edit As Boolean, Simpan As Boolean, Batal As Boolean, Cari As Boolean, Hapus As Boolean, Cetak As Boolean, Tutup As Boolean)
On Error Resume Next
    f.cmdNew.Enabled = Baru
    f.cmdEdit.Enabled = Edit
    f.cmdSave.Enabled = Simpan
    f.cmdCancel.Enabled = Batal
    f.cmdFind.Enabled = Cari
    f.cmdDelete.Enabled = Hapus
    f.cmdPrint.Enabled = Cetak
    f.cmdCLose.Enabled = Tutup
End Sub
Function FileTitle(ByVal FileName As String) As String
    FileTitle = Split(FileName, "\")(UBound(Split(FileName, "\")))
End Function
Function FilePath(ByVal FileName As String) As String
    FilePath = Replace(FileName, FileTitle(FileName), "")
End Function

Sub GantiKodeUnit(ByVal KodeUnit As String)
On Error Resume Next
    KodeUnit = Left(KodeUnit, 2) + "." + Right(KodeUnit, 2)
    
    g_objConn.Execute "Update A set KDA=Left(KDA,2)+'" & KodeUnit + ".'+ Mid(KDA, 9, Len(KDA))"
    
    g_objConn.Execute "Update B set KDA=Left(KDA,2)+'" & KodeUnit + ".'+ Mid(KDA, 9, Len(KDA))"
    g_objConn.Execute "Update B set KDB=Left(KDB,2)+'" & KodeUnit + ".'+ Mid(KDB, 9, Len(KDB))"
    
    g_objConn.Execute "Update C set KDB=Left(KDB,2)+'" & KodeUnit + ".'+ Mid(KDB, 9, Len(KDB))"
    g_objConn.Execute "Update C set KDC=Left(KDC,2)+'" & KodeUnit + ".'+ Mid(KDC, 9, Len(KDC))"
    
    g_objConn.Execute "Update D set KDC=Left(KDC,2)+'" & KodeUnit + ".'+ Mid(KDC, 9, Len(KDC))"
    g_objConn.Execute "Update D set KDD=Left(KDD,2)+'" & KodeUnit + ".'+ Mid(KDD, 9, Len(KDD))"
    
    g_objConn.Execute "Update E set KDD=Left(KDD,2)+'" & KodeUnit + ".'+ Mid(KDD, 9, Len(KDD))"
    g_objConn.Execute "Update E set KDE=Left(KDE,2)+'" & KodeUnit + ".'+ Mid(KDE, 9, Len(KDE))"
    
    
    g_objConn.Execute "Update Uraian set KDA=Left(KDA,2)+'" & KodeUnit + ".'+ Mid(KDA, 9, Len(KDA))"
    g_objConn.Execute "Update Uraian set KDB=Left(KDB,2)+'" & KodeUnit + ".'+ Mid(KDB, 9, Len(KDB))"
    g_objConn.Execute "Update Uraian set KDC=Left(KDC,2)+'" & KodeUnit + ".'+ Mid(KDC, 9, Len(KDC))"
    g_objConn.Execute "Update Uraian set KDD=Left(KDD,2)+'" & KodeUnit + ".'+ Mid(KDD, 9, Len(KDD))"
    g_objConn.Execute "Update Uraian set KDE=Left(KDE,2)+'" & KodeUnit + ".'+ Mid(KDE, 9, Len(KDE))"
    g_objConn.Execute "Update Uraian set KodeRek=Left(KodeRek,2)+'" & KodeUnit + ".'+ Mid(KodeRek, 9, Len(KodeRek))"
    
    g_objConn.Execute "Update Uraian set KodeUnitDinas='" & Replace(KodeUnit, ".", "") & "'"
    g_objConn.Execute "Update UnitDinas set KodeUnitDinas='" & Replace(KodeUnit, ".", "") & "'"
    g_objConn.Execute "Update S2A set KodeUnitDinas='" & Replace(KodeUnit, ".", "") & "'"
    g_objConn.Execute "Update S2 set KodeUnitDinas='" & Replace(KodeUnit, ".", "") & "'"
    g_objConn.Execute "Update S1 set KodeUnitDinas='" & Replace(KodeUnit, ".", "") & "'"
    g_objConn.Execute "Update Indikator set KodeUnitDinas='" & Replace(KodeUnit, ".", "") & "'"
    
End Sub
Sub AmbilNama()
Dim rs As New ADODB.Recordset
        rs.Open "Select * From UnitDinas", g_objConn, adOpenDynamic, adLockReadOnly
        If Not (rs.EOF And rs.BOF) Then
            g_UnitDinas = rs!KodeUnitDinas & ""
            g_NamaUnitDinas = rs!NamaUnitDinas & ""
            g_NamaKepalaDinas = rs!KepalaDinas & ""
            g_NIPKepalaDinas = rs!NIP & ""
            g_Jabatan = rs!Jabatan & ""
            g_Bidang = rs!Bidang & ""
            g_PemegangKas = rs!PemegangKas & ""
            g_JabatanPemegangKas = rs!JabatanPemegangKas & ""
            g_nipPemegangKas = rs!NipPemegangKas & ""
            
            If g_NamaKepalaDinas = "" Then
                MsgBox "Informasi Dinas/Unit Kerja belum lengkap. Silakan tekan ENTER untuk melengkapi informasi dinas/unit kerja", vbInformation, "Info"
                frmSettingUnitKerja.Show 1
            End If
            rs.Close
            rs.Open "Select * From NamaTim", g_objConn, adOpenDynamic, adLockOptimistic
            If Not (rs.EOF And rs.BOF) Then
                For i = 1 To 9
                    NIPAnggaran(i) = rs.Fields("nip" & i) & ""
                Next i
                For i = 1 To 9
                    NMAnggaran(i) = rs.Fields("nm" & i) & ""
                Next i
                For i = 1 To 9
                    JBTAnggaran(i) = rs.Fields("jbt" & i) & ""
                Next i
            End If
        Else
            g_UnitDinas = ""
            g_NamaUnitDinas = ""
            g_NamaKepalaDinas = ""
            g_NIPKepalaDinas = ""
            g_Jabatan = ""
            g_Bidang = ""
            g_PemegangKas = ""
            g_JabatanPemegangKas = ""
            g_nipPemegangKas = ""
            MsgBox "Informasi Dinas/Unit Kerja belum diisi. Silakan tekan ENTER untuk mengisi data dinas/unit kerja", vbInformation, "Info"
            frmSettingUnitKerja.Show 1
        End If
        rs.Close
        Set rs = Nothing
End Sub
Function TableExist(ByVal s As String) As Boolean
On Error GoTo Err1
Dim rs As New ADODB.Recordset
    rs.Open "Select * from " & s, g_objConn, adOpenDynamic, adLockOptimistic
    rs.Close
    Set rs = Nothing
    TableExist = True
    Exit Function
Err1:
    TableExist = False
End Function

Sub CreateTables()
On Error Resume Next
Dim dB As Database
Dim tb As TableDef
Dim fn As Field
Dim rs As Recordset
Dim temp As String
Dim idx As Index

    Set dB = DBEngine.OpenDatabase(g_strDBName)
    Set tb = dB.CreateTableDef("IndikatorPerubahan")
    With tb
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("MASUKKAN", dbText, 225)
        .Fields.Append .CreateField("KELUARAN", dbText, 225)
        .Fields.Append .CreateField("HASIL", dbText, 225)
        .Fields.Append .CreateField("MANFAAT", dbText, 225)
        .Fields.Append .CreateField("DAMPAK", dbText, 225)
        .Fields.Append .CreateField("KINERJA1", dbText, 225)
        .Fields.Append .CreateField("KINERJA2", dbText, 225)
        .Fields.Append .CreateField("KINERJA3", dbText, 225)
        .Fields.Append .CreateField("KINERJA4", dbText, 225)
        .Fields.Append .CreateField("KINERJA5", dbText, 225)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("NAMAPIMPINAN", dbText, 50)
        .Fields.Append .CreateField("NIP", dbText, 15)
        .Fields.Append .CreateField("JABATAN", dbText, 50)
        .Fields.Append .CreateField("LokasiKegiatan", dbText, 120)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX GABUNG ON IndikatorPerubahan (TH,KEGIATAN)"
    
    Set tb = dB.CreateTableDef("S2APerubahan")
    With tb
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("PROGRAM", dbText, 225)
        .Fields.Append .CreateField("JUMLAH", dbCurrency)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX GABUNG ON S2APerubahan (TH,KEGIATAN)"
    
    Set tb = dB.CreateTableDef("S2Perubahan")
    With tb
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("PROGRAM", dbText, 225)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX Indeks ON S2Perubahan (TH,PROGRAM,KodeUnitDinas)"
    
    
    Set tb = dB.CreateTableDef("TempIndikator")
    With tb
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("MASUKKAN", dbText, 225)
        .Fields.Append .CreateField("KELUARAN", dbText, 225)
        .Fields.Append .CreateField("HASIL", dbText, 225)
        .Fields.Append .CreateField("MANFAAT", dbText, 225)
        .Fields.Append .CreateField("DAMPAK", dbText, 225)
        .Fields.Append .CreateField("KINERJA1", dbText, 50)
        .Fields.Append .CreateField("KINERJA2", dbText, 50)
        .Fields.Append .CreateField("KINERJA3", dbText, 50)
        .Fields.Append .CreateField("KINERJA4", dbText, 50)
        .Fields.Append .CreateField("KINERJA5", dbText, 50)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("NAMAPIMPINAN", dbText, 50)
        .Fields.Append .CreateField("NIP", dbText, 15)
        .Fields.Append .CreateField("JABATAN", dbText, 50)
        .Fields.Append .CreateField("LokasiKegiatan", dbText, 120)
        .Fields.Append .CreateField("Jenis", dbText, 1)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX GABUNG ON TempIndikator (TH,KEGIATAN,KodeUnitDinas,Jenis)"
    
    Set tb = dB.CreateTableDef("TempS2")
    With tb
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("PROGRAM", dbText, 225)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("Jenis", dbText, 1)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX Indeks ON TempS2 (TH,Program,KodeUnitDinas,Jenis)"
    
    
    Set tb = dB.CreateTableDef("TempS2A")
    With tb
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("PROGRAM", dbText, 225)
        .Fields.Append .CreateField("JUMLAH", dbCurrency)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("Jenis", dbText, 1)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX GABUNG ON TempS2A (TH,KEGIATAN,Jenis,KodeUnitDinas)"
    
    
    'Uraian
    Set tb = dB.CreateTableDef("TempUraian")
    With tb
        .Fields.Append .CreateField("URAIAN", dbText, 225)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("VOL", dbLong)
        .Fields.Append .CreateField("SAT", dbText, 15)
        .Fields.Append .CreateField("HGSAT", dbCurrency)
        .Fields.Append .CreateField("JLF", dbCurrency)
        .Fields.Append .CreateField("KodeRek", dbText, 21)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("NoUrut", dbSingle)
        .Fields.Append .CreateField("KDA", dbText, 21)
        .Fields.Append .CreateField("KDB", dbText, 21)
        .Fields.Append .CreateField("KDC", dbText, 21)
        .Fields.Append .CreateField("KDD", dbText, 21)
        .Fields.Append .CreateField("KDE", dbText, 21)
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("JLFPerubahan", dbCurrency)
        .Fields.Append .CreateField("JENIS", dbText, 1)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then tb.Fields(i).AllowZeroLength = True
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX IND ON TempUraian (kde,KodeUnitDinas,Kegiatan,NoUrut,Jenis)"
    
    
    Set tb = dB.CreateTableDef("UraianPerubahan")
    With tb
        .Fields.Append .CreateField("URAIAN", dbText, 225)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("VOL", dbLong)
        .Fields.Append .CreateField("SAT", dbText, 15)
        .Fields.Append .CreateField("HGSAT", dbCurrency)
        .Fields.Append .CreateField("JLF", dbCurrency)
        .Fields.Append .CreateField("KodeRek", dbText, 21)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("NoUrut", dbSingle)
        .Fields.Append .CreateField("KDA", dbText, 21)
        .Fields.Append .CreateField("KDB", dbText, 21)
        .Fields.Append .CreateField("KDC", dbText, 21)
        .Fields.Append .CreateField("KDD", dbText, 21)
        .Fields.Append .CreateField("KDE", dbText, 21)
        .Fields.Append .CreateField("TH", dbText, 4)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then
            If (UCase(Left(tb.Fields(i).Name, 2)) = "KD") Or (UCase(tb.Fields(i).Name) = "KEGIATAN") Then tb.Fields(i).DefaultValue = "x"
            tb.Fields(i).AllowZeroLength = True
        End If
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX IND ON UraianPerubahan (th,KodeRek,KodeUnitDinas,NoUrut,Kegiatan,KDA,KDB,KDC,KDD,KDE)"
    
    Set tb = dB.CreateTableDef("TempLaporanUraian")
    With tb
        .Fields.Append .CreateField("URAIAN", dbText, 225)
        .Fields.Append .CreateField("KEGIATAN", dbText, 225)
        .Fields.Append .CreateField("VOL", dbLong)
        .Fields.Append .CreateField("SAT", dbText, 15)
        .Fields.Append .CreateField("HGSAT", dbCurrency)
        .Fields.Append .CreateField("JLF", dbCurrency)
        .Fields.Append .CreateField("KodeRek", dbText, 21)
        .Fields.Append .CreateField("KodeUnitDinas", dbText, 5)
        .Fields.Append .CreateField("NoUrut", dbSingle)
        .Fields.Append .CreateField("KDA", dbText, 21)
        .Fields.Append .CreateField("KDB", dbText, 21)
        .Fields.Append .CreateField("KDC", dbText, 21)
        .Fields.Append .CreateField("KDD", dbText, 21)
        .Fields.Append .CreateField("KDE", dbText, 21)
        .Fields.Append .CreateField("TH", dbText, 4)
        .Fields.Append .CreateField("URAIANPERUBAHAN", dbText, 225)
        .Fields.Append .CreateField("VOLPERUBAHAN", dbLong)
        .Fields.Append .CreateField("SATPERUBAHAN", dbText, 15)
        .Fields.Append .CreateField("HGSATPERUBAHAN", dbCurrency)
        .Fields.Append .CreateField("JLFPERUBAHAN", dbCurrency)
    End With
    For i = 0 To tb.Fields.Count - 1
        If tb.Fields(i).Type = dbText Then
            If (UCase(Left(tb.Fields(i).Name, 2)) = "KD") Or (UCase(tb.Fields(i).Name) = "KEGIATAN") Then tb.Fields(i).DefaultValue = "x"
            tb.Fields(i).AllowZeroLength = True
        End If
    Next i
    dB.TableDefs.Append tb
    dB.Execute "CREATE INDEX IND ON TempLaporanUraian (kde,KodeUnitDinas,Kegiatan,NoUrut,Jenis)"
    
    dB.Close
    Set dB = Nothing
End Sub
Function FieldAda(ByVal Tabel As String, ByVal Field As String) As Boolean
On Error GoTo ErrF1
Dim rs As New ADODB.Recordset
    rs.Open "Select " & Field & " From " & Tabel, g_objConn, adOpenDynamic, adLockOptimistic
    FieldAda = True
    Exit Function
ErrF1:
    FieldAda = False
End Function

Sub SiapkanLaporan(ProgressBar1 As ProgressBar)
Dim rs As New ADODB.Recordset
Dim rsA As New ADODB.Recordset
Dim rsP As New ADODB.Recordset
Dim rsBaru As New ADODB.Recordset

    ProgressBar1.Value = 0
    
    g_objConn.Execute "Delete From TempLaporanUraian"
    pos = 0
    
    If rs.State = 1 Then rs.Close
    rs.Open "Select KodeUnitDinas,KodeRek,Kegiatan from TempUraian Group By KodeUnitDinas,KodeRek,Kegiatan", g_objConn, adOpenDynamic, adLockOptimistic
    N = 0
    While Not rs.EOF
        rs.MoveNext
        N = N + 1
    Wend
    rs.MoveFirst
    
    While Not rs.EOF
        If rsA.State = 1 Then rsA.Close
        rsA.Open "Select count(Uraian) as TJumlah from TempUraian where Jenis='A' and KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
        
        If rsP.State = 1 Then rsP.Close
        rsP.Open "Select count(Uraian) as TJumlah from TempUraian where Jenis='P' and KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
        
        If rsA.Fields("TJumlah") > rsP.Fields("TJumlah") Then
            rsA.Close
            rsA.Open "Select * from TempUraian where Jenis='A' and KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
            
            If rsBaru.State = 1 Then rsBaru.Close
            rsBaru.Open "Select * from TempLaporanUraian", g_objConn, adOpenDynamic, adLockOptimistic
            
            While Not rsA.EOF
                rsBaru.AddNew
                rsBaru.Fields("KodeUnitDinas") = rs!KodeUnitDinas
                rsBaru.Fields("KodeRek") = rs!KodeRek
                rsBaru.Fields("Kegiatan") = rs!Kegiatan
                rsBaru.Fields("Uraian") = rsA!Uraian
                rsBaru.Fields("Vol") = rsA!Vol
                rsBaru.Fields("Sat") = rsA!Sat
                rsBaru.Fields("HGSAT") = rsA!HGSAT
                rsBaru.Fields("JLF") = rsA!JLF
                rsBaru.Fields("KDA") = rsA!KDA
                rsBaru.Fields("KDB") = rsA!KDB
                rsBaru.Fields("KDC") = rsA!KDC
                rsBaru.Fields("KDD") = rsA!KDD
                rsBaru.Fields("KDE") = rsA!KDE
                rsBaru.Fields("NoUrut") = rsA!NoUrut
                rsBaru.Fields("TH") = rsA!Th
                'rsBaru.Fields("DasarHukum") = rsA!DasarHukum
                rsBaru.Update
                rsA.MoveNext
            Wend
            
            rsP.Close
            rsP.Open "Select * from TempUraian where Jenis='P' and KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
            
            rsBaru.Close
            rsBaru.Open "Select * from TempLaporanUraian where KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
            
            While Not rsP.EOF
                rsBaru.Fields("UraianPerubahan") = rsP!Uraian
                rsBaru.Fields("VolPerubahan") = rsP!Vol
                rsBaru.Fields("SatPerubahan") = rsP!Sat
                rsBaru.Fields("HGSATPerubahan") = rsP!HGSAT
                rsBaru.Fields("JLFPerubahan") = IIf(rsP!JLFPerubahan <> 0, rsP!JLFPerubahan, rsP!JLF)
                rsBaru.Update
                rsBaru.MoveNext
                rsP.MoveNext
            Wend
            
        Else
            
            rsP.Close
            rsP.Open "Select * from TempUraian where Jenis='P' and KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
            
            If rsBaru.State = 1 Then rsBaru.Close
            rsBaru.Open "Select * from TempLaporanUraian", g_objConn, adOpenDynamic, adLockOptimistic
            
            While Not rsP.EOF
                rsBaru.AddNew
                rsBaru.Fields("KodeUnitDinas") = rs!KodeUnitDinas
                rsBaru.Fields("KodeRek") = rs!KodeRek
                rsBaru.Fields("Kegiatan") = rs!Kegiatan
                rsBaru.Fields("UraianPerubahan") = rsP!Uraian
                rsBaru.Fields("VolPerubahan") = rsP!Vol
                rsBaru.Fields("SatPerubahan") = rsP!Sat
                rsBaru.Fields("HGSATPerubahan") = rsP!HGSAT
                rsBaru.Fields("JLFPerubahan") = rsP!JLFPerubahan
                rsBaru.Fields("KDA") = rsP!KDA
                rsBaru.Fields("KDB") = rsP!KDB
                rsBaru.Fields("KDC") = rsP!KDC
                rsBaru.Fields("KDD") = rsP!KDD
                rsBaru.Fields("KDE") = rsP!KDE
                rsBaru.Fields("NoUrut") = rsP!NoUrut
                rsBaru.Fields("TH") = rsP!Th
                'rsBaru.Fields("DasarHukum") = rsP!DasarHukum
                rsBaru.Update
                rsP.MoveNext
            Wend
            
            rsA.Close
            rsA.Open "Select * from TempUraian where Jenis='A' and KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
            
            rsBaru.Close
            rsBaru.Open "Select * from TempLaporanUraian where KodeUnitDinas='" & rs!KodeUnitDinas & "' " & _
                 "And KodeRek='" & rs!KodeRek & "' and Kegiatan='" & rs!Kegiatan & "'", g_objConn, adOpenDynamic, adLockOptimistic
            
            While Not rsA.EOF
                rsBaru.Fields("Uraian") = rsA!Uraian
                rsBaru.Fields("Vol") = rsA!Vol
                rsBaru.Fields("Sat") = rsA!Sat
                rsBaru.Fields("HGSAT") = rsA!HGSAT
                rsBaru.Fields("JLF") = rsA!JLF
                
                rsBaru.Update
                rsBaru.MoveNext
                rsA.MoveNext
            Wend
        End If
        
        pos = pos + 1
        ProgressBar1.Value = (pos / N) * 100
        rs.MoveNext
    Wend
End Sub


