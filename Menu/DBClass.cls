VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'====================================================
'PROJECT NAME   : DBFunction Class Modul
'PROGRAMMER     : Agusman Yulianto, S.Si, M.Kom
'LAST MODIFIED  : 9 July 2004
'====================================================

Private mvarConnectionString As String
Public DbConnection As New ADODB.Connection
Public PersenSelesaiExport As Single
Private Function TambahSpasi(ByVal s As String, ByVal N As Integer) As String
    While Len(s) < N
        s = s + Space(1)
    Wend
    TambahSpasi = s
End Function
Public Sub EksportKeFileTeks(ByVal NamaFile As String, ByVal NamaTabel As String, Optional Fields As String, Optional LebarFields As String, Optional ByVal Kriteria As String, Optional ByVal TandaPemisah As String)
Dim rs As New ADODB.Recordset
Dim FileNumber As Integer
Dim FieldSize() As String
Dim N As Long
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Sub
    End If
    If Trim(LebarFields) <> "" Then
        FieldSize = Split(LebarFields, ",")
    Else
        ReDim FieldSize(0)
    End If
    If Trim(Fields) = "" Then Fields = "*"
    If Trim(Kriteria) <> "" Then
        rs.Open "Select " & Fields & " From " & NamaTabel & " Where " & Kriteria, DbConnection, adOpenStatic, adLockReadOnly
    Else
        rs.Open "Select " & Fields & " From " & NamaTabel, DbConnection, adOpenStatic, adLockReadOnly
    End If
    N = rs.RecordCount: Pos = 0
    If Not (rs.EOF And rs.BOF) Then
        Open NamaFile For Output As 1
        While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                If i <> rs.Fields.Count - 1 Then
                    If UBound(FieldSize) <> 0 Then
                        If Val(FieldSize(i)) <> 0 Then
                            Print #1, TambahSpasi(rs.Fields(i), FieldSize(i)) & "" & TandaPemisah;
                        Else
                            Print #1, rs.Fields(i) & "" & TandaPemisah;
                        End If
                    Else
                        Print #1, rs.Fields(i) & "" & TandaPemisah;
                    End If
                Else
                    Print #1, rs.Fields(i) & ""
                End If
            Next i
            Pos = Pos + 1
            PersenSelesaiExport = (Pos / N) * 100
            DoEvents
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Sub
Public Sub EksportSQLKeFileTeks(ByVal NamaFile As String, ByVal SQL As String, Optional LebarFields As String, Optional ByVal TandaPemisah As String)
Dim rs As New ADODB.Recordset
Dim FileNumber As Integer
Dim FieldSize() As String
Dim N As Long
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Sub
    End If
    If Trim(LebarFields) <> "" Then
        FieldSize = Split(LebarFields, ",")
    Else
        ReDim FieldSize(0)
    End If
    If Trim(Fields) = "" Then Fields = "*"
    rs.Open SQL, DbConnection, adOpenStatic, adLockReadOnly
    N = rs.RecordCount: Pos = 0
    If Not (rs.EOF And rs.BOF) Then
        Open NamaFile For Output As 1
        While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                If i <> rs.Fields.Count - 1 Then
                    If UBound(FieldSize) <> 0 Then
                        If Val(FieldSize(i)) <> 0 Then
                            Print #1, TambahSpasi(rs.Fields(i), FieldSize(i)) & "" & TandaPemisah;
                        Else
                            Print #1, rs.Fields(i) & "" & TandaPemisah;
                        End If
                    Else
                        Print #1, rs.Fields(i) & "" & TandaPemisah;
                    End If
                Else
                    Print #1, rs.Fields(i) & ""
                End If
            Next i
            Pos = Pos + 1
            PersenSelesaiExport = (Pos / N) * 100
            DoEvents
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Property Let ConnectionString(ByVal vData As String)
On Error GoTo Err4
    mvarConnectionString = vData
    With DbConnection
        .ConnectionString = ConnectionString
        .ConnectionTimeout = 5
        .Open
    End With
    Exit Property
Err4:
    MsgBox Err.Description, vbInformation, "Error"
End Property

Public Property Get ConnectionString() As String
    ConnectionString = mvarConnectionString
End Property

Public Function JumlahRekord(ByVal NamaTabel As String, Optional Kriteria As String, Optional ByVal NumberFormat As String)
On Error GoTo Err1
Dim rs As New ADODB.Recordset
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    If Trim(Kriteria) <> "" Then
        rs.Open "Select * From " & NamaTabel & " Where " & Kriteria, DbConnection, adOpenStatic, adLockReadOnly
    Else
        rs.Open "Select * From " & NamaTabel, DbConnection, adOpenStatic, adLockReadOnly
    End If
    If Trim(NumberFormat) <> "" Then
        JumlahRekord = Format(IIf(IsNull(rs.RecordCount), 0, rs.RecordCount), NumberFormat)
    Else
        JumlahRekord = IIf(IsNull(rs.RecordCount), 0, rs.RecordCount)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
Err1:
    MsgBox Err.Description, vbInformation, "Error"
End Function

Private Function ConnectionOK() As Boolean
    ConnectionOK = (DbConnection.State = 1)
End Function

Public Function HapusRekord(ByVal NamaTabel As String, ByVal Kriteria As String, Optional ByVal Pesan As String, Optional ByVal PesanJikaBerhasil As String)
On Error GoTo Err2
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    If Trim(Kriteria) = "" Then
        MsgBox "You must fill Kriteria expression", vbInformation, "Info"
        Exit Function
    End If
    If Trim(Pesan) <> "" Then
        Y = MsgBox(Pesan, 36, "Confirmation")
        If Y = 6 Then
            DbConnection.Execute "Delete From " & NamaTabel & " Where " & Kriteria
            If Trim(PesanJikaBerhasil) <> "" Then MsgBox PesanJikaBerhasil, vbInformation, "Info"
        End If
    Else
        DbConnection.Execute "Delete From " & NamaTabel & " Where " & Kriteria
        If Trim(PesanJikaBerhasil) <> "" Then MsgBox PesanJikaBerhasil, vbInformation, "Info"
    End If
    Exit Function
Err2:
    MsgBox Err.Description, vbInformation, "Error"
End Function


'========================
'Specific Function
'========================
Public Function ParCharReferensi1(ByVal RefKey1 As String, ByVal RefKey2 As String) As String
On Error GoTo Err3
Dim rs As New ADODB.Recordset
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    
    rs.Open "Select CharPar1 From Refference where RefKey1='" & RefKey1 & "' and RefKey2='" & RefKey2 & "'", DbConnection, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        ParCharReferensi1 = rs!CharPar1 & ""
    Else
        ParCharReferensi1 = ""
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
Err3:
    MsgBox Err.Description, vbInformation, "Error"
End Function

Public Function KodePropinsi(ByVal NamaPropinsi As String)
    Dim rs As New ADODB.Recordset
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    rs.Open "Select Kd_Propinsi From Propinsi where Nm_Propinsi='" & NamaPropinsi & "'", DbConnection, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        KodePropinsi = rs!Kd_Propinsi & ""
    Else
        KodePropinsi = ""
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
Err3:
    MsgBox Err.Description, vbInformation, "Error"
End Function

Public Function KodeKabupaten(ByVal NamaKabupaten As String)
    Dim rs As New ADODB.Recordset
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    rs.Open "Select Kd_Kabupaten From Kabupaten where Nm_Kabupaten='" & NamaKabupaten & "'", DbConnection, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        KodeKabupaten = rs!Kd_Kabupaten & ""
    Else
        KodeKabupaten = ""
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
Err3:
    MsgBox Err.Description, vbInformation, "Error"
End Function

Public Function KodeKecamatan(ByVal NamaKecamatan As String)
    Dim rs As New ADODB.Recordset
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    rs.Open "Select Kd_Kecamatan From Kecamatan where Nm_Kecamatan='" & NamaKecamatan & "'", DbConnection, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        KodeKecamatan = rs!Kd_Kecamatan & ""
    Else
        KodeKecamatan = ""
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
Err3:
    MsgBox Err.Description, vbInformation, "Error"
End Function

Public Function KodeDesa(ByVal NamaDesa As String)
    Dim rs As New ADODB.Recordset
    If Not ConnectionOK Then
        MsgBox "Koneksi ke database belum di set ", vbInformation, "Info"
        Exit Function
    End If
    rs.Open "Select Kd_Desa From Desa where Nm_Desa='" & NamaDesa & "'", DbConnection, adOpenDynamic, adLockOptimistic
    If Not (rs.EOF And rs.BOF) Then
        KodeDesa = rs!Kd_Desa & ""
    Else
        KodeDesa = ""
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
Err3:
    MsgBox Err.Description, vbInformation, "Error"
End Function
