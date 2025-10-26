VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_D_出荷・移動データ入力"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim frmWidth As Long
Dim frmHeight As Long
Dim subFrmWidth As Long
Dim subFrmHeight As Long

Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Form_Open
    
    PreForm = AFNAME
    AFNAME = Me.name

    Application.SetOption KEYMACRO, "入力キー"
'    Me.Visible = True

    Call 初期設定(Me.name)
    
    frmWidth = Me.WindowWidth
    frmHeight = Me.WindowHeight
    subFrmWidth = Me.SUB.Width
    subFrmHeight = Me.SUB.Height
    
    Call FRMCLS

Exit_Form_Open:
    Exit Sub
    
Err_Form_Open:
    MsgBox Err.Description, , "Form_Open"
    Call エラーログ記録(Me.name, "Form_Open", Err.Number, Err.Description)
    Resume Exit_Form_Open
End Sub

Private Sub Form_Resize()
On Error Resume Next

'    Me.SUB.Width = subFrmWidth + (Me.WindowWidth - frmWidth)
'    DoEvents
    Me.SUB.Height = subFrmHeight + (Me.WindowHeight - frmHeight)
    DoEvents
End Sub

Public Function FRMCLS()
On Error Resume Next

    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
           ctl = Null
        ElseIf ctl.ControlType = acComboBox Then
           ctl = Null
        ElseIf ctl.ControlType = acCheckBox Then
           ctl = 0
        End If
    Next

    Set ctl = Nothing
    
    Call モード変更(Me.MODE, "新規")
    
    Me.工場元CD = FACTORY
    Me.処理日 = Date
    Me.日付 = Date
    Me.読込FLG = False
    
    Me.処理区分 = 3
    Call 処理区分_AfterUpdate
    
    Call D_DEL(CurrentDb, "W_出荷・移動データ入力_移動")
    Call D_DEL(CurrentDb, "W_出荷・移動データ入力_出荷")
    
    Me.SUB.Form.Requery
    
    Call FORMFOCUS
End Function

Function FORMFOCUS()
On Error GoTo Err_FORMFOCUS
    
    Me.P5.Enabled = False
    Me.P6.Enabled = False
    Me.P7.Enabled = False
    Me.P8.Enabled = False
    Me.P9.Enabled = False
    Me.P10.Enabled = False
    
    Me.処理区分.Enabled = False
    Me.工場元CD.Enabled = False
    Me.工場先CD.Enabled = False
    Me.出荷先CD.Enabled = False
    Me.担当者CD.Enabled = False
    Me.処理日.Enabled = False
    Me.日付.Enabled = False
    Me.伝票番号.Enabled = False
    
    Me.印刷済み1.Visible = False
    Me.印刷済み2.Visible = False
    Me.印刷済み3.Visible = False
    Me.MSG1.Visible = False
    Me.MSG2.Visible = False
    Me.MSG3.Visible = False

    If Nz(Me.MODE, "") = "" Then
    ElseIf Nz(Me.MODE, "") = "新規" Then
        Me.P5.Enabled = True
        Me.P7.Enabled = True
        Me.P8.Enabled = True
        Me.P9.Enabled = True
        Me.P10.Enabled = True
        
        Me.処理区分.Enabled = True
        Me.工場元CD.Enabled = True
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            Me.工場先CD.Enabled = True
        Else
            Me.出荷先CD.Enabled = True
        End If
        
        Me.担当者CD.Enabled = True
        Me.処理日.Enabled = True
        Me.日付.Enabled = True
        
        Me.SUB.Form.処理連番.Locked = False
    ElseIf Nz(Me.MODE, "") = "修正" Then
        Me.P5.Enabled = True
        Me.P6.Enabled = True
        Me.P8.Enabled = True
        Me.P9.Enabled = True
        Me.P10.Enabled = True
        
        Me.処理区分.Enabled = True
        Me.工場元CD.Enabled = True
        Me.処理日.Enabled = True
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            Me.工場先CD.Enabled = True
        Else
            Me.印刷済み1.Visible = True
            Me.印刷済み2.Visible = True
            Me.印刷済み3.Visible = True
            Me.MSG1.Visible = True
            Me.MSG2.Visible = True
            Me.MSG3.Visible = True
        End If
        Me.伝票番号.Enabled = True
        
        Me.SUB.Form.処理連番.Locked = False
    ElseIf Nz(Me.MODE, "") = "削除" Then
        Me.P5.Enabled = True
        Me.P6.Enabled = True
        Me.P7.Enabled = True
        
        Me.処理区分.Enabled = True
        Me.工場元CD.Enabled = True
        Me.処理日.Enabled = True
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            Me.工場先CD.Enabled = True
        End If
        Me.伝票番号.Enabled = True
        
        Me.SUB.Form.処理連番.Locked = True
    End If
    
    Me.工場元CD.SetFocus

Exit_FORMFOCUS:
    Exit Function
Err_FORMFOCUS:
    MsgBox Err.Description, , "FORMFOCUS"
    Call エラーログ記録(Me.name, "FORMFOCUS", Err.Number, Err.Description)
    Resume Exit_FORMFOCUS
End Function

Public Function DATAREAD1()
On Error GoTo Err_DATAREAD1
    
    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim Sql As String
    Dim 行番号 As Long

    DATAREAD1 = False
    DoCmd.Hourglass True
    
    Set MyDB = CurrentDb
    
    Call D_DEL(MyDB, "W_出荷・移動データ入力_移動")
    
    Set WKDB = MyDB.OpenRecordset("W_出荷・移動データ入力_移動")
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    Sql = ""
    Sql = Sql & " SELECT SD17.*, "
    Sql = Sql & " SM04.* "
    Sql = Sql & " FROM KSD17IDOU AS SD17 "
    Sql = Sql & " LEFT JOIN KSM04SEIH AS SM04 ON (SD17.SD17011 = SM04.SM04001 AND SD17.SD17012 = SM04.SM04002) "
    Sql = Sql & " WHERE SD17.SD17002 = " & SQL_MOJICHK(Me.伝票番号)
    Sql = Sql & " ORDER BY SD17.SD17003 "
    
    rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
    
    If rst.EOF Then
        MsgBox "データが1件もありません。", vbOKOnly, "確認"
        
        Call FRMCLS
        
        ' DELETE 2025/02/28 SUZUKI フォーカスエラー発生のため
'        Me.伝票番号.SetFocus

        GoTo Exit_DATAREAD1
    End If
    
    行番号 = 1
    Do Until rst.EOF
        ' 1行目の情報をヘッダーに表示
        If rst!SD17003 = 1 Then
            Me.工場元CD = rst!SD17005
            Me.工場先CD = rst!SD17007
            If Trim(Nz(rst!SD17010, "")) <> "" Then
                Me.担当者CD = rst!SD17010
            End If
            Me.処理日 = DateSerial(Left(rst!SD17008, 4), Mid(rst!SD17008, 5, 2), Right(rst!SD17008, 2))
            Me.日付 = DateSerial(Left(rst!SD17006, 4), Mid(rst!SD17006, 5, 2), Right(rst!SD17006, 2))
        End If
        
        WKDB.AddNew
            
            WKDB!SD17001 = rst!SD17001
            WKDB!SD17002 = rst!SD17002
            WKDB!SD17003 = 行番号
            WKDB!SD17004 = rst!SD17004
            WKDB!SD17005 = rst!SD17005
            WKDB!SD17006 = rst!SD17006
            WKDB!SD17007 = rst!SD17007
            WKDB!SD17008 = rst!SD17008
            WKDB!SD17009 = rst!SD17009
            WKDB!SD17010 = rst!SD17010
            WKDB!SD17011 = rst!SD17011
            WKDB!SD17012 = rst!SD17012
            WKDB!SD17013 = rst!SD17013
            WKDB!SD17014 = rst!SD17014
            WKDB!SD17015 = rst!SD17015
            WKDB!SD17016 = rst!SD17016
            WKDB!SD17017 = rst!SD17017
            WKDB!SD17018 = rst!SD17018 * 1000
            WKDB!SD17019 = rst!SD17019
            WKDB!SD17020 = rst!SD17020
            WKDB!SD17021 = rst!SD17021
            WKDB!SD17022 = rst!SD17022
            
            WKDB!登録日 = rst!登録日
            WKDB!変更日 = rst!変更日
            
            WKDB!製品名 = rst!SM04003
            
            WKDB!登録済 = True
            
            WKDB!旧処理連番 = rst!SD17001
            
        WKDB.Update
        
        If 行番号 Mod LINEUNIT = 0 Then
            Me.SUB.Form.Requery
            DoEvents
        End If
        
        行番号 = 行番号 + 1
        
        rst.MoveNext
    Loop
    
    Me.読込FLG = True
    
    Call FORMFOCUS
    
    DATAREAD1 = True
    
Exit_DATAREAD1:
    Me.SUB.Form.Requery
    Me.SUB.Form.処理連番.SetFocus
    
    Call CloseADORecordset(rst)
    Call CloseADOConnection(cnn)
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    DoCmd.Hourglass False
    Exit Function

Err_DATAREAD1:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "DATAREAD1"
    Call エラーログ記録(Me.name, "DATAREAD1", Err.Number, Err.Description, "Sql：" & Sql)
    Resume Exit_DATAREAD1
End Function

Public Function DATAREAD2()
On Error GoTo Err_DATAREAD2
    
    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim Sql As String
    Dim 行番号 As Long

    DATAREAD2 = False
    DoCmd.Hourglass True
    
    Set MyDB = CurrentDb
    
    Call D_DEL(MyDB, "W_出荷・移動データ入力_出荷")
    
    Set WKDB = MyDB.OpenRecordset("W_出荷・移動データ入力_出荷")
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    Sql = ""
    Sql = Sql & " SELECT SD18.*, "
    Sql = Sql & " SM04.* "
    Sql = Sql & " FROM KSD18SYUK AS SD18 "
    Sql = Sql & " LEFT JOIN KSM04SEIH AS SM04 ON (SD18.SD18008 = SM04.SM04001 AND SD18.SD18009 = SM04.SM04002) "
    Sql = Sql & " WHERE SD18.SD18005 = " & SQL_MOJICHK(Me.伝票番号)
    Sql = Sql & " ORDER BY SD18.SD18006 "
    
    rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
    
    If rst.EOF Then
        MsgBox "データが1件もありません。", vbOKOnly, "確認"
        
        ' DELETE 2025/02/28 SUZUKI フォーカスエラー発生のため
'        Me.伝票番号.SetFocus

        GoTo Exit_DATAREAD2
    End If
    
    行番号 = 1
    Do Until rst.EOF
        ' 1行目の情報をヘッダーに表示
        If rst!SD18006 = 1 Then
            Me.工場元CD = rst!SD18002
            Me.出荷先CD = rst!SD18004
            If Trim(Nz(rst!SD18007, "")) <> "" Then
                Me.担当者CD = rst!SD18007
            End If
            Me.処理日 = DateSerial(Left(rst!SD18018, 4), Mid(rst!SD18018, 5, 2), Right(rst!SD18018, 2))
            Me.日付 = DateSerial(Left(rst!SD18003, 4), Mid(rst!SD18003, 5, 2), Right(rst!SD18003, 2))
        End If
        
        WKDB.AddNew
            
            WKDB!SD18001 = rst!SD18001
            WKDB!SD18002 = rst!SD18002
            WKDB!SD18003 = rst!SD18003
            WKDB!SD18004 = rst!SD18004
            WKDB!SD18005 = rst!SD18005
            WKDB!SD18006 = 行番号
            WKDB!SD18007 = rst!SD18007
            WKDB!SD18008 = rst!SD18008
            WKDB!SD18009 = rst!SD18009
            WKDB!SD18010 = rst!SD18010
            WKDB!SD18011 = rst!SD18011
            WKDB!SD18012 = rst!SD18012
            WKDB!SD18013 = rst!SD18013
            WKDB!SD18014 = rst!SD18014
            WKDB!SD18015 = rst!SD18015 * 1000
            WKDB!SD18016 = rst!SD18016
            WKDB!SD18017 = rst!SD18017
            WKDB!SD18018 = rst!SD18018
            WKDB!SD18019 = rst!SD18019
            WKDB!SD18020 = rst!SD18020
            WKDB!SD18021 = rst!SD18021
            WKDB!SD18022 = rst!SD18022
            WKDB!SD18023 = rst!SD18023
            
            WKDB!登録日 = rst!登録日
            WKDB!変更日 = rst!変更日
            
            WKDB!製品名 = rst!SM04003
            
            WKDB!登録済 = True
            
            WKDB!旧処理連番 = rst!SD18001
            
        WKDB.Update
        
        If 行番号 Mod LINEUNIT = 0 Then
            Me.SUB.Form.Requery
            DoEvents
        End If
        
        行番号 = 行番号 + 1
        
        rst.MoveNext
    Loop
    
    Me.読込FLG = True
    
    Call FORMFOCUS
    
    DATAREAD2 = True
    
Exit_DATAREAD2:
    Me.SUB.Form.Requery
    Me.SUB.Form.処理連番.SetFocus
    
    Call 未発行更新
    
    Call CloseADORecordset(rst)
    Call CloseADOConnection(cnn)
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    DoCmd.Hourglass False
    Exit Function

Err_DATAREAD2:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "DATAREAD2"
    Call エラーログ記録(Me.name, "DATAREAD2", Err.Number, Err.Description, "Sql：" & Sql)
    Resume Exit_DATAREAD2
End Function

Private Function DATAWRITE1() As Boolean
On Error GoTo Err_DATAWRITE1
    
    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim Sql As String
    Dim 行番号 As Long
    
    If 入力チェック() = False Then
        Exit Function
    End If
    
    DATAWRITE1 = False
    DoCmd.Hourglass True
    
    Set MyDB = CurrentDb()
    
    Set WKDB = MyDB.OpenRecordset("SELECT * FROM W_出荷・移動データ入力_移動 WHERE SD17001 <> 0 AND SD17001 IS NOT NULL ORDER BY SD17003")
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst2 = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    cnn.BeginTrans
    
    ' 新規の場合は伝票番号を再取得する
    If Me.MODE = "新規" Then
        Call 伝票番号取得(True)
        If Nz(Me.伝票番号, "") = "" Then
            MsgBox "エラーが発生しました", vbCritical, "伝票番号採番エラー"
            GoTo Exit_DATAWRITE1
        End If
    
    ' 編集の場合は一度データを削除して全てAddNewする
    ElseIf Me.MODE = "修正" Then
        Sql = ""
        Sql = Sql & " DELETE FROM KSD17IDOU "
        Sql = Sql & " WHERE SD17002 = " & SQL_MOJICHK(Me.伝票番号)
        
        Call cnn.Execute(Sql)
    
    ' 削除の場合はすべて削除
    ElseIf Me.MODE = "削除" Then
        Sql = ""
        Sql = Sql & " DELETE FROM KSD17IDOU "
        Sql = Sql & " WHERE SD17002 = " & SQL_MOJICHK(Me.伝票番号)
        
        Call cnn.Execute(Sql)
    End If
    
    Sql = ""
    Sql = Sql & " SELECT * FROM KSD17IDOU "
    Sql = Sql & " WHERE SD17002 = " & SQL_MOJICHK(Me.伝票番号)
    
    rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
    
    行番号 = 1
    Do Until WKDB.EOF
        If WKDB!削除 = False Then
            rst.AddNew
                
                rst!SD17001 = WKDB!SD17001                  ' 処理連番
                rst!SD17002 = Me.伝票番号                   ' 伝票№
                rst!SD17003 = 行番号                        ' 行№
                rst!SD17004 = Me.処理区分                   ' 移動区分
                rst!SD17005 = Me.工場元CD                   ' 入力工場
                rst!SD17006 = Format(Me.日付, "yyyymmdd")   ' 伝票日付
                rst!SD17007 = Me.工場先CD                   ' 入出庫工場
                rst!SD17008 = Format(Me.処理日, "yyyymmdd") ' 更新日
                rst!SD17009 = Nz(WKDB!SD17009, Format(Now(), "hhmmss")) ' 処理時刻
                If Nz(Me.担当者CD, "") <> "" Then
                    rst!SD17010 = Trim(Me.担当者CD)         ' 担当者コード
                End If
                rst!SD17011 = WKDB!SD17011                  ' 製品区分
                rst!SD17012 = WKDB!SD17012                  ' 製品コード
                rst!SD17013 = Nz(WKDB!SD17013, 0)           ' 箱№
                rst!SD17014 = WKDB!SD17014                  ' M/Cコード
                rst!SD17015 = WKDB!SD17015                  ' ロット日
                rst!SD17016 = WKDB!SD17016                  ' ロット№日付
                rst!SD17017 = Nz(WKDB!SD17017, 0)           ' 数量
                rst!SD17018 = Nz(WKDB!SD17018, 0) / 1000    ' 単価
                rst!SD17019 = Nz(WKDB!SD17019, 0)           ' 金額
                rst!SD17020 = Nz(WKDB!SD17020, "00")        ' POT-ID
                rst!SD17021 = Nz(WKDB!SD17021, 0)           ' 桶№
                rst!SD17022 = Nz(WKDB!SD17022, 0)           ' 出荷先コード
                
                If WKDB!登録済 = False Then
                    rst!登録日 = Now()
                Else
                    rst!登録日 = WKDB!登録日
                    rst!変更日 = Now()
                End If
                
            rst.Update
            
            行番号 = 行番号 + 1
            
            ' リード
            If Me.工場元CD.Column(2) = 1 Then
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD04KOUT "
                Sql = Sql & " WHERE SD04001 = " & Nz(WKDB!SD17001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    ' 移動出荷
                    If Me.処理区分 = 1 Then
                        rst2!SD04017 = Me.工場先CD
                        rst2!SD04030 = Format(Me.日付, "yyyymmdd")
                        
                        rst2.Update
                        
                    ' 移動入荷
                    ElseIf Me.処理区分 = 2 Then
                        rst2!SD04017 = Me.工場元CD
                        rst2!SD04030 = Format(Me.日付, "yyyymmdd")
                        
                        rst2.Update
                    End If
                End If
                
                rst2.Close
                
            ' ケース
            Else
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD05KOUT "
                Sql = Sql & " WHERE SD05001 = " & Nz(WKDB!SD17001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    ' 移動出荷
                    If Me.処理区分 = 1 Then
                        rst2!SD05016 = Me.工場先CD
                        rst2!SD05029 = Format(Me.日付, "yyyymmdd")
                        
                        rst2.Update
                        
                    ' 移動入荷
                    ElseIf Me.処理区分 = 2 Then
                        rst2!SD05016 = Me.工場元CD
                        rst2!SD05029 = Format(Me.日付, "yyyymmdd")
                        
                        rst2.Update
                    End If
                End If
                
                rst2.Close
            End If
        
        ' 登録済みの行削除(画面で新規追加された行は未処理なので考慮しない)
        ElseIf WKDB!登録済 = True Then
            ' リード
            If Me.工場元CD.Column(2) = 1 Then
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD04KOUT "
                Sql = Sql & " WHERE SD04001 = " & Nz(WKDB!SD17001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    ' 移動出荷
                    If Me.処理区分 = 1 Then
                        rst2!SD04017 = Me.工場元CD
                        rst2!SD04030 = WKDB!SD17006
                        
                        rst2.Update
                        
                    ' 移動入荷
                    ElseIf Me.処理区分 = 2 Then
                        rst2!SD04017 = Me.工場先CD
                        rst2!SD04030 = WKDB!SD17006
                        
                        rst2.Update
                    End If
                End If
                
                rst2.Close
                
            ' ケース
            Else
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD05KOUT "
                Sql = Sql & " WHERE SD05001 = " & Nz(WKDB!SD17001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    ' 移動出荷
                    If Me.処理区分 = 1 Then
                        rst2!SD05016 = Me.工場元CD
                        rst2!SD05029 = WKDB!SD17006
                        
                        rst2.Update
                        
                    ' 移動入荷
                    ElseIf Me.処理区分 = 2 Then
                        rst2!SD05016 = Me.工場先CD
                        rst2!SD05029 = WKDB!SD17006
                        
                        rst2.Update
                    End If
                End If
                
                rst2.Close
            End If
        End If
        
        WKDB.MoveNext
    Loop
    
    cnn.CommitTrans
    
    DATAWRITE1 = True
    
    Call FRMCLS

Exit_DATAWRITE1:
    Call CloseADORecordset(rst)
    Call CloseADORecordset(rst2)
    Call CloseADOConnection(cnn)
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    DoCmd.Hourglass False
    Exit Function

Err_DATAWRITE1:
    cnn.RollbackTrans
    DoCmd.Hourglass False
    MsgBox Err.Description, , "DATAWRITE1"
    Call エラーログ記録(Me.name, "DATAWRITE1", Err.Number, Err.Description, "Sql：" & Sql)
    Resume Exit_DATAWRITE1
End Function

Private Function DATAWRITE2() As Boolean
On Error GoTo Err_DATAWRITE2
    
    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim Sql As String
    Dim 行番号 As Long
    
    If 入力チェック() = False Then
        Exit Function
    End If
    
    DATAWRITE2 = False
    DoCmd.Hourglass True
    
    Set MyDB = CurrentDb()
    
    Set WKDB = MyDB.OpenRecordset("SELECT * FROM W_出荷・移動データ入力_出荷 WHERE SD18001 <> 0 AND SD18001 IS NOT NULL ORDER BY SD18006")
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst2 = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    cnn.BeginTrans
    
    ' 新規の場合は伝票番号を再取得する
    If Me.MODE = "新規" Then
        Call 伝票番号取得(True)
        If Nz(Me.伝票番号, "") = "" Then
            MsgBox "エラーが発生しました", vbCritical, "伝票番号採番エラー"
            GoTo Exit_DATAWRITE2
        End If
    
    ' 編集の場合は一度データを削除して全てAddNewする
    ElseIf Me.MODE = "修正" Then
        Sql = ""
        Sql = Sql & " DELETE FROM KSD18SYUK "
        Sql = Sql & " WHERE SD18005 = " & SQL_MOJICHK(Me.伝票番号)
        
        Call cnn.Execute(Sql)
    
    ' 削除の場合はすべて削除
    ElseIf Me.MODE = "削除" Then
        Sql = ""
        Sql = Sql & " DELETE FROM KSD18SYUK "
        Sql = Sql & " WHERE SD18005 = " & SQL_MOJICHK(Me.伝票番号)
        
        Call cnn.Execute(Sql)
    End If
    
    Sql = ""
    Sql = Sql & " SELECT * FROM KSD18SYUK "
    Sql = Sql & " WHERE SD18005 = " & SQL_MOJICHK(Me.伝票番号)
    
    rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
    
    行番号 = 1
    Do Until WKDB.EOF
        If WKDB!削除 = False Then
            rst.AddNew
                
                rst!SD18001 = WKDB!SD18001                  ' 処理連番
                rst!SD18002 = Me.工場元CD                   ' 工場コード
                rst!SD18003 = Format(Me.日付, "yyyymmdd")   ' 伝票日付
                rst!SD18004 = Me.出荷先CD                   ' 出荷先コード
                rst!SD18005 = Me.伝票番号                   ' 伝票№
                rst!SD18006 = 行番号                        ' 行№
                If Nz(Me.担当者CD, "") <> "" Then
                    rst!SD18007 = Trim(Me.担当者CD)         ' 担当者コード
                End If
                rst!SD18008 = WKDB!SD18008                  ' 製品区分
                rst!SD18009 = WKDB!SD18009                  ' 製品コード
                rst!SD18010 = Nz(WKDB!SD18010, 0)           ' 箱№
                rst!SD18011 = WKDB!SD18011                  ' M/Cコード
                rst!SD18012 = WKDB!SD18012                  ' ロット日
                rst!SD18013 = WKDB!SD18013                  ' ロット№日付
                rst!SD18014 = Nz(WKDB!SD18014, 0)           ' 数量
                rst!SD18015 = Nz(WKDB!SD18015, 0) / 1000    ' 単価
                rst!SD18016 = Nz(WKDB!SD18016, 0)           ' 金額
                rst!SD18017 = WKDB!SD18017                  ' 製造工場コード
                rst!SD18018 = Format(Me.処理日, "yyyymmdd") ' 更新日
                rst!SD18019 = Nz(WKDB!SD18019, "00")        ' POT-ID
                rst!SD18020 = Nz(WKDB!SD18020, 0)           ' 梱包明細発行区分
                rst!SD18021 = Nz(WKDB!SD18021, 0)           ' 出荷明細発行区分
                rst!SD18022 = Nz(WKDB!SD18022, 0)           ' 本社伝票発行区分
                rst!SD18023 = Nz(WKDB!SD18023, 0)           ' 桶№
                
                If WKDB!登録済 = False Then
                    rst!登録日 = Now()
                Else
                    rst!登録日 = WKDB!登録日
                    rst!変更日 = Now()
                End If
                
            rst.Update
            
            行番号 = 行番号 + 1
            
            ' リード
            If Me.工場元CD.Column(2) = 1 Then
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD04KOUT "
                Sql = Sql & " WHERE SD04001 = " & Nz(WKDB!SD18001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    rst2!SD04008 = Nz(WKDB!SD18015, 0) / 1000
                    rst2!SD04028 = Format(Me.日付, "yyyymmdd")
                    rst2!SD04029 = Me.工場元CD
                    
                    rst2.Update
                End If
                
                rst2.Close
                
            ' ケース
            Else
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD05KOUT "
                Sql = Sql & " WHERE SD05001 = " & Nz(WKDB!SD18001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    rst2!SD05008 = Nz(WKDB!SD18015, 0) / 1000
                    rst2!SD05027 = Format(Me.日付, "yyyymmdd")
                    rst2!SD05028 = Me.工場元CD
                    
                    rst2.Update
                End If
                
                rst2.Close
            End If
        
        ' 登録済みの行削除(画面で新規追加された行は未処理なので考慮しない)
        ElseIf WKDB!登録済 = True Then
            ' リード
            If Me.工場元CD.Column(2) = 1 Then
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD04KOUT "
                Sql = Sql & " WHERE SD04001 = " & Nz(WKDB!SD18001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    rst2!SD04028 = 0
                    rst2!SD04029 = 0
                    
                    rst2.Update
                End If
                
                rst2.Close
                
            ' ケース
            Else
                Sql = ""
                Sql = Sql & " SELECT * FROM KSD05KOUT "
                Sql = Sql & " WHERE SD05001 = " & Nz(WKDB!SD18001, 0)
                
                rst2.Open Sql, cnn, adOpenKeyset, adLockOptimistic
                
                If rst2.EOF = False Then
                    rst2!SD05027 = 0
                    rst2!SD05028 = 0
                    
                    rst2.Update
                End If
                
                rst2.Close
            End If
        End If
        
        WKDB.MoveNext
    Loop
    
    cnn.CommitTrans
    
    DATAWRITE2 = True
    
    Call FRMCLS

Exit_DATAWRITE2:
    Call CloseADORecordset(rst)
    Call CloseADORecordset(rst2)
    Call CloseADOConnection(cnn)
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    DoCmd.Hourglass False
    Exit Function

Err_DATAWRITE2:
    cnn.RollbackTrans
    DoCmd.Hourglass False
    MsgBox Err.Description, , "DATAWRITE2"
    Call エラーログ記録(Me.name, "DATAWRITE2", Err.Number, Err.Description, "Sql：" & Sql)
    Resume Exit_DATAWRITE2
End Function

Public Function 入力チェック() As Boolean
On Error GoTo Err_入力チェック
    
    入力チェック = False
    
    If Nz(Me.MODE, "") = "" Then
        MsgBox "MODEが不正です", vbCritical, "エラー"
        Exit Function
    End If
    
    If Nz(Me.処理区分, 0) = 0 Then
        MsgBox "処理区分を選択してください", vbCritical, "エラー"
        Me.処理区分.SetFocus
        Exit Function
    End If
    
    If Me.MODE = "新規" Then
        If Nz(Me.工場元CD, 0) = 0 Then
            MsgBox "工場元を入力してください", vbCritical, "エラー"
            Me.工場元CD.SetFocus
            Exit Function
        End If
        
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            If Nz(Me.工場先CD, 0) = 0 Then
                MsgBox "工場先を入力してください", vbCritical, "エラー"
                Me.工場先CD.SetFocus
                Exit Function
            End If
        Else
            If Nz(Me.出荷先CD, 0) = 0 Then
                MsgBox "出荷先を入力してください", vbCritical, "エラー"
                Me.出荷先CD.SetFocus
                Exit Function
            End If
        End If
        
        If Nz(Me.MODE, "") = "新規" Then
            If Nz(Me.担当者CD, "") = "" Then
                MsgBox "担当者を入力してください", vbCritical, "エラー"
                Me.担当者CD.SetFocus
                Exit Function
            End If
        End If
        
        If Nz(Me.処理日, "") = "" Then
            MsgBox "処理日を入力してください", vbCritical, "エラー"
            Me.処理日.SetFocus
            Exit Function
        End If
        
        If Nz(Me.日付, "") = "" Then
            MsgBox "日付を入力してください", vbCritical, "エラー"
            Me.日付.SetFocus
            Exit Function
        End If
    Else
        If Me.読込FLG = False Then
            Exit Function
        End If
    End If
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        If DCount("*", "W_出荷・移動データ入力_移動") = 0 Then
            MsgBox "明細が入力されていません", vbCritical, "エラー"
            Me.SUB.SetFocus
            Exit Function
        End If
        
        If DCount("*", "W_出荷・移動データ入力_移動", "SD17001 = 0 OR SD17001 IS NULL") > 0 Then
            MsgBox "処理連番が入力されていない行があります", vbCritical, "エラー"
            Me.SUB.SetFocus
            Exit Function
        End If
    Else
        If DCount("*", "W_出荷・移動データ入力_出荷") = 0 Then
            MsgBox "明細が入力されていません", vbCritical, "エラー"
            Me.SUB.SetFocus
            Exit Function
        End If
        
        If DCount("*", "W_出荷・移動データ入力_出荷", "SD18001 = 0 OR SD18001 IS NULL") > 0 Then
            MsgBox "処理連番が入力されていない行があります", vbCritical, "エラー"
            Me.SUB.SetFocus
            Exit Function
        End If
    End If
    
    入力チェック = True
    
Exit_入力チェック:
    Exit Function

Err_入力チェック:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "入力チェック"
    Call エラーログ記録(Me.name, "入力チェック", Err.Number, Err.Description)
    Resume Exit_入力チェック
End Function

Private Sub P3_Click()
On Error GoTo Err_P3_Click
    
    Call FRMCLS

Exit_P3_Click:
    Exit Sub

Err_P3_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P3_Click"
    Call エラーログ記録(Me.name, "P3_Click", Err.Number, Err.Description)
    Resume Exit_P3_Click
End Sub

Private Sub P4_Click()
On Error GoTo Err_P4_Click
    
    Select Case ActCon
        Case "工場元CD"
            Call 工場元CD_DblClick(False)
        Case "工場先CD"
            Call 工場先CD_DblClick(False)
        Case "出荷先CD"
            Call 出荷先CD_DblClick(False)
        Case "担当者CD"
            Call 担当者CD_DblClick(False)
        Case "伝票番号"
            Call 伝票番号_DblClick(False)
    End Select

Exit_P4_Click:
    Exit Sub

Err_P4_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P4_Click"
    Call エラーログ記録(Me.name, "P4_Click", Err.Number, Err.Description)
    Resume Exit_P4_Click
End Sub

Private Sub P5_Click()
On Error GoTo Err_P5_Click
    
    If Me.MODE <> "削除" Then
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            Call DATAWRITE1
        Else
            Call DATAWRITE2
        End If
    Else
        If 入力チェック() = False Then
            Exit Sub
        End If
        
        Call アクションログ記録(4, Me.name, "データ削除　伝票番号：" & Me.伝票番号)
        
        ' 画面上で操作してから削除されると完全に削除できない可能性があるので、
        ' 再度読込して削除フラグを立てて登録処理を行う。
        
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            If DATAREAD1() = False Then
                MsgBox "削除中にエラーが発生しました", vbCritical, "エラー"
                Exit Sub
            End If
            
            Call CurrentDb.Execute("UPDATE W_出荷・移動データ入力_移動 SET 削除 = TRUE")
            
            Call DATAWRITE1
        Else
            If DATAREAD2() = False Then
                MsgBox "削除中にエラーが発生しました", vbCritical, "エラー"
                Exit Sub
            End If
            
            Call CurrentDb.Execute("UPDATE W_出荷・移動データ入力_出荷 SET 削除 = TRUE")
            
            Call DATAWRITE2
        End If
        
        Call FRMCLS
    End If

Exit_P5_Click:
    Exit Sub

Err_P5_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P5_Click"
    Call エラーログ記録(Me.name, "P5_Click", Err.Number, Err.Description)
    Resume Exit_P5_Click
End Sub

Private Sub P6_Click()
On Error GoTo Err_P6_Click
    
    If Me.MODE <> "新規" Then
        Call FRMCLS
    End If

Exit_P6_Click:
    Exit Sub

Err_P6_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P6_Click"
    Call エラーログ記録(Me.name, "P6_Click", Err.Number, Err.Description)
    Resume Exit_P6_Click
End Sub

Private Sub P7_Click()
On Error GoTo Err_P7_Click
    
    If Me.MODE <> "修正" Then
        Call FRMCLS
        
        Call モード変更(Me.MODE, "修正")
        
        Call FORMFOCUS
    End If

Exit_P7_Click:
    Exit Sub

Err_P7_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P7_Click"
    Call エラーログ記録(Me.name, "P7_Click", Err.Number, Err.Description)
    Resume Exit_P7_Click
End Sub

Private Sub P8_Click()
On Error GoTo Err_P8_Click
    
    If Me.MODE <> "削除" Then
        Call FRMCLS
        
        Call モード変更(Me.MODE, "削除")
        
        Call FORMFOCUS
    End If

Exit_P8_Click:
    Exit Sub

Err_P8_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P8_Click"
    Call エラーログ記録(Me.name, "P8_Click", Err.Number, Err.Description)
    Resume Exit_P8_Click
End Sub

Private Sub P9_Click()
On Error GoTo Err_P9_Click

    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim Sql As String
    Dim num As Long
    
    Set MyDB = CurrentDb
    
    num = Nz(Me.SUB.Form.行番号, 0)
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        If DCount("*", "W_出荷・移動データ入力_移動") = 0 Then
            Exit Sub
        End If
        
        Sql = ""
        Sql = Sql & " SELECT * FROM W_出荷・移動データ入力_移動 "
        Sql = Sql & " WHERE SD17003 >= " & num
        Sql = Sql & "   AND 削除 = FALSE "
        Sql = Sql & " ORDER BY SD17003 DESC "
        
        Set WKDB = MyDB.OpenRecordset(Sql)
        
        Do Until WKDB.EOF
            WKDB.Edit
                
                WKDB!SD17003 = WKDB!SD17003 + 1
                
            WKDB.Update
            
            WKDB.MoveNext
        Loop
        
        WKDB.AddNew
            
            WKDB!SD17003 = num
            
        WKDB.Update
    Else
        If DCount("*", "W_出荷・移動データ入力_出荷") = 0 Then
            Exit Sub
        End If
        
        Sql = ""
        Sql = Sql & " SELECT * FROM W_出荷・移動データ入力_出荷 "
        Sql = Sql & " WHERE SD18006 >= " & num
        Sql = Sql & "   AND 削除 = FALSE "
        Sql = Sql & " ORDER BY SD18006 DESC "
        
        Set WKDB = MyDB.OpenRecordset(Sql)
        
        Do Until WKDB.EOF
            WKDB.Edit
                
                WKDB!SD18006 = WKDB!SD18006 + 1
                
            WKDB.Update
            
            WKDB.MoveNext
        Loop
        
        WKDB.AddNew
            
            WKDB!SD18006 = num
            
        WKDB.Update
    End If
    
    Me.SUB.SetFocus
    Me.SUB.Form.Requery
    
    DoCmd.GoToRecord , , acGoTo, num

Exit_P9_Click:
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    DoCmd.Hourglass False
    Exit Sub

Err_P9_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P9_Click"
    Call エラーログ記録(Me.name, "P9_Click", Err.Number, Err.Description)
    Resume Exit_P9_Click
End Sub

Private Sub P10_Click()
On Error GoTo Err_P10_Click

    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim Sql As String
    Dim num As Long
    
    num = Nz(Me.SUB.Form.行番号, 0)
    
    If Me.SUB.Form.NewRecord Then
        On Error Resume Next
        Me.SUB.SetFocus
        Me.SUB.Form.Undo
        DoCmd.GoToRecord , , acGoTo, num
        Exit Sub
    End If
    
    Set MyDB = CurrentDb
    
    ' 登録済みのデータの行削除の場合は、工程管理ファイルの更新の都合上ワークテーブルの削除は行わない
    If Me.SUB.Form.登録済 = True Then
        Me.SUB.Form.削除 = True
    Else
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            Call D_DEL(MyDB, "W_出荷・移動データ入力_移動", "SD17003 = " & num & " AND 削除 = FALSE")
        Else
            Call D_DEL(MyDB, "W_出荷・移動データ入力_出荷", "SD18006 = " & num & " AND 削除 = FALSE")
        End If
    End If
    
    Me.SUB.Form.Requery
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        If DCount("*", "W_出荷・移動データ入力_移動") = 0 Then
            Exit Sub
        End If
        
        Sql = ""
        Sql = Sql & " SELECT * FROM W_出荷・移動データ入力_移動 "
        Sql = Sql & " WHERE SD17003 > " & num
        Sql = Sql & "   AND 削除 = FALSE "
        Sql = Sql & " ORDER BY SD17003 "
        
        Set WKDB = MyDB.OpenRecordset(Sql)
        
        Do Until WKDB.EOF
            WKDB.Edit
                
                WKDB!SD17003 = WKDB!SD17003 - 1
                
            WKDB.Update
            
            WKDB.MoveNext
        Loop
    Else
        If DCount("*", "W_出荷・移動データ入力_出荷") = 0 Then
            Exit Sub
        End If
        
        Sql = ""
        Sql = Sql & " SELECT * FROM W_出荷・移動データ入力_出荷 "
        Sql = Sql & " WHERE SD18006 > " & num
        Sql = Sql & "   AND 削除 = FALSE "
        Sql = Sql & " ORDER BY SD18006 "
        
        Set WKDB = MyDB.OpenRecordset(Sql)
        
        Do Until WKDB.EOF
            WKDB.Edit
                
                WKDB!SD18006 = WKDB!SD18006 - 1
                
            WKDB.Update
            
            WKDB.MoveNext
        Loop
    End If
    
    Me.SUB.SetFocus
    Me.SUB.Form.Requery
    
    DoCmd.GoToRecord , , acGoTo, num

Exit_P10_Click:
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    DoCmd.Hourglass False
    Exit Sub

Err_P10_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P10_Click"
    Call エラーログ記録(Me.name, "P10_Click", Err.Number, Err.Description)
    Resume Exit_P10_Click
End Sub

Private Sub P12_Click()
On Error GoTo Err_P12_Click

    Call 閉じる

Exit_P12_Click:
    Exit Sub

Err_P12_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P12_Click"
    Call エラーログ記録(Me.name, "P12_Click", Err.Number, Err.Description)
    Resume Exit_P12_Click
End Sub

Private Sub SUB_Enter()
On Error Resume Next
    
    If Me.MODE = "新規" Then
        If Nz(Me.工場元CD, "") = "" Then
            MsgBox "先に工場元を入力してください", vbCritical, "確認"
            Me.工場元CD.SetFocus
            Exit Sub
        End If
        
        ' 移動
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            If Nz(Me.工場先CD, "") = "" Then
                MsgBox "先に工場先を入力してください", vbCritical, "確認"
                Me.工場先CD.SetFocus
                Exit Sub
            End If
            
            If Me.工場元CD.Column(2) <> Me.工場先CD.Column(2) Then
                MsgBox "製品区分が異なる工場同士の移動はできません", vbCritical, "確認"
                Me.工場先CD.SetFocus
                Exit Sub
            End If
        
        ' 出荷
        Else
            If Nz(Me.出荷先CD, "") = "" Then
                MsgBox "先に出荷先を入力してください", vbCritical, "確認"
                Me.出荷先CD.SetFocus
                Exit Sub
            End If
        End If
    Else
        If Me.読込FLG = False Then
            Call 伝票番号_DblClick(False)
        End If
    End If
    
    Call GETACTIVE
End Sub

Private Sub 工場元CD_DblClick(Cancel As Integer)
On Error GoTo Err_工場元CD_DblClick

    検索CD = Null
    DoCmd.OpenForm "K_工場検索", , , , , acDialog
    
    If Nz(検索CD, "") = "" Then
        Me.工場元CD.SetFocus
    Else
        Me.工場元CD = CLng(検索CD)
        Me.工場元CD.SetFocus
        Call 工場元CD_AfterUpdate
    End If
    
Exit_工場元CD_DblClick:
    Exit Sub

Err_工場元CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "工場元CD_DblClick"
    Call エラーログ記録(Me.name, "工場元CD_DblClick", Err.Number, Err.Description)
    Resume Exit_工場元CD_DblClick
End Sub

Private Sub 工場先CD_DblClick(Cancel As Integer)
On Error GoTo Err_工場先CD_DblClick

    検索CD = Null
    DoCmd.OpenForm "K_工場検索", , , , , acDialog
    
    If Nz(検索CD, "") = "" Then
        Me.工場先CD.SetFocus
    Else
        Me.工場先CD = CLng(検索CD)
        Me.工場先CD.SetFocus
    End If
    
Exit_工場先CD_DblClick:
    Exit Sub

Err_工場先CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "工場先CD_DblClick"
    Call エラーログ記録(Me.name, "工場先CD_DblClick", Err.Number, Err.Description)
    Resume Exit_工場先CD_DblClick
End Sub

Private Sub 出荷先CD_BeforeUpdate(Cancel As Integer)
On Error Resume Next
    
    If 出荷先チェック(True) = False Then
        Cancel = True
    End If
End Sub

Private Sub 出荷先CD_AfterUpdate()
On Error GoTo Exit_出荷先CD_AfterUpdate
    
    If IsNumeric(Me.出荷先CD) Then
        Me.出荷先CD = CLng(Me.出荷先CD)
    End If

Exit_出荷先CD_AfterUpdate:
    Exit Sub

Err_出荷先CD_AfterUpdate:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "出荷先CD_AfterUpdate"
    Call エラーログ記録(Me.name, "出荷先CD_AfterUpdate", Err.Number, Err.Description)
    Resume Exit_出荷先CD_AfterUpdate
End Sub

Private Sub 出荷先CD_DblClick(Cancel As Integer)
On Error GoTo Err_出荷先CD_DblClick

    検索CD = Null
    DoCmd.OpenForm "K_出荷先検索", , , , , acDialog
    
    If Nz(検索CD, "") = "" Then
        Me.出荷先CD.SetFocus
    Else
        Me.出荷先CD = CLng(検索CD)
        Me.出荷先CD.SetFocus
    End If
    
Exit_出荷先CD_DblClick:
    Exit Sub

Err_出荷先CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "出荷先CD_DblClick"
    Call エラーログ記録(Me.name, "出荷先CD_DblClick", Err.Number, Err.Description)
    Resume Exit_出荷先CD_DblClick
End Sub

Private Sub 処理区分_AfterUpdate()
On Error Resume Next
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        Me.SUB.SourceObject = "D_出荷・移動データ入力SUB_移動"
        Me.出荷先CD = Null
    Else
        Me.SUB.SourceObject = "D_出荷・移動データ入力SUB_出荷"
        Me.工場先CD = Null
    End If
    
    Call 工場元CD_AfterUpdate
    
    Call FORMFOCUS
End Sub

Private Sub 処理日_AfterUpdate()
On Error Resume Next
    
    If Nz(Me.処理日, "") <> "" Then
        Me.日付 = Me.処理日
    End If
End Sub

Private Sub 処理日_DblClick(Cancel As Integer)
On Error Resume Next

    Call CallCalendar
    
    Call 処理日_AfterUpdate
End Sub

Private Sub 日付_DblClick(Cancel As Integer)
On Error Resume Next

    Call CallCalendar
End Sub

Private Sub 担当者CD_DblClick(Cancel As Integer)
On Error GoTo Err_担当者CD_DblClick

    検索CD = Null
    DoCmd.OpenForm "K_担当者検索", , , , , acDialog
    
    If Nz(検索CD, "") = "" Then
        Me.担当者CD.SetFocus
    Else
        Me.担当者CD = 検索CD
        Me.担当者CD.SetFocus
    End If
    
Exit_担当者CD_DblClick:
    Exit Sub

Err_担当者CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "担当者CD_DblClick"
    Call エラーログ記録(Me.name, "担当者CD_DblClick", Err.Number, Err.Description)
    Resume Exit_担当者CD_DblClick
End Sub

Private Sub 伝票番号_AfterUpdate()
On Error GoTo Err_伝票番号_DblClick
    
    If Nz(Me.処理区分, 0) = 0 Then
        MsgBox "処理区分を選択してください", vbCritical, "エラー"
        Me.処理区分.SetFocus
        Exit Sub
    End If
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        Call DATAREAD1
    Else
        Call DATAREAD2
    End If
    
Exit_伝票番号_DblClick:
    Exit Sub

Err_伝票番号_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "伝票番号_DblClick"
    Call エラーログ記録(Me.name, "伝票番号_DblClick", Err.Number, Err.Description)
    Resume Exit_伝票番号_DblClick
End Sub

Private Sub 伝票番号_DblClick(Cancel As Integer)
On Error GoTo Err_伝票番号_DblClick
    
    If Nz(Me.処理区分, 0) = 0 Then
        MsgBox "処理区分を選択してください", vbCritical, "エラー"
        Me.処理区分.SetFocus
        Exit Sub
    End If
    
    If Nz(Me.工場元CD, 0) = 0 Then
        MsgBox "工場元を入力してください", vbCritical, "エラー"
        Me.工場元CD.SetFocus
        Exit Sub
    End If
    
    If Nz(Me.処理日, "") = "" Then
        MsgBox "処理日を入力してください", vbCritical, "エラー"
        Me.処理日.SetFocus
        Exit Sub
    End If
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        If Nz(Me.工場先CD, 0) = 0 Then
            MsgBox "工場先を入力してください", vbCritical, "エラー"
            Me.工場先CD.SetFocus
            Exit Sub
        End If
    Else
        
    End If

    検索CD = Null
    
    P_CODE = Me.処理区分
    P_CODE2 = Me.工場元CD
    P_CODE3 = Me.処理日
    
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        P_CODE4 = Me.工場先CD
        DoCmd.OpenForm "K_移動データ検索", , , , , acDialog
        
        If Nz(検索CD, "") = "" Then
            Me.伝票番号.SetFocus
        Else
            Me.伝票番号 = 検索CD
            Call DATAREAD1
        End If
    Else
        P_CODE4 = Null
        DoCmd.OpenForm "K_出荷データ検索", , , , , acDialog
        
        If Nz(検索CD, "") = "" Then
            Me.伝票番号.SetFocus
        Else
            Me.伝票番号 = 検索CD
            Call DATAREAD2
        End If
    End If
    
Exit_伝票番号_DblClick:
    Exit Sub

Err_伝票番号_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "伝票番号_DblClick"
    Call エラーログ記録(Me.name, "伝票番号_DblClick", Err.Number, Err.Description)
    Resume Exit_伝票番号_DblClick
End Sub

Private Sub 工場元CD_AfterUpdate()
On Error Resume Next
    
    If Nz(Me.工場元CD, "") <> "" Then
        ' リード
        If Me.工場元CD.Column(2) = 1 Then
            Me.SUB.Form.番号ラベル.Caption = "箱№"
            
            If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
                Me.SUB.Form.番号.ControlSource = "SD17013"
            Else
                Me.SUB.Form.番号.ControlSource = "SD18010"
            End If
            
        ' ケース
        Else
            Me.SUB.Form.番号ラベル.Caption = "桶№"
            
            If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
                Me.SUB.Form.番号.ControlSource = "SD17021"
            Else
                Me.SUB.Form.番号.ControlSource = "SD18023"
            End If
        End If
    End If
    
    Call 伝票番号取得
End Sub

Private Function 伝票番号取得(Optional UPDFLG As Boolean = False)
    If Nz(Me.工場元CD, "") <> "" Then
        ' リード
        If Me.工場元CD.Column(2) = 1 Then
            If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
                Me.伝票番号 = 移動出荷連番取得(1, 1, UPDFLG)
            Else
                Me.伝票番号 = 移動出荷連番取得(1, 2, UPDFLG)
            End If
            
        ' ケース
        Else
            If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
                Me.伝票番号 = 移動出荷連番取得(2, 1, UPDFLG)
            Else
                Me.伝票番号 = 移動出荷連番取得(2, 2, UPDFLG)
            End If
        End If
    Else
        Me.伝票番号 = Null
    End If
End Function

Public Function 処理連番データ取得() As Boolean
On Error GoTo Err_処理連番データ取得
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim Sql As String
    Dim subSql As String
    Dim num As Long
    
    処理連番データ取得 = False
    
    If Nz(Me.SUB.Form.処理連番, "") = "" Then
'        Me.SUB.Form.処理連番 = Null
        Me.SUB.Form.製品区分 = Null
        Me.SUB.Form.製品CD = Null
        Me.SUB.Form.製品名 = Null
        Me.SUB.Form.番号 = Null
        Me.SUB.Form.ロット番号 = Null
        Me.SUB.Form.ロット番号日付 = Null
        Me.SUB.Form.マシンコード = Null
        Me.SUB.Form.数量 = Null
        Me.SUB.Form.単価 = Null
        Me.SUB.Form.金額 = Null
        Me.SUB.Form.出荷先CD = Null
        ' 出荷
        If Me.処理区分 = 3 Then
            Me.SUB.Form.製造工場コード = Null
        End If
        GoTo Exit_処理連番データ取得
    End If
    
    num = Me.SUB.Form.行番号
    
    ' 移動
    If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
        If DCount("*", "W_出荷・移動データ入力_移動", "SD17001 = " & Nz(Me.SUB.Form.処理連番, 0) & " AND 削除 = FALSE AND SD17003 <> " & num) > 0 Then
            処理連番データ取得 = True
            MsgBox "処理連番が既に入力されています", vbCritical, "確認"
            GoTo Exit_処理連番データ取得
        End If
    
    ' 出荷
    Else
        If DCount("*", "W_出荷・移動データ入力_出荷", "SD18001 = " & Nz(Me.SUB.Form.処理連番, 0) & " AND 削除 = FALSE AND SD18006 <> " & num) > 0 Then
            処理連番データ取得 = True
            MsgBox "処理連番が既に入力されています", vbCritical, "確認"
            GoTo Exit_処理連番データ取得
        End If
    End If
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    ' リード
    If Me.工場元CD.Column(2) = 1 Then
        Sql = ""
        Sql = Sql & " SELECT SD04.*, "
        Sql = Sql & " SM05.* "
        Sql = Sql & " FROM KSD04KOUT AS SD04 "
        
        ' 移動出荷・移動入荷
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            ' 製品単価マスタ　製品区分、出荷先コード、製品コード
            Sql = Sql & " LEFT JOIN KSM05STAN AS SM05 ON (CAST(SD04.SD04010 AS INTEGER) = SM05.SM05001 AND CAST(SD04.SD04003 AS INTEGER) = SM05.SM05002 AND SD04.SD04004 = SM05.SM05003) "
        
        ' 出荷
        Else
            ' 製品単価マスタ　製品区分、製品コード　※出荷先は画面上のコードで抽出
            subSql = ""
            subSql = subSql & " SELECT * FROM KSM05STAN "
            subSql = subSql & " WHERE SM05002 = " & Nz(Me.出荷先CD, 0)
            Sql = Sql & " LEFT JOIN (" & subSql & ") AS SM05 ON (CAST(SD04.SD04010 AS INTEGER) = SM05.SM05001 AND SD04.SD04004 = SM05.SM05003) "
        End If
        
        Sql = Sql & " WHERE SD04001 = " & Nz(Me.SUB.Form.処理連番, 0)
        
        rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
        
        If rst.EOF Then
            処理連番データ取得 = True
            MsgBox "データが1件もありません。", vbOKOnly, "確認"
            GoTo Exit_処理連番データ取得
        Else
            If Nz(rst!SD04028, 0) <> 0 Then
                処理連番データ取得 = True
                MsgBox "出荷済みです", vbCritical, "確認"
                GoTo Exit_処理連番データ取得
            End If
            
            ' 出荷
            If Me.処理区分 = 1 Or Me.処理区分 = 3 Then
                If Nz(rst!SD04017, 0) <> Nz(Me.工場元CD, 0) Then
                    処理連番データ取得 = True
                    MsgBox "在庫工場が違います", vbCritical, "確認"
                    GoTo Exit_処理連番データ取得
                End If
            
            ' 入荷
            Else
                If Nz(rst!SD04017, 0) <> Nz(Me.工場先CD, 0) Then
                    処理連番データ取得 = True
                    MsgBox "在庫工場が違います", vbCritical, "確認"
                    GoTo Exit_処理連番データ取得
                End If
            End If
            
            If Nz(rst!SD04032, 0) = 3 Then
                処理連番データ取得 = True
                MsgBox "削除済みです", vbCritical, "確認"
                GoTo Exit_処理連番データ取得
            End If
            
'            Me.SUB.Form.処理連番 = rst!SD04001
            Me.SUB.Form.製品区分 = rst!SD04010
            Me.SUB.Form.製品CD = rst!SD04004
            Me.SUB.Form.製品名 = rst!SD04011
            Me.SUB.Form.番号 = rst!SD04007
            Me.SUB.Form.ロット番号 = rst!SD04006
            Me.SUB.Form.ロット番号日付 = rst!SD04012
            Me.SUB.Form.マシンコード = rst!SD04005
            Me.SUB.Form.数量 = Nz(rst!SD04009, 0)
            
            ' 単価変更日以降は変更後単価を適用
            If Nz(rst!SM05006, "") <> "" Then
                If Val(Format(Me.日付, "yyyymmdd")) >= Val(rst!SM05006) Then
                    Me.SUB.Form.単価 = Nz(rst!SM05007, 0) * 1000
                Else
                    Me.SUB.Form.単価 = Nz(rst!SM05004, 0) * 1000
                End If
            Else
                Me.SUB.Form.単価 = 0
            End If
            
            Me.SUB.Form.金額 = ROUND(Nz(rst!SD04009, 0) * Nz(Me.SUB.Form.単価, 0), 5, 0)
            Me.SUB.Form.出荷先CD = rst!SD04003
            
            ' 出荷
            If Me.処理区分 = 3 Then
                Me.SUB.Form.製造工場コード = rst!SD04002
            End If
        End If
        
    ' ケース
    Else
        Sql = ""
        Sql = Sql & " SELECT SD05.*, "
        Sql = Sql & " SM05.* "
        Sql = Sql & " FROM KSD05KOUT AS SD05 "
        
        ' 移動出荷・移動入荷
        If Me.処理区分 = 1 Or Me.処理区分 = 2 Then
            ' 製品単価マスタ　製品区分、出荷先コード、製品コード
            Sql = Sql & " LEFT JOIN KSM05STAN AS SM05 ON (CAST(SD05.SD05010 AS INTEGER) = SM05.SM05001 AND CAST(SD05.SD05003 AS INTEGER) = SM05.SM05002 AND SD05.SD05005 = SM05.SM05003) "
        
        ' 出荷
        Else
            ' 製品単価マスタ　製品区分、製品コード　※出荷先は画面上のコードで抽出
            subSql = ""
            subSql = subSql & " SELECT * FROM KSM05STAN "
            subSql = subSql & " WHERE SM05002 = " & Nz(Me.出荷先CD, 0)
            Sql = Sql & " LEFT JOIN (" & subSql & ") AS SM05 ON (CAST(SD05.SD05010 AS INTEGER) = SM05.SM05001 AND SD05.SD05005 = SM05.SM05003) "
        End If
        
        Sql = Sql & " WHERE SD05001 = " & Nz(Me.SUB.Form.処理連番, 0)
        
        rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
        
        If rst.EOF Then
            処理連番データ取得 = True
            MsgBox "データが1件もありません。", vbOKOnly, "確認"
            GoTo Exit_処理連番データ取得
        Else
            If Nz(rst!SD05027, 0) <> 0 Then
                処理連番データ取得 = True
                MsgBox "出荷済みです", vbCritical, "確認"
                GoTo Exit_処理連番データ取得
            End If
            
            ' 出荷
            If Me.処理区分 = 1 Or Me.処理区分 = 3 Then
                If Nz(rst!SD05016, 0) <> Nz(Me.工場元CD, 0) Then
                    処理連番データ取得 = True
                    MsgBox "在庫工場が違います", vbCritical, "確認"
                    GoTo Exit_処理連番データ取得
                End If
            
            ' 入荷
            Else
                If Nz(rst!SD05016, 0) <> Nz(Me.工場先CD, 0) Then
                    処理連番データ取得 = True
                    MsgBox "在庫工場が違います", vbCritical, "確認"
                    GoTo Exit_処理連番データ取得
                End If
            End If
            
            If Nz(rst!SD05031, 0) = 3 Then
                処理連番データ取得 = True
                MsgBox "削除済みです", vbCritical, "確認"
                GoTo Exit_処理連番データ取得
            End If
            
'            Me.SUB.Form.処理連番 = rst!SD05001
            Me.SUB.Form.製品区分 = rst!SD05010
            Me.SUB.Form.製品CD = rst!SD05005
            Me.SUB.Form.製品名 = rst!SD05011
            Me.SUB.Form.番号 = rst!SD05007
            Me.SUB.Form.ロット番号 = rst!SD05006
            Me.SUB.Form.ロット番号日付 = rst!SD05012
            Me.SUB.Form.マシンコード = rst!SD05004
            Me.SUB.Form.数量 = Nz(rst!SD05009, 0)
            
            ' 単価変更日以降は変更後単価を適用
            If Nz(rst!SM05006, "") <> "" Then
                If Val(Format(Me.日付, "yyyymmdd")) >= Val(rst!SM05006) Then
                    Me.SUB.Form.単価 = Nz(rst!SM05007, 0) * 1000
                Else
                    Me.SUB.Form.単価 = Nz(rst!SM05004, 0) * 1000
                End If
            Else
                Me.SUB.Form.単価 = 0
            End If
            
            Me.SUB.Form.金額 = ROUND(Nz(rst!SD05009, 0) * Nz(Me.SUB.Form.単価, 0), 5, 0)
            Me.SUB.Form.出荷先CD = rst!SD05003
            
            ' 出荷
            If Me.処理区分 = 3 Then
                Me.SUB.Form.製造工場コード = rst!SD05002
            End If
        End If
    End If
    
Exit_処理連番データ取得:
    Call CloseADORecordset(rst)
    Call CloseADOConnection(cnn)
    Exit Function

Err_処理連番データ取得:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "処理連番データ取得"
    Call エラーログ記録(Me.name, "処理連番データ取得", Err.Number, Err.Description, "Sql：" & Sql)
    Resume Exit_処理連番データ取得
End Function

Private Sub 登録へ_Enter()
On Error Resume Next

    Me.P5.SetFocus
End Sub

' 2024/12/17
Private Sub 印刷済み1_Click()
On Error GoTo Err_印刷済み1_Click

    If Me.処理区分 <> 3 Then
        Exit Sub
    End If
    
'    If MsgBox("表示された伝票の梱包明細を全て印刷済みにしてもよろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認") = vbNo Then
'        Exit Sub
'    End If
    
    Call CurrentDb().Execute("UPDATE W_出荷・移動データ入力_出荷 SET SD18020 = 1")
    
    Me.SUB.Form.Requery
    
    Call 未発行更新
    
'    MsgBox "印刷済み処理が完了しました" & vbNewLine & "登録にて反映されます", vbOKOnly, "完了"
    
Exit_印刷済み1_Click:
    DoCmd.Hourglass False
    Exit Sub

Err_印刷済み1_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "印刷済み1_Click"
    Call エラーログ記録(Me.name, "印刷済み1_Click", Err.Number, Err.Description)
    Resume Exit_印刷済み1_Click
End Sub

' 2024/12/17
Private Sub 印刷済み2_Click()
On Error GoTo Err_印刷済み2_Click

    If Me.処理区分 <> 3 Then
        Exit Sub
    End If
    
'    If MsgBox("表示された伝票の出荷明細を全て印刷済みにしてもよろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認") = vbNo Then
'        Exit Sub
'    End If
    
    Call CurrentDb().Execute("UPDATE W_出荷・移動データ入力_出荷 SET SD18021 = 1")
    
    Me.SUB.Form.Requery
    
    Call 未発行更新
    
'    MsgBox "印刷済み処理が完了しました" & vbNewLine & "登録にて反映されます", vbOKOnly, "完了"
    
Exit_印刷済み2_Click:
    DoCmd.Hourglass False
    Exit Sub

Err_印刷済み2_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "印刷済み2_Click"
    Call エラーログ記録(Me.name, "印刷済み2_Click", Err.Number, Err.Description)
    Resume Exit_印刷済み2_Click
End Sub

' 2024/12/17
Private Sub 印刷済み3_Click()
On Error GoTo Err_印刷済み3_Click

    If Me.処理区分 <> 3 Then
        Exit Sub
    End If
    
'    If MsgBox("表示された伝票の本社伝票を全て印刷済みにしてもよろしいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認") = vbNo Then
'        Exit Sub
'    End If
    
    Call CurrentDb().Execute("UPDATE W_出荷・移動データ入力_出荷 SET SD18022 = 1")
    
    Me.SUB.Form.Requery
    
    Call 未発行更新
    
'    MsgBox "印刷済み処理が完了しました" & vbNewLine & "登録にて反映されます", vbOKOnly, "完了"
    
Exit_印刷済み3_Click:
    DoCmd.Hourglass False
    Exit Sub

Err_印刷済み3_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "印刷済み3_Click"
    Call エラーログ記録(Me.name, "印刷済み3_Click", Err.Number, Err.Description)
    Resume Exit_印刷済み3_Click
End Sub

Public Function 未発行更新()
On Error Resume Next
    
    Me.MSG1.Requery
    Me.MSG2.Requery
    Me.MSG3.Requery
End Function
