VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_P_検査成績表_リード"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const maxKoumoku As Long = 13

' 警告表示する件数
Private Const keikokuCnt As Long = 50

Private SM17 As SM17_S
Private SM18 As SM18_S
Private SD06 As SD06_S

Private WSM17(0 To maxKoumoku) As SM17_S
Private WSM18(0 To maxKoumoku) As SM18_S
Private WSD06(0 To maxKoumoku) As SD06_S

Private sSD06002 As String  ' 処理日
Private sSD06009 As String  ' M/Cコード
Private sSD06012 As String  ' 製品コード
Private sSD06027 As String  ' 通常追加区分

' 空欄項目に表示する文字
Private Const nullStr As String = "―"
Private Const nullStr2 As String = "＿"

Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Form_Open
    
    PreForm = AFNAME
    AFNAME = Me.name

    Application.SetOption KEYMACRO, "入力キー"
    Me.Visible = True

    Call 初期設定(Me.name)

    Call FRMCLS

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description, , "フォームオープン"
    Resume Exit_Form_Open
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
    
    Me.S工場CD = FACTORY
    Me.E工場CD = FACTORY
    Me.S日付 = Date
    Me.E日付 = Date
    
    Call D_DEL(CurrentDb(), "W_検査成績表_リード")
    
    Me.SUB.Form.Requery
End Function

Function PRT(PFLG As Integer)
On Error GoTo Err_PRT

    Dim MyDB As DAO.Database
    Dim WKDB As DAO.Recordset
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim rst3 As ADODB.Recordset
    Dim Sql As String
    Dim subSql As String
    Dim Crit As String
    Dim SCODE As Variant
    Dim ECODE As Variant
    Dim SDATE As Variant
    Dim EDATE As Variant
    Dim SCODE2 As Variant
    Dim ECODE2 As Variant
    Dim SCODE3 As Variant
    Dim ECODE3 As Variant
    Dim cnt As Long
    Dim tmpMAX As Double
    Dim tmpMIN As Double
    Dim tmpSum As Double
    Dim i As Long
    Dim arr() As Double
    Dim tmp規格上限 As Double
    Dim tmp規格下限 As Double
    Dim tmpCpk1 As Double
    Dim tmpCpk2 As Double
    Dim tmp測定値 As Double
    Dim tmpKeta As Long
    Dim tmpHasuu As Long
    Dim tmp検査日 As Long
    Dim DTCNT As Long
    Dim tmp項目順 As Long
    Dim HEADSW As Boolean
    Dim bSkipFlg As Boolean
    Dim k As Long
    
    #If Win64 Then
        Dim maxCnt As LongLong
    #Else
        Dim maxCnt As Long
    #End If

    DoCmd.Hourglass True

    Set MyDB = CurrentDb
    
    Call D_DEL(MyDB, "W_検査成績表_リード")
    Me.SUB.Form.Requery
    
    Set WKDB = MyDB.OpenRecordset("W_検査成績表_リード")

    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst2 = New ADODB.Recordset
    Set rst3 = New ADODB.Recordset

    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    Call ShowProgressBar(0, "出力処理", "出力前処理中", True)
    
    Crit = "" & vbNewLine
    
    ' 工場コード
    SCODE = Nz(Me.S工場CD, 0)
    ECODE = Nz(Me.E工場CD, 99)
    Crit = Crit & " SD04002 BETWEEN " & SCODE & " AND " & ECODE & vbNewLine
    
    ' BOXラベル発行日
    SDATE = Nz(Me.S日付, "1900/01/01")
    EDATE = Nz(Me.E日付, "9999/12/31")
    Crit = Crit & "   AND SD04037 BETWEEN " & Format(SDATE, "yyyymmdd") & " AND " & Format(EDATE, "yyyymmdd") & vbNewLine
    
    ' 製品
    SCODE2 = Nz(StrConv(Me.S製品CD, vbUpperCase), "00000")
    ECODE2 = Nz(StrConv(Me.E製品CD, vbUpperCase), "zzzzz")
    Crit = Crit & "   AND SD04004 BETWEEN " & SQL_MOJICHK(SCODE2) & " AND " & SQL_MOJICHK(ECODE2) & vbNewLine
    
    ' M/C
    SCODE3 = Nz(Me.SMCCD, "0")
    ECODE3 = Nz(Me.EMCCD, "zzzz")
    Crit = Crit & "   AND SD04005 BETWEEN " & SQL_MOJICHK(SCODE3) & " AND " & SQL_MOJICHK(ECODE3) & vbNewLine
    
    ' 新規発行
    If Nz(Me.FL_発行区分, 0) = 0 Then
        Crit = Crit & "   AND SD04016 = 0 " & vbNewLine
    
    ' 再発行
    Else
        Crit = Crit & "   AND SD04016 <> 0 " & vbNewLine
    End If
    
    ' 削除は除く
    Crit = Crit & "   AND SD04032 <> 3 " & vbNewLine
    
    Sql = "" & vbNewLine
    Sql = Sql & " SELECT SD04.*, " & vbNewLine
    Sql = Sql & " SM01.*, " & vbNewLine
    Sql = Sql & " SM02.*, " & vbNewLine
    Sql = Sql & " SM03.*, " & vbNewLine
    Sql = Sql & " SM04.*, " & vbNewLine
    Sql = Sql & " SM06.* " & vbNewLine
    
    ' リード工程管理ファイル
    subSql = "" & vbNewLine
    subSql = subSql & " SELECT * FROM KSD04KOUT " & vbNewLine
    subSql = subSql & " WHERE " & Crit & vbNewLine
    
    Sql = Sql & " FROM (" & subSql & ") AS SD04 " & vbNewLine
    
    ' 工場マスタ　工場コード
    Sql = Sql & " LEFT JOIN KSM01KOJY AS SM01 ON (SD04.SD04002 = SM01.SM01001) " & vbNewLine
    
    ' 得意先マスタ　得意先コード(出荷先コードの左2桁)　※出荷先コードは4桁固定ではないため4桁に合わせてから切り取る
    Sql = Sql & " LEFT JOIN KSM02TOKU AS SM02 ON (CAST(LEFT(FORMAT(SD04.SD04003, '0000'), 2) AS INTEGER) = SM02.SM02001) " & vbNewLine
    
    ' 出荷先マスタ　出荷先コード
    Sql = Sql & " LEFT JOIN KSM03SYUK AS SM03 ON (SD04.SD04003 = SM03.SM03001) " & vbNewLine
    
    ' 製品マスタ　製品区分、製品コード
    Sql = Sql & " LEFT JOIN KSM04SEIH AS SM04 ON (SD04.SD04010 = SM04.SM04001 AND SD04.SD04004 = SM04.SM04002) " & vbNewLine
    
    ' 検査成績表備考マスタ　製品区分、製品コード
    Sql = Sql & " LEFT JOIN KSM06BIKO AS SM06 ON (SD04.SD04010 = SM06.SM06001 AND SD04.SD04004 = SM06.SM06002) " & vbNewLine
    
    ' 工場コード, 登録日, 製品コード, M/Cコード
    Sql = Sql & " ORDER BY SD04.SD04002, SD04.SD04037, SD04.SD04004, SD04.SD04005 " & vbNewLine
    
    rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
    
    If rst.EOF Then
        Call ShowProgressBar(, , , False)
        MsgBox "データが1件もありません。", vbOKOnly, "確認"
        GoTo Exit_PRT
    End If
    
    DTCNT = 1
    maxCnt = rst.RecordCount
    
    If maxCnt > keikokuCnt Then
        If MsgBox(keikokuCnt & "件を超えるデータが選択されましたが、処理を継続してもよろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, "確認") = vbNo Then
            GoTo Exit_PRT
        End If
    End If
    
    Do Until rst.EOF
        If ProgressBarCancelFlag = True Then
            DoCmd.Hourglass False
            Call ShowProgressBar(, , , False)
            MsgBox "処理がキャンセルされました。", vbCritical, "キャンセル"
            GoTo Exit_PRT
        End If
        
        Call ShowProgressBar((DTCNT / maxCnt) * 100, "出力処理 (" & DTCNT & " / " & maxCnt & ")", "出力中", True, True)
        
        ' 初期化
        For i = 1 To maxKoumoku
            Call CLR_SM17(WSM17(i))
            Call CLR_SM18(WSM18(i))
            Call CLR_SD06(WSD06(i))
        Next
        
        HEADSW = False
        bSkipFlg = False
        
        ' 再発行はチェックなし
        If Nz(Me.FL_発行区分, 0) = 1 Then
            GoTo MRED_050
        End If
        
        ' D04と同じM/C、製品で検査日<=ロット№日付で最大のもの
        Sql = "" & vbNewLine
        Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
        Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0) & vbNewLine        ' 工場コード
        Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(rst!SD04005) & vbNewLine  ' M/Cコード
        Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(rst!SD04004) & vbNewLine  ' 製品コード
        Sql = Sql & "   AND SD06002 <= " & Nz(rst!SD04012, 0) & vbNewLine       ' 処理日
        Sql = Sql & " ORDER BY SD06001, SD06002 DESC, SD06012, SD06009, SD06014, SD06010, SD06027 DESC " & vbNewLine ' 工場コード, 処理日降順, 製品コード, M/Cコード, 検査項目№, 項目順, 通常追加区分降順
        
        rst2.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
        
        If rst2.EOF Then
            HEADSW = True
            sSD06002 = ""
            sSD06009 = ""
            sSD06012 = ""
            rst2.Close
            GoTo MRED_050
        Else
            Call SET_SD06(SD06, rst2)
            
            sSD06002 = Nz(rst2!SD06002, 0)  ' 処理日
            sSD06009 = Nz(rst2!SD06009, "") ' M/Cコード
            sSD06012 = Nz(rst2!SD06012, "") ' 製品コード
            sSD06027 = Nz(rst2!sd06027, 0)  ' 通常追加区分
            
            rst2.Close
            
            ' 検査日 = ロット№日付
            If CLng(SD06.S002) = Nz(rst!SD04012, 0) Then
                GoTo MRED_040
            Else
                GoTo MRED_030
            End If
        End If
        
MRED_022:
        DoEvents
        
        Sql = "" & vbNewLine
        Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
        Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0) & vbNewLine        ' 工場コード
        Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(rst!SD04005) & vbNewLine  ' M/Cコード
        Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(rst!SD04004) & vbNewLine  ' 製品コード
        Sql = Sql & "   AND SD06002 < " & Nz(sSD06002, 0) & vbNewLine           ' 処理日
        Sql = Sql & " ORDER BY SD06001, SD06002 DESC, SD06012, SD06009, SD06014, SD06010, SD06027 DESC " & vbNewLine ' 工場コード, 処理日降順, 製品コード, M/Cコード, 検査項目№, 項目順, 通常追加区分降順
        
        rst2.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
        
        If rst2.EOF Then
            HEADSW = True
            sSD06002 = ""
            sSD06009 = ""
            sSD06012 = ""
            rst2.Close
            GoTo MRED_050
        Else
            Call SET_SD06(SD06, rst2)
            
            sSD06002 = Nz(rst2!SD06002, 0)  ' 処理日
            sSD06009 = Nz(rst2!SD06009, "") ' M/Cコード
            sSD06012 = Nz(rst2!SD06012, "") ' 製品コード
            sSD06027 = Nz(rst2!sd06027, 0)  ' 通常追加区分
            
            rst2.Close
        End If
        
MRED_030:
        DoEvents
        
        ' D04と同じM/C、違う製品で検査日<=ロット№日付で最大のもの
        Sql = "" & vbNewLine
        Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
        Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0) & vbNewLine        ' 工場コード
        Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(rst!SD04005) & vbNewLine  ' M/Cコード
        Sql = Sql & "   AND SD06012 <> " & SQL_MOJICHK(rst!SD04004) & vbNewLine ' 製品コード
        Sql = Sql & "   AND SD06002 <= " & Nz(rst!SD04012, 0) & vbNewLine       ' 処理日
        Sql = Sql & " ORDER BY SD06001, SD06002 DESC, SD06012, SD06009, SD06014, SD06010, SD06027 DESC " & vbNewLine ' 工場コード, 処理日降順, 製品コード, M/Cコード, 検査項目№, 項目順, 通常追加区分降順
        
        rst2.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
        
        If rst2.EOF Then
            rst2.Close
        Else
            Call SET_SD06(SD06, rst2)
            
            rst2.Close
            
            ' 処理日
            If SD06.S002 > sSD06002 Then
                HEADSW = True
                GoTo MRED_050
            End If
        End If
        
MRED_040:
        DoEvents
        
        ' 対象データが計測済みか調べる
        Sql = "" & vbNewLine
        Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
        Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0) & vbNewLine        ' 工場コード
        Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(sSD06009) & vbNewLine     ' M/Cコード
        Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(sSD06012) & vbNewLine     ' 製品コード
        Sql = Sql & "   AND SD06002 = " & Nz(sSD06002, 0) & vbNewLine           ' 処理日
        Sql = Sql & "   AND SD06027 = " & Nz(sSD06027, 0) & vbNewLine           ' 通常追加区分
        Sql = Sql & "   AND SD06025 = 0 " & vbNewLine                           ' 計測済み区分
        Sql = Sql & " ORDER BY SD06001, SD06002, SD06012, SD06009, SD06014, SD06010 " & vbNewLine ' 工場コード, 処理日, 製品コード, M/Cコード, 検査項目№, 項目順
        
        rst2.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
        
        If rst2.EOF Then
            rst2.Close
        
        ' 未計測なら計測済みのデータになるまで探し続ける
        Else
            Call SET_SD06(SD06, rst2)
            
            rst2.Close
            
            GoTo MRED_022
        End If
        
MRED_050:
        DoEvents
        
        ' 検査値マスタの取得
        Sql = "" & vbNewLine
        Sql = Sql & " SELECT * FROM KSM18KSCH " & vbNewLine
        Sql = Sql & " WHERE SM18001 = " & SQL_MOJICHK(Me.製品区分) & vbNewLine  ' 製品区分
        Sql = Sql & "   AND SM18002 = " & SQL_MOJICHK(rst!SD04004) & vbNewLine  ' 製品コード
        Sql = Sql & " ORDER BY SM18003 " & vbNewLine    ' 項目順
        
        rst2.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
        
        k = 0
        Do Until rst2.EOF
            DoEvents
            
            Call SET_SM18(SM18, rst2)
            Call CLR_SM17(SM17)
            Call CLR_SD06(SD06)
            
            ' 検査項目名が入っていて、検査値が0以外であれば表示する
            If Trim(SM18.S004) <> "" And SM18.D005 <> 0 Then
                If HEADSW = False Then
                    ' 新規発行
                    If Nz(Me.FL_発行区分, 0) = 0 Then
                        Sql = "" & vbNewLine
                        Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
                        Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0) & vbNewLine        ' 工場コード
                        Sql = Sql & "   AND SD06002 = " & Nz(sSD06002, 0) & vbNewLine           ' 処理日
                        Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(rst!SD04004) & vbNewLine  ' 製品コード
                        Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(rst!SD04005) & vbNewLine  ' M/Cコード
                        Sql = Sql & "   AND SD06011 = " & Nz(Me.製品区分, 0) & vbNewLine        ' 製品区分
                        Sql = Sql & "   AND SD06014 = " & SQL_MOJICHK(SM18.L003) & vbNewLine    ' 検査項目№
                        
                        ' UPDATE 2025/01/20 SUZUKI 通常と追加が混在する場合に通常が表示されてしまう。sSD06027を出しているのでイコールでいいはず
'                        Sql = Sql & "   AND SD06027 <= " & Nz(sSD06027, 0) & vbNewLine          ' 通常追加区分
                        Sql = Sql & "   AND SD06027 = " & Nz(sSD06027, 0) & vbNewLine          ' 通常追加区分

                        Sql = Sql & " ORDER BY SD06001, SD06002, SD06012, SD06009, SD06014, SD06010 " & vbNewLine ' 工場コード, 処理日, 製品コード, M/Cコード, 検査項目№, 項目順
                        
                        rst3.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
                        
                        If rst3.EOF = False Then
                            Call SET_SD06(SD06, rst3)
                        End If
                        
                        rst3.Close
                        
                    ' 再発行
                    Else
                        Sql = "" & vbNewLine
                        Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
                        Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0) & vbNewLine        ' 工場コード
                        Sql = Sql & "   AND SD06002 = " & Nz(rst!SD04016, 0) & vbNewLine        ' 処理日
                        Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(rst!SD04004) & vbNewLine  ' 製品コード
                        Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(rst!SD04005) & vbNewLine  ' M/Cコード
                        Sql = Sql & "   AND SD06011 = " & Nz(Me.製品区分, 0) & vbNewLine        ' 製品区分
                        Sql = Sql & "   AND SD06014 = " & SQL_MOJICHK(SM18.L003) & vbNewLine    ' 検査項目№
                        Sql = Sql & " ORDER BY SD06001, SD06002, SD06012, SD06009, SD06014, SD06010, SD06027 DESC " & vbNewLine ' 工場コード, 処理日, 製品コード, M/Cコード, 検査項目№, 項目順, 通常追加区分降順
                        
                        rst3.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
                        
                        If rst3.EOF = False Then
                            Call SET_SD06(SD06, rst3)
                        End If
                        
                        rst3.Close
                    End If
                    
                    ' 検査項目マスタの取得
                    Sql = "" & vbNewLine
                    Sql = Sql & " SELECT * FROM KSM17KSKM " & vbNewLine
                    Sql = Sql & " WHERE SM17001 = " & SQL_MOJICHK(Me.製品区分) & vbNewLine                      ' 製品区分
                    Sql = Sql & "   AND SM17002 = " & CInt(Left(Format(rst!SD04003, "0000"), 2)) & vbNewLine    ' 得意先コード
                    Sql = Sql & "   AND SM17003 = " & SQL_MOJICHK(SM18.L003) & vbNewLine                        ' 項目順
                    
                    rst3.Open Sql, cnn, adOpenForwardOnly, adLockReadOnly
                    
                    If rst3.EOF = False Then
                        Call SET_SM17(SM17, rst3)
                    End If
                    
                    rst3.Close
                    
                    ' 検査結果ファイルが存在しない場合
                    If Nz(SD06.S001, "") = "" Then
                        ' 測定区分が文字項目の場合
                        If SM18.S008 = "0" Then
                            If k >= maxKoumoku Then
                                GoTo SM18_END
                            End If
                            
                            k = k + 1
                            WSM17(k) = SM17
                            WSM18(k) = SM18
                        End If
                    Else
                        If k >= maxKoumoku Then
                            GoTo SM18_END
                        End If
                        
                        k = k + 1
                        WSM17(k) = SM17
                        WSM18(k) = SM18
                        WSD06(k) = SD06
                    End If
                Else
                    If k >= maxKoumoku Then
                        GoTo SM18_END
                    End If
                    k = k + 1
                    WSM18(k) = SM18
                End If
            End If
            
            rst2.MoveNext
        Loop

SM18_END:
        rst2.Close
        
        ' 表示する検査項目が1つもない場合
        If k = 0 Then
            GoTo NextLoop
        End If
        
        WKDB.AddNew
            
            ' UPDATE 2025/02/17 SUZUKI 未測定の項目がある場合は印刷しない
            WKDB!印刷FLG = Not HEADSW
            
            WKDB!印字区分 = rst!SM03012
            WKDB!処理連番 = rst!SD04001
            WKDB!工場コード = rst!SD04002
            WKDB!出荷先コード = rst!SD04003
            WKDB!得意先コード = Left(Format(rst!SD04003, "0000"), 2)
            WKDB!得意先名 = RTrim(Nz(rst!SM02002, "")) & "　殿"
            WKDB!出荷先名1 = RTrim(rst!SM02010)
            WKDB!出荷先名2 = RTrim(rst!SM02011)
            WKDB!出荷先名3 = RTrim(rst!SM02012)
            WKDB!出荷先名4 = RTrim(rst!SM02013)
            WKDB!成績表タイトル = rst!SM02014
            WKDB!製品コード = rst!SD04004
            WKDB!マシンコード = rst!SD04005
            WKDB!品名 = RTrim(rst!SD04011)
            WKDB!付加記号 = RTrim(rst!SM04022)
            WKDB!数量 = rst!SD04009 * 1000
            If Nz(rst!SD04039, 0) = 0 Then
                ' UPDATE 2024/12/20 年の下1桁が0の時に桁数がずれるため
'                WKDB!ロット№ = Right(rst!SD04005, 3) & " " & rst!SD04006
                WKDB!ロット№ = Right(rst!SD04005, 3) & " " & Format(rst!SD04006, "00000")
            Else
                ' UPDATE 2024/12/20 年の下1桁が0の時に桁数がずれるため
'                WKDB!ロット№ = Right(rst!SD04005, 3) & " " & rst!SD04006 & "-" & Format(rst!SD04039, "00")
                WKDB!ロット№ = Right(rst!SD04005, 3) & " " & Format(rst!SD04006, "00000") & "-" & Format(rst!SD04039, "00")
            End If
            WKDB!検査年月日 = DateSerial(Left(rst!SD04037, 4), Mid(rst!SD04037, 5, 2), Right(rst!SD04037, 2))
            
            WKDB!箱番号 = rst!SD04007
            WKDB!備考 = rst!SM06015
            
            ' UPDATE 2025/03/05 SUZUKI 合格印追加
            WKDB!合格印印字区分 = Nz(rst!SM03029, 0)
            
            ' UPDATE 2025/03/05 SUZUKI 確認者印字
            WKDB!確認者印字区分 = Nz(rst!SM03031, 0)
            
            ' 項目
            For cnt = 1 To maxKoumoku
                ReDim arr(0)
                
                ' 項目名
                WKDB("項目" & cnt) = Trim(WSM18(cnt).S004)
                
                If Trim(WSM18(cnt).S004) <> "" Then
                    ' 規格（上段）
                    WKDB("規格1-" & cnt) = Trim(WSM18(cnt).S006)
                    
                    ' 規格（下段）
                    WKDB("規格2-" & cnt) = Trim(WSM18(cnt).S007)
                    
                    ' 対外公差＋, 対外公差－
                    If WSM18(cnt).D009 = 0 And WSM18(cnt).D010 = 0 Then
                    Else
                        ' 公差＋
                        If WSM18(cnt).D009 >= 0 Then
                            WKDB("公差プラス" & cnt) = "+" & Abs(WSM18(cnt).D009)
                        Else
                            WKDB("公差プラス" & cnt) = "-" & Abs(WSM18(cnt).D009)
                        End If
                        
                        ' 公差－
                        If WSM18(cnt).D010 > 0 Then
                            WKDB("公差マイナス" & cnt) = "+" & Abs(WSM18(cnt).D010)
                        Else
                            WKDB("公差マイナス" & cnt) = "-" & Abs(WSM18(cnt).D010)
                        End If
                    End If
                    
                    ' 結果ファイルがない場合は表示しない
                    If WSD06(cnt).S001 <> "" Then
                        ' 測定値(計測回数分)
                        For i = 1 To WSD06(cnt).L019
                            Select Case WSM18(cnt).S008
                                ' 数字項目
                                Case "1"
                                    Select Case i
                                        Case 1
                                            tmp測定値 = WSD06(cnt).D020
                                        Case 2
                                            tmp測定値 = WSD06(cnt).D021
                                        Case 3
                                            tmp測定値 = WSD06(cnt).D022
                                        Case 4
                                            tmp測定値 = WSD06(cnt).D023
                                        Case 5
                                            tmp測定値 = WSD06(cnt).D024
                                    End Select
                                    
                                    tmp測定値 = tmp測定値 + WSM18(cnt).D014
                                    
                                    tmpKeta = WSM17(cnt).L007
                                    tmpHasuu = WSM17(cnt).S008
                                    
                                    ' 単位変換
                                    Select Case WSM17(cnt).S012
                                        ' 無し
                                        Case "0"
                                            WKDB("K" & cnt & "-" & i) = tmp測定値
                                            
                                        ' N
                                        Case "1"
                                            WKDB("K" & cnt & "-" & i) = tmp測定値 * 9.8
                                            tmpKeta = 1
                                            tmpHasuu = 0
                                            
                                        ' Pa
                                        Case "2"
                                            WKDB("K" & cnt & "-" & i) = tmp測定値 * 9.8
                                            tmpKeta = 0
                                            tmpHasuu = 0
                                            
                                        ' Mpa
                                        Case "3"
                                            WKDB("K" & cnt & "-" & i) = tmp測定値 * 0.098
                                            tmpKeta = 2
                                            tmpHasuu = 0
                                    End Select
                                    
                                    ' バリの場合は1000で割る
                                    If WKDB("項目" & cnt) = "バリ" Then
                                        WKDB("K" & cnt & "-" & i) = WKDB("K" & cnt & "-" & i) / 1000
                                    End If
                                    
                                    ' 端数処理
                                    Select Case tmpHasuu
                                        ' 切り捨て
                                        Case 0
                                            WKDB("K" & cnt & "-" & i) = ROUND(WKDB("K" & cnt & "-" & i), 5 + tmpKeta, 2)
                                        
                                        ' 切り上げ
                                        Case 1
                                            WKDB("K" & cnt & "-" & i) = ROUND(WKDB("K" & cnt & "-" & i), 5 + tmpKeta, 1)
                                            
                                        ' 四捨五入
                                        Case 2
                                            WKDB("K" & cnt & "-" & i) = ROUND(WKDB("K" & cnt & "-" & i), 5 + tmpKeta, 0)
                                    End Select
                                    
                                    If i = 1 Then
                                        ReDim arr(0)
                                    Else
                                        ReDim Preserve arr(UBound(arr) + 1)
                                    End If
                                    arr(UBound(arr)) = WKDB("K" & cnt & "-" & i)
                                    
                                    ' 表示桁設定
                                    If tmpKeta > 0 Then
                                        WKDB("K" & cnt & "-" & i) = Format(WKDB("K" & cnt & "-" & i), "0." & String(tmpKeta, "0"))
                                    End If
                                
                                ' 文字項目
                                Case "0"
                                    
                            End Select
                        Next
                        
                        Select Case WSM18(cnt).S008
                            ' 数字項目
                            Case "1"
                                ' 最大値
                                WKDB("MAX" & cnt) = getMaxValue(arr)
                                If tmpKeta > 0 Then
                                    WKDB("MAX" & cnt) = Format(WKDB("MAX" & cnt), "0." & String(tmpKeta, "0"))
                                End If
                                
                                ' 最小値
                                WKDB("MIN" & cnt) = getMinValue(arr)
                                If tmpKeta > 0 Then
                                    WKDB("MIN" & cnt) = Format(WKDB("MIN" & cnt), "0." & String(tmpKeta, "0"))
                                End If
                                
                                ' 平均値
                                WKDB("AVE" & cnt) = Format(getAverage(arr), "0.000")
                                
                                ' 差
                                WKDB("R" & cnt) = WKDB("MAX" & cnt) - WKDB("MIN" & cnt)
                                If tmpKeta > 0 Then
                                    WKDB("R" & cnt) = Format(WKDB("R" & cnt), "0." & String(tmpKeta, "0"))
                                End If
                                
                                ' 検査項目№が1と3のみ計算する
                                If WSD06(cnt).S014 = "1" Or WSD06(cnt).S014 = "3" Then
                                    ' 計測回数が1回の場合は計算しない
                                    If WSD06(cnt).L019 = 1 Then
                                        ' 標準偏差
'                                        WKDB("STD" & cnt) = " /  "
                                        
                                        ' CPK
'                                        WKDB("CPK" & cnt) = " /  "
                                    Else
                                        ' 標準偏差
                                        WKDB("STD" & cnt) = Format(getStandardDeviation(arr), "0.0000")
                                        
                                        ' CPK
                                        ' 規格値1行目が0の時は表示しない
                                        ' UPDATE 2025/01/08 文字の時にエラー発生するため
'                                        If WSM18(cnt).S006 = 0 Then
                                        If Val(WSM18(cnt).S006) = 0 Then
'                                            WKDB("CPK" & cnt) = " /  "
                                        Else
                                            If IsNumeric(WSM18(cnt).S006) Then
                                                tmp規格上限 = CDbl(Trim(Nz(WSM18(cnt).S006, ""))) + WSM18(cnt).D009
                                                tmp規格下限 = CDbl(Trim(Nz(WSM18(cnt).S006, ""))) + WSM18(cnt).D010
                                                
                                                If tmp規格上限 = tmp規格下限 Or CDbl(WKDB("STD" & cnt)) = 0 Then
'                                                    WKDB("CPK" & cnt) = " /  "
                                                Else
                                                    ' （USL―X）/3×σまたは（X―LSL）/3×σのいずれかの小さい値
                                                    tmpCpk1 = (tmp規格上限 - CDbl(WKDB("AVE" & cnt))) / (3 * CDbl(WKDB("STD" & cnt)))
                                                    tmpCpk2 = (CDbl(WKDB("AVE" & cnt)) - tmp規格下限) / (3 * CDbl(WKDB("STD" & cnt)))
                                                    
                                                    If tmpCpk1 <= tmpCpk2 Then
                                                        WKDB("CPK" & cnt) = Format(tmpCpk1, "0.00")
                                                    Else
                                                        WKDB("CPK" & cnt) = Format(tmpCpk2, "0.00")
                                                    End If
                                                End If
                                            Else
                                                WKDB("CPK" & cnt) = "NaN"
                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                    End If
                End If
                
                Select Case WSM18(cnt).S008
                    ' 文字項目
                    Case "0"
                        ' 計測回数が0の場合
                        If WSD06(cnt).L019 = 0 Then
                            WKDB("K" & cnt & "-1") = WSM18(cnt).S013
                        End If
                End Select
            Next
            
            ' UPDATE 2025/03/05 SUZUKI 空欄チェック
            For cnt = 1 To maxKoumoku
                If Nz(WKDB("項目" & cnt), "") = "" Then
                    WKDB("項目" & cnt) = nullStr
                End If
                
                If Nz(WKDB("規格1-" & cnt), "") = "" And Nz(WKDB("規格2-" & cnt), "") = "" Then
                    WKDB("規格1-" & cnt) = nullStr2
                End If
                
                If Nz(WKDB("公差プラス" & cnt), "") = "" Then
                    WKDB("公差プラス" & cnt) = nullStr
                End If
                
                If Nz(WKDB("公差マイナス" & cnt), "") = "" Then
                    WKDB("公差マイナス" & cnt) = nullStr
                End If
                
                For i = 1 To 5
                    If Nz(WKDB("K" & cnt & "-" & i), "") = "" Then
                        WKDB("K" & cnt & "-" & i) = nullStr
                    End If
                Next
                
                If Nz(WKDB("MAX" & cnt), "") = "" Then
                    WKDB("MAX" & cnt) = nullStr
                End If
                
                If Nz(WKDB("MIN" & cnt), "") = "" Then
                    WKDB("MIN" & cnt) = nullStr
                End If
                
                If Nz(WKDB("AVE" & cnt), "") = "" Then
                    WKDB("AVE" & cnt) = nullStr
                End If
                
                If Nz(WKDB("R" & cnt), "") = "" Then
                    WKDB("R" & cnt) = nullStr
                End If
                
                If Nz(WKDB("STD" & cnt), "") = "" Then
                    WKDB("STD" & cnt) = nullStr
                End If
                
                If Nz(WKDB("CPK" & cnt), "") = "" Then
                    WKDB("CPK" & cnt) = nullStr
                End If
            Next
            
        WKDB.Update
        
        ' UPDATE 2025/02/17 SUZUKI 未測定項目がない場合のみ更新する
        If HEADSW = False Then
            ' 印刷時は区分の更新を行う
            If PFLG = 1 Then
                ' リード工程管理ファイル
                Sql = ""
                Sql = Sql & " UPDATE KSD04KOUT "
                Sql = Sql & " SET SD04016 = " & IIf(IsNumeric(WSD06(1).S002), WSD06(1).S002, 0) ' 測定日
                Sql = Sql & " WHERE SD04002 = " & Nz(rst!SD04002, 0)                            ' 工場コード
                Sql = Sql & "   AND SD04037 = " & Nz(rst!SD04037, 0)                            ' 登録日
                Sql = Sql & "   AND SD04004 = " & SQL_MOJICHK(rst!SD04004)                      ' 製品コード
                Sql = Sql & "   AND SD04005 = " & SQL_MOJICHK(rst!SD04005)                      ' M/Cコード
                Sql = Sql & "   AND SD04001 = " & Nz(rst!SD04001, 0)                            ' 処理連番

                Call cnn.Execute(Sql)

                ' 今は更新かけていない模様
    '            ' 検査結果ファイル
    '            Sql = ""
    '            Sql = Sql & " UPDATE KSD06KEKA "
    '            Sql = Sql & " SET SD06028 = 1 "
    '            Sql = Sql & " WHERE SD06001 = " & Nz(rst!SD04002, 0)        ' 工場コード
    '            Sql = Sql & "   AND SD06002 = " & tmp検査日                 ' 検査日<=ロット№日付で最大のもの
    '            Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(rst!SD04004)  ' 製品コード
    '            Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(rst!SD04005)  ' M/Cコード
    '            Sql = Sql & "   AND SD06011 = 1 "                           ' 製品区分
    '
    '            Call cnn.Execute(Sql)
            End If
        End If
        
NextLoop:
        
        DTCNT = DTCNT + 1
        
        rst.MoveNext
    Loop
    
    Call ShowProgressBar(, , , False)
    
    stDocName = "R_検査成績表_リード"
    
    ' UPDATE 2025/02/17 SUZUKI 未測定の項目がある場合は印刷しない
'    Call 印刷処理(Me.name, stDocName, PFLG)
    Call 印刷処理(Me.name, stDocName, PFLG, "印刷FLG = TRUE")

'    If PFLG = 0 Then
'        DoCmd.OpenReport stDocName, acPreview
'    Else
'        DoCmd.OpenReport stDocName, acNormal
'        Me.Visible = True
'    End If
    
    Me.SUB.Form.Requery
    Me.SUB.Form.FilterOn = False
    Me.SUB.Form.OrderByOn = False

Exit_PRT:
    Call CloseADORecordset(rst)
    Call CloseADORecordset(rst2)
    Call CloseADORecordset(rst3)
    Call CloseADOConnection(cnn)
    Call CloseRecordset(WKDB)
    Call CloseDB(MyDB)
    Call ShowProgressBar(, , , False)
    DoCmd.Hourglass False
    Exit Function

Err_PRT:
    ' データが1件もなかった場合は除く
    If Err.Number <> 2501 Then
        MsgBox Err.Description, , "PRT"
        Call エラーログ記録(Me.name, "PRT", Err.Number, Err.Description, "日付：" & Me.S日付 & "～" & Me.E日付 & ", 製品CD：" & Me.S製品CD & "～" & Me.E製品CD & ", M/C：" & Me.SMCCD & "～" & Me.EMCCD & ", Sql：" & Sql)
    End If
    Resume Exit_PRT
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
        Case "S工場CD"
            Call S工場CD_DblClick(False)
        Case "E工場CD"
            Call E工場CD_DblClick(False)
        Case "S製品CD"
            Call S製品CD_DblClick(False)
        Case "E製品CD"
            Call E製品CD_DblClick(False)
        Case "SMCCD"
            Call SMCCD_DblClick(False)
        Case "EMCCD"
            Call EMCCD_DblClick(False)
    End Select

Exit_P4_Click:
    Exit Sub

Err_P4_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P4_Click"
    Call エラーログ記録(Me.name, "P4_Click", Err.Number, Err.Description)
    Resume Exit_P4_Click
End Sub

Private Sub P9_Click()
    Call PRT(0)
End Sub

Private Sub P10_Click()
    Call PRT(1)
End Sub

Private Sub P11_Click()
    Call PRT(2)
End Sub

Private Sub P12_Click()
    Call 閉じる
End Sub

Private Sub S工場CD_DblClick(Cancel As Integer)
On Error GoTo Err_S工場CD_DblClick

    検索CD = Null
    DoCmd.OpenForm "K_工場検索", , , , , acDialog

    If Nz(検索CD, "") = "" Then
        Me.S工場CD.SetFocus
    Else
        Me.S工場CD = CLng(検索CD)
        Me.S工場CD.SetFocus
        Call S工場CD_AfterUpdate
    End If

Exit_S工場CD_DblClick:
    Exit Sub

Err_S工場CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "S工場CD_DblClick"
    Call エラーログ記録(Me.name, "S工場CD_DblClick", Err.Number, Err.Description)
    Resume Exit_S工場CD_DblClick
End Sub

Private Sub E工場CD_DblClick(Cancel As Integer)
On Error GoTo Err_E工場CD_DblClick

    検索CD = Null
    DoCmd.OpenForm "K_工場検索", , , , , acDialog

    If Nz(検索CD, "") = "" Then
        Me.E工場CD.SetFocus
    Else
        Me.E工場CD = CLng(検索CD)
        Me.E工場CD.SetFocus
    End If

Exit_E工場CD_DblClick:
    Exit Sub

Err_E工場CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "E工場CD_DblClick"
    Call エラーログ記録(Me.name, "E工場CD_DblClick", Err.Number, Err.Description)
    Resume Exit_E工場CD_DblClick
End Sub

Private Sub S製品CD_DblClick(Cancel As Integer)
On Error GoTo Err_S製品CD_DblClick

    検索CD = Null
    P_CODE = Me.製品区分

    DoCmd.OpenForm "K_製品検索", , , , , acDialog

    If Nz(検索CD, "") = "" Then
        Me.S製品CD.SetFocus
    Else
        Me.S製品CD = 検索CD
        Me.S製品CD.SetFocus
        Call S製品CD_AfterUpdate
    End If

Exit_S製品CD_DblClick:
    Exit Sub

Err_S製品CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "S製品CD_DblClick"
    Call エラーログ記録(Me.name, "S製品CD_DblClick", Err.Number, Err.Description)
    Resume Exit_S製品CD_DblClick
End Sub

Private Sub E製品CD_DblClick(Cancel As Integer)
On Error GoTo Err_E製品CD_DblClick

    検索CD = Null
    P_CODE = Me.製品区分

    DoCmd.OpenForm "K_製品検索", , , , , acDialog

    If Nz(検索CD, "") = "" Then
        Me.E製品CD.SetFocus
    Else
        Me.E製品CD = 検索CD
        Me.E製品CD.SetFocus
    End If

Exit_E製品CD_DblClick:
    Exit Sub

Err_E製品CD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "E製品CD_DblClick"
    Call エラーログ記録(Me.name, "E製品CD_DblClick", Err.Number, Err.Description)
    Resume Exit_E製品CD_DblClick
End Sub

Private Sub SMCCD_DblClick(Cancel As Integer)
On Error GoTo Err_SMCCD_DblClick

    検索CD = Null
    If Nz(Me.S工場CD, "") = "" Then
        MsgBox "工場コードを入力してください", vbCritical, "エラー"
        Exit Sub
    End If
    P_CODE = Me.S工場CD
    DoCmd.OpenForm "K_マシン検索", , , , , acDialog
    
    If Nz(検索CD, "") = "" Then
        Me(ActCon).SetFocus
    Else
        Me(ActCon) = 検索CD
        Call SMCCD_AfterUpdate
    End If

Exit_SMCCD_DblClick:
    Exit Sub

Err_SMCCD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "SMCCD_DblClick"
    Call エラーログ記録(Me.name, "SMCCD_DblClick", Err.Number, Err.Description)
    Resume Exit_SMCCD_DblClick
End Sub

Private Sub EMCCD_DblClick(Cancel As Integer)
On Error GoTo Err_EMCCD_DblClick

    検索CD = Null
    If Nz(Me.E工場CD, "") = "" Then
        MsgBox "工場コードを入力してください", vbCritical, "エラー"
        Exit Sub
    End If
    P_CODE = Me.E工場CD
    DoCmd.OpenForm "K_マシン検索", , , , , acDialog
    
    If Nz(検索CD, "") = "" Then
        Me(ActCon).SetFocus
    Else
        Me(ActCon) = 検索CD
    End If

Exit_EMCCD_DblClick:
    Exit Sub

Err_EMCCD_DblClick:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "EMCCD_DblClick"
    Call エラーログ記録(Me.name, "EMCCD_DblClick", Err.Number, Err.Description)
    Resume Exit_EMCCD_DblClick
End Sub

Private Sub S工場CD_AfterUpdate()
    If Nz(Me.S工場CD, "") <> "" Then
        Me.E工場CD = Me.S工場CD
    End If
End Sub

Private Sub S日付_AfterUpdate()
    If Nz(Me.S日付, "") <> "" Then
        Me.E日付 = Me.S日付
    End If
End Sub

Private Sub S製品CD_AfterUpdate()
    If Nz(Me.S製品CD, "") <> "" Then
        Me.E製品CD = Me.S製品CD
    End If
End Sub

Private Sub SMCCD_AfterUpdate()
    If Nz(Me.SMCCD, "") <> "" Then
        Me.EMCCD = Me.SMCCD
    End If
End Sub

Private Sub S日付_DblClick(Cancel As Integer)
On Error Resume Next

    Call CallCalendar
    
    Call S日付_AfterUpdate
End Sub

Private Sub E日付_DblClick(Cancel As Integer)
On Error Resume Next

    Call CallCalendar
End Sub

Private Sub 明細へ_Enter()
On Error Resume Next

    Me.P9.SetFocus
End Sub

