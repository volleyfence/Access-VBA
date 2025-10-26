VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_D_測定処理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' 測定器の同時入力を行えるようにする
' VBAでは並列処理(スレッド)ができないので、時間でCOMポートを変えて処理を行い、並列処理に見せかけている

Private Const maxPort As Long = 3

' スキップを認めるか
Private Const allowSkip As Boolean = False

' 前回結果の修正を認めるか
Private Const allowEdit As Boolean = False

' 手入力を認めるか(テスト用)
Private Const allowInput As Boolean = False

' ポート番号(1～4)からアルファベットを取得する用(FRMCLSで格納)
Private portInfo1 As Variant

' ポート番号(1～4)からを項目数を取得する用(FRMCLSで格納)
Private portInfo2 As Variant

Private currentPort As Long

Private Comm(0 To maxPort) As Object

Private cnn_common As ADODB.Connection
Private rst_common(0 To maxPort) As ADODB.Recordset           ' Index => 1:A / 2:B / 3:C / 4:D

Private SD06KEY(0 To maxPort, 0 To 3) As SD06KEY_S          ' Index => 1:A / 2:B / 3:C / 4:D, 1:前回 / 2:現在 / 3:次回
Private SD06DAT(0 To maxPort, 0 To 4) As SD06_S             ' Index => 1:A / 2:B / 3:C / 4:D, 1:項目1 / 2:項目2 / 3:項目3 / 4:項目4

Private ZSD06DAT(0 To maxPort, 0 To 4) As SD06_S            ' Index => 1:A / 2:B / 3:C / 4:D, 1:項目1 / 2:項目2 / 3:項目3 / 4:項目4 前回結果用

Private KeisokuFLG(0 To maxPort, 0 To 4) As Boolean         ' Index => 1:A / 2:B / 3:C / 4:D, 1:項目1 / 2:項目2 / 3:項目3 / 4:項目4　True: 計測終了

#If Win64 Then
    Private KeisokuCnt(0 To maxPort, 0 To 2) As LongLong    ' Index => 1:A / 2:B / 3:C / 4:D, 1:総計 / 2:計測回数　(画面を開いてからDATAWRITEの回数を記録する)
#Else
    Private KeisokuCnt(0 To maxPort, 0 To 2) As Long        ' Index => 1:A / 2:B / 3:C / 4:D, 1:総計 / 2:計測回数　(画面を開いてからDATAWRITEの回数を記録する)
#End If

Private Sub Form_Close()
On Error GoTo Err_Form_Close
    
    Dim port As Long
    
    ' 終了処理
    For port = 1 To maxPort
        Call CloseComm(port)
        
        Call CloseADORecordset(rst_common(port))
    Next
    Call CloseADOConnection(cnn_common)

Exit_Form_Close:
    Exit Sub
    
Err_Form_Close:
    MsgBox Err.Description, , "Form_Close"
    Call エラーログ記録(Me.name, "Form_Close", Err.Number, Err.Description)
    Resume Exit_Form_Close
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Form_Open
    
    Dim port As Long
    Dim 測定データ存在 As Boolean
    
    PreForm = AFNAME
    AFNAME = Me.name

    Application.SetOption KEYMACRO, "入力キー"
    Me.Visible = True

    Call 初期設定(Me.name)
    
    Call FRMCLS
    
    ' 共通cnnセット
    Set cnn_common = New ADODB.Connection
    
    cnn_common.ConnectionString = ACPASS
    cnn_common.CommandTimeout = Timeout
    cnn_common.Open
    
    ' rst_commonセット
    測定データ存在 = False
    For port = 1 To maxPort
        If setRST(port) Then
            測定データ存在 = True
        End If
    Next
    
    If 測定データ存在 = False Then
        MsgBox "測定データがありません", vbCritical, "エラー"
        DoCmd.Close acForm, Me.name
        Exit Sub
    End If
    
    ' SD06KEY, SD06DATセット
    For port = 1 To maxPort
        Call DATAREAD(port, True)
    Next
    
    ' INIファイルからポート情報取得
    Call GETPORT
    
    ' COMポート初期化
    Call ClrComm
    
    ' COMポートオープン①
    For port = 1 To maxPort
        Call OpenComm(port)
    Next
    
    ' COMポートオープン②
    ' ①でエラーが起きた際は全てCloseしてしまうので再度開く
    For port = 1 To maxPort
        Call OpenComm(port)
    Next

Exit_Form_Open:
    DoCmd.Hourglass False
    Exit Sub
    
Err_Form_Open:
    MsgBox Err.Description, , "Form_Open"
    Call エラーログ記録(Me.name, "Form_Open", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_Form_Open
End Sub

Private Sub Form_Timer()
On Error Resume Next

    Dim port As Long
    
    For port = 1 To maxPort
        If Comm(port).PortOpen = True Then
            Call OnComm(port)
        End If
    Next
End Sub

Public Function FRMCLS()
On Error Resume Next
    
    Dim alp As String
    Dim i As Long

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
    
    Me.工場CD = P_CODE
    Me.日付 = P_CODE2
    Me.モード = P_CODE3 ' 0:通常　1:追加　2:続き　3:スキップ
    
    ' 画面位置文字と項目数(1番目は使用しない)
    portInfo1 = Array("", "A", "B", "C", "D")
    portInfo2 = Array(0, 4, 2, 2, 2)
    
    For i = 1 To maxPort
        alp = portInfo1(i)
        
        Me("修正登録" & alp).Visible = False
        
        If allowSkip Then
            Me("スキップ" & alp).Visible = True
        Else
            Me("スキップ" & alp).Visible = False
        End If
    Next
End Function

Private Function GETPORT()
On Error GoTo Err_GETPORT

    Dim tmpIni As Variant
    Dim port As Long
    Dim alp As String
    
    For port = 1 To maxPort
        alp = portInfo1(port)
        
        tmpIni = Nz(ReadIniData("PORT", "ポート" & alp), 0)
        
        If tmpIni = 0 Then
            MsgBox "INIファイルが不正です", vbInformation, "エラー"
            Exit Function
        End If
        
        Me("ポート" & alp) = tmpIni
    Next

Exit_GETPORT:
    Exit Function

Err_GETPORT:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "GETPORT"
    Call エラーログ記録(Me.name, "GETPORT", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_GETPORT
End Function

Private Function ClrComm()
On Error Resume Next

    Dim port As Long
    
    For port = 1 To maxPort
        Set Comm(port) = CreateObject("MSCommLib.MSComm")
        Me("COM状態" & port) = True
    Next
End Function

' port -> 1:A / 2:B / 3:C / 4:D
Private Function OpenComm(port As Long)
On Error GoTo Err_OpenComm
    
    Dim alp As String
    
    If Comm(port).PortOpen = True Then
        Exit Function
    End If
    
    If Me("COM状態" & port) = False Then
        Exit Function
    End If
    
    alp = portInfo1(port)

    ' COMポート番号を指定
    Comm(port).CommPort = Me("ポート" & alp)
    
    ' 通信設定を指定
    Comm(port).Settings = "1200,N,8,1"
    
    Comm(port).InputLen = 0
    Comm(port).InBufferSize = 1024
    Comm(port).RThreshold = 1
    
    ' COMポートを開く
    Comm(port).PortOpen = True
    
    Me("COM状態" & port) = True

Exit_OpenComm:
    Exit Function

Err_OpenComm:
    MsgBox "COMポートのOPENに失敗しました。" & vbNewLine & "ポート番号 " & Me("ポート" & alp), vbCritical, "OpenComm"
    Me("COM状態" & port) = False
    ' 全てCloseしないとなぜか他が動かない
    Call CloseComm
    Resume Exit_OpenComm
End Function

' port -> 1:A / 2:B / 3:C / 4:D
Private Function OnComm(port As Long)
On Error GoTo Err_OnComm
    
    Dim comInput As Variant
    Dim comEvent As Long
    Dim msgStr As String
    Dim alp As String
    Dim koumokuNum As Long
    Dim X As Long   ' 横位置(項目)
    Dim Y As Long   ' 縦位置(測定値)
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    comEvent = Comm(port).CommEvent
    DoEvents
    
    Select Case comEvent
        Case comEventBreak
            msgStr = "中断信号が受信されました。"
            GoTo BreakComm
        Case ComEventCDTO
            msgStr = "CD (RLSD) タイムアウトです。"
            GoTo BreakComm
        Case comEventCTSTO
            msgStr = "CTS タイムアウトです。"
            GoTo BreakComm
        Case comEventDSRTO
            msgStr = "DSR タイムアウトです。"
            GoTo BreakComm
        Case ComEventFrame
            msgStr = "フレーム エラーです。"
            GoTo BreakComm
        Case ComEventOverrun
            msgStr = "データが失われました。"
            GoTo BreakComm
        Case ComEventRxOver
            msgStr = "受信バッファがオーバーフローしました。"
            GoTo BreakComm
        Case ComEventRxParity
            msgStr = "パリティ エラーです。"
            GoTo BreakComm
        Case ComEventTxFull
            msgStr = "送信バッファがいっぱいです。"
            GoTo BreakComm
        Case ComEventDCB
            msgStr = "デバイス コントロール ブロック (DCB)が取得される予期しないエラーです。"
            GoTo BreakComm
        
        ' RThreshold プロパティで指定された数の文字を受信しました。
        Case comEvReceive
            comInput = Comm(port).Input
            
            ' 2回に分けて返される場合があるのでもう1回取得する
            comInput = comInput & Comm(port).Input
            
            ' このポートの測定が終了している場合は受信しない
            If Me("測定FLG_" & alp) = False Then
                GoTo Exit_OnComm
            End If
        
            If Left(comInput, 1) = "9" Then
                Select Case Mid(comInput, 3, 1)
                    Case 1
                        Call MsgBox("データが入力できません。" & vbCrLf & "計測機器の電源をONにして下さい。", vbCritical + vbOKOnly, Me.Caption)
                    Case 2
                        Call MsgBox("取り込んだデータが規定フォーマット外", vbCritical + vbOKOnly, Me.Caption)
                End Select
                comInput = ""
            End If
    
            comInput = Right(comInput, 9)
            If IsNumeric(comInput) Then
                X = ((Nz(Me("回数" & alp), 0) - 1) Mod koumokuNum) + 1
                Y = Int((Nz(Me("回数" & alp), 0) - 1) / koumokuNum) + 1
                
                Me("測定値" & Y & alp & X) = CDbl(comInput)
                
                If CHK_DATA(port, CDbl(comInput)) Then
                    Call NextRead(port)
                End If
            End If
            DoEvents
        End Select

Exit_OnComm:
    Exit Function

Err_OnComm:
    Resume Exit_OnComm

BreakComm:
    MsgBox msgStr & vbNewLine & "COM：" & port, vbCritical, "OnComm"
    DoCmd.Close acForm, Me.name
    GoTo Exit_OnComm
End Function

' port -> 1:A / 2:B / 3:C / 4:D　引数なしなら全て
Private Function CloseComm(Optional port As Long = 0)
On Error GoTo Err_CloseComm
    
    Dim i As Long
    
    If port = 0 Then
        For i = 1 To maxPort
            If Comm(i).PortOpen Then
                Comm(i).PortOpen = False
            End If
        Next
    Else
        If Comm(port).PortOpen Then
            Comm(port).PortOpen = False
        End If
    End If

Exit_CloseComm:
    Exit Function

Err_CloseComm:
'    MsgBox Err.Description, vbCritical, "CloseComm"
    Resume Exit_CloseComm
End Function

' port -> 1:A / 2:B / 3:C / 4:D
Private Function setRST(port As Long) As Boolean
On Error GoTo Err_setRST
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim Sql As String
    Dim subSql As String
    Dim alp As String
    Dim Crit As String
    
    setRST = False
    
    alp = portInfo1(port)
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rst2 = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    Crit = ""
    Crit = Crit & "       SD06001 = " & Nz(Me.工場CD, 0) & vbNewLine              ' 工場コード
    Crit = Crit & "   AND SD06002 = " & Format(Me.日付, "yyyymmdd") & vbNewLine   ' 処理日
    Crit = Crit & "   AND SD06003 = 1 " & vbNewLine                               ' 自動手動区分
    Crit = Crit & "   AND SD06004 = " & SQL_MOJICHK(alp) & vbNewLine              ' 画面位置
'    Crit = Crit & "   AND SD06025 = 0 " & vbNewLine                               ' 計測済み区分
    ' 通常
    If Nz(Me.モード, 0) = 0 Then
        Crit = Crit & "   AND SD06027 = 0 " & vbNewLine                           ' 通常追加区分
        
    ' 追加
    ElseIf Nz(Me.モード, 0) = 1 Then
        Crit = Crit & "   AND SD06027 = 1 " & vbNewLine                           ' 通常追加区分
    End If
    
    ' スキップ
    If Nz(Me.モード, 0) = 3 Then
        Crit = Crit & "   AND SD06029 = 1 " & vbNewLine                           ' スキップ区分
        Crit = Crit & "   AND SD06030 = 1 " & vbNewLine                           ' スキップ測定区分
    Else
        Crit = Crit & "   AND SD06029 = 0 " & vbNewLine                           ' スキップ区分
    End If
    
    ' 計測回数の合計を計算する(DATEWRITEの回数)
    Sql = ""
    Sql = Sql & " SELECT SD06005, SD06006, SD06007, SD06008, SD06009 FROM KSD06KEKA " & vbNewLine
    Sql = Sql & " WHERE " & Crit & vbNewLine
    Sql = Sql & "   AND SD06025 = 0 " & vbNewLine                               ' 計測済み区分
    Sql = Sql & " GROUP BY SD06005, SD06006, SD06007, SD06008, SD06009 " & vbNewLine
    
    rst.Open Sql, cnn, adOpenKeyset, adLockReadOnly
    
    ' UPDATE 2024/12/18 測定済み回数のカウント
    Sql = ""
    Sql = Sql & " SELECT SD06005, SD06006, SD06007, SD06008, SD06009 FROM KSD06KEKA " & vbNewLine
    Sql = Sql & " WHERE " & Crit & vbNewLine
    Sql = Sql & "   AND SD06025 = 1 " & vbNewLine                               ' 計測済み区分
    Sql = Sql & " GROUP BY SD06005, SD06006, SD06007, SD06008, SD06009 " & vbNewLine
    
    rst2.Open Sql, cnn, adOpenKeyset, adLockReadOnly
    
    If rst.EOF = False Then
        rst.MoveLast
        KeisokuCnt(port, 1) = rst.RecordCount
        If rst2.EOF Then
            KeisokuCnt(port, 2) = 0
        Else
            rst2.MoveLast
            KeisokuCnt(port, 2) = rst2.RecordCount
        End If
        Me("未測定" & alp) = KeisokuCnt(port, 1)
        Me("測定済" & alp) = KeisokuCnt(port, 2)
    End If
    
    rst.Close
    rst2.Close
    
    Set rst_common(port) = New ADODB.Recordset
    
    ' 画面を開いている間はこのレコードセットを使い続ける
    Sql = ""
    Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
    Sql = Sql & " WHERE " & Crit & vbNewLine
    Sql = Sql & "   AND SD06025 = 0 " & vbNewLine                               ' 計測済み区分
    Sql = Sql & " ORDER BY SD06001, SD06002, SD06003, SD06004, SD06005, "       ' 工場コード, 処理日, 自動手動区分, 画面位置, 設定順
    Sql = Sql & "          SD06006, SD06007, SD06008, SD06010 "                 ' ライン順, M/C順, ライン№, 項目順
    
    rst_common(port).Open Sql, cnn_common, adOpenKeyset, adLockReadOnly
    
    If rst_common(port).EOF Then
        Me("測定FLG_" & alp) = False
    Else
        setRST = True
        Me("測定FLG_" & alp) = True
    End If
    
Exit_setRST:
    Call CloseADORecordset(rst)
    Call CloseADORecordset(rst2)
    Call CloseADOConnection(cnn)
    Exit Function

Err_setRST:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "setRST"
    Call エラーログ記録(Me.name, "setRST", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_setRST
End Function

' port      -> 1:A / 2:B / 3:C / 4:D
' FirstREAD -> 画面を開いて最初の読み込みか
Public Function DATAREAD(port As Long, Optional FirstREAD As Boolean = False) As Boolean
On Error GoTo Err_DATAREAD
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim iPart As Long
    Dim alp As String
    Dim koumokuNum As Long
    Dim i As Long
    Dim Sql As String
    
    DATAREAD = False
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    ' 今回→前回
    With SD06KEY(port, 1)
        .S001 = SD06KEY(port, 2).S001
        .S002 = SD06KEY(port, 2).S002
        .S003 = SD06KEY(port, 2).S003
        .S004 = SD06KEY(port, 2).S004
        .L005 = SD06KEY(port, 2).L005
        .L006 = SD06KEY(port, 2).L006
        .L007 = SD06KEY(port, 2).L007
        .L008 = SD06KEY(port, 2).L008
        .S009 = SD06KEY(port, 2).S009
        .L010 = SD06KEY(port, 2).L010
        .S012 = SD06KEY(port, 2).S012
        .S040 = SD06KEY(port, 2).S040
        .S050 = SD06KEY(port, 2).S050
    End With
    
    ' 次回→今回
    With SD06KEY(port, 2)
        .S001 = SD06KEY(port, 3).S001
        .S002 = SD06KEY(port, 3).S002
        .S003 = SD06KEY(port, 3).S003
        .S004 = SD06KEY(port, 3).S004
        .L005 = SD06KEY(port, 3).L005
        .L006 = SD06KEY(port, 3).L006
        .L007 = SD06KEY(port, 3).L007
        .L008 = SD06KEY(port, 3).L008
        .S009 = SD06KEY(port, 3).S009
        .L010 = SD06KEY(port, 3).L010
        .S012 = SD06KEY(port, 3).S012
        .S040 = SD06KEY(port, 3).S040
        .S050 = SD06KEY(port, 3).S050
    End With
    
    Call CLR_SD06KEY(SD06KEY(port, 3))
    
    If rst_common(port).EOF Then
        SD06KEY(port, 2).S040 = "END"
        Me("測定FLG_" & alp) = False
    End If
    
    ' 初期化
    For iPart = 1 To koumokuNum
        Call CLR_SD06(SD06DAT(port, iPart))
        
        ' 測定終了時は初期化しない
        If Me("測定FLG_" & alp) = True Then
            ' 測定完了フラグ初期化
            KeisokuFLG(port, iPart) = False
        End If
        
        Me("項目名" & alp & iPart) = Null
        Me("規格" & alp & iPart) = Null
        Me("公差プラス" & alp & iPart) = Null
        Me("公差マイナス" & alp & iPart) = Null
        For i = 1 To 5
            Me("測定値" & i & alp & iPart) = Null
            Me("測定値" & i & alp & iPart).Enabled = False
            Me("測定値" & i & alp & iPart).Locked = False
        Next
    Next
    
    If Me("測定FLG_" & alp) = True Then
        ' 初回データは現在に格納
        If FirstREAD Then
            With SD06KEY(port, 2)
                .S001 = Nz(rst_common(port)!SD06001, 0)
                .S002 = Format(rst_common(port)!SD06002, "00000000")
                .S003 = Nz(rst_common(port)!SD06003, 0)
                .S004 = rst_common(port)!SD06004
                .L005 = rst_common(port)!SD06005
                .L006 = rst_common(port)!SD06006
                .L007 = rst_common(port)!SD06007
                .L008 = rst_common(port)!SD06008
                .S009 = rst_common(port)!SD06009
                .L010 = rst_common(port)!SD06010
                .S012 = rst_common(port)!SD06012
                .S040 = rst_common(port)!SD06009
                .S050 = rst_common(port)!SD06013
            End With
        End If
        
        Do Until rst_common(port).EOF
            rst_common(port).MoveNext
            
            ' 次のデータが無い場合はENDとする
            If rst_common(port).EOF Then
                SD06KEY(port, 3).S040 = "END"
                SD06KEY(port, 3).S050 = ""
            
            ' データがある場合は次回データに格納する
            Else
                With SD06KEY(port, 3)
                    .S001 = Nz(rst_common(port)!SD06001, 0)
                    .S002 = Format(rst_common(port)!SD06002, "00000000")
                    .S003 = Nz(rst_common(port)!SD06003, 0)
                    .S004 = rst_common(port)!SD06004
                    .L005 = rst_common(port)!SD06005
                    .L006 = rst_common(port)!SD06006
                    .L007 = rst_common(port)!SD06007
                    .L008 = rst_common(port)!SD06008
                    .S009 = rst_common(port)!SD06009
                    .L010 = rst_common(port)!SD06010
                    .S012 = rst_common(port)!SD06012
                    .S040 = rst_common(port)!SD06009
                    .S050 = rst_common(port)!SD06013
                End With
            End If
            
            If (SD06KEY(port, 2).L005 <> SD06KEY(port, 3).L005) _
            Or (SD06KEY(port, 2).L006 <> SD06KEY(port, 3).L006) _
            Or (SD06KEY(port, 2).L007 <> SD06KEY(port, 3).L007) _
            Or (SD06KEY(port, 2).L008 <> SD06KEY(port, 3).L008) _
            Or (SD06KEY(port, 2).S009 <> SD06KEY(port, 3).S009) Then
                Exit Do
            End If
        Loop
        
        Me("回数" & alp) = 1
        
        ' 項目別に呼び出し
        For iPart = 1 To koumokuNum
            Sql = ""
            Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
            Sql = Sql & " WHERE SD06001 = " & SD06KEY(port, 2).S001 & vbNewLine                 ' 工場コード
            Sql = Sql & "   AND SD06002 = " & SD06KEY(port, 2).S002 & vbNewLine                 ' 処理日
            Sql = Sql & "   AND SD06003 = " & SD06KEY(port, 2).S003 & vbNewLine                 ' 自動手動区分
            Sql = Sql & "   AND SD06004 = " & SQL_MOJICHK(SD06KEY(port, 2).S004) & vbNewLine    ' 画面位置
            Sql = Sql & "   AND SD06005 = " & SD06KEY(port, 2).L005 & vbNewLine                 ' 設定順
            Sql = Sql & "   AND SD06006 = " & SD06KEY(port, 2).L006 & vbNewLine                 ' ライン順
            Sql = Sql & "   AND SD06007 = " & SD06KEY(port, 2).L007 & vbNewLine                 ' Ｍ／Ｃ順
            Sql = Sql & "   AND SD06008 = " & SD06KEY(port, 2).L008 & vbNewLine                 ' ライン№
            Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(SD06KEY(port, 2).S009) & vbNewLine    ' Ｍ／Ｃコード
            Sql = Sql & "   AND SD06010 = " & iPart & vbNewLine                                 ' 項目順
            Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(SD06KEY(port, 2).S012) & vbNewLine    ' 製品コード
            Sql = Sql & " ORDER BY SD06001, SD06002, SD06003, SD06004, SD06005, "               ' 工場コード, 処理日, 自動手動区分, 画面位置, 設定順
            Sql = Sql & "          SD06006, SD06007, SD06008, SD06010 "                         ' ライン順, M/C順, ライン№, 項目順
            
            rst.Open Sql, cnn, adOpenKeyset, adLockReadOnly
            
            If rst.EOF = False Then
                Call SET_SD06(SD06DAT(port, iPart), rst)
                
                ' 画面表示
                Me("項目名" & alp & iPart) = SD06DAT(port, iPart).S015
                Me("規格" & alp & iPart) = SD06DAT(port, iPart).D016
                Me("公差プラス" & alp & iPart) = SD06DAT(port, iPart).D017
                Me("公差マイナス" & alp & iPart) = SD06DAT(port, iPart).D018
                
                For i = 1 To Nz(SD06DAT(port, iPart).L019, 1)
                    Me("測定値" & i & alp & iPart).Enabled = True
                Next
            End If
            
            ' 項目が入っていなかったらその項目の測定は完了とする
            If Nz(SD06DAT(port, iPart).S001, "") = "" Then
                KeisokuFLG(port, iPart) = True
                If Nz(Me("回数" & alp), 0) = iPart Then
                    Me("回数" & alp) = Me("回数" & alp) + 1
                End If
            End If
            
            rst.Close
        Next
    End If
    
    Me("前回" & alp) = SD06KEY(port, 1).S040
    Me("今回" & alp) = SD06KEY(port, 2).S040
    Me("次回" & alp) = SD06KEY(port, 3).S040
    Me("製品名" & alp) = SD06KEY(port, 2).S050
    
    Call FOCUS(port)
    
    DATAREAD = True
    
Exit_DATAREAD:
    Call CloseADORecordset(rst)
    Call CloseADOConnection(cnn)
    Exit Function

Err_DATAREAD:
    MsgBox Err.Description, , "DATAREAD"
    Call エラーログ記録(Me.name, "DATAREAD", Err.Number, Err.Description, "モード：" & Me.モード & ", Sql：" & Sql)
    Resume Exit_DATAREAD
End Function

' port      -> 1:A / 2:B / 3:C / 4:D
' EditFLG   -> 前回結果の修正
' SkipFLG   -> スキップするデータか
Public Function DATAWRITE(port As Long, Optional EditFLG As Boolean = False, Optional SkipFLG As Boolean) As Boolean
On Error GoTo Err_DATAWRITE
    
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim iPart As Long
    Dim alp As String
    Dim koumokuNum As Long
    Dim i As Long
    Dim KBN As Long
    Dim fieldNmAdd As String
    Dim Sql As String
    
    DATAWRITE = False
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    ' 前回
    If EditFLG = True Then
        KBN = 1
        fieldNmAdd = "Z"
    
    ' 今回
    Else
        KBN = 2
        fieldNmAdd = ""
    End If
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cnn.ConnectionString = ACPASS
    cnn.CommandTimeout = Timeout
    cnn.Open
    
    cnn.BeginTrans
    
    ' 項目別に登録
    For iPart = 1 To koumokuNum
        If Nz(Me(fieldNmAdd & "項目名" & alp & iPart), "") <> "" Then
            Sql = ""
            Sql = Sql & " SELECT * FROM KSD06KEKA " & vbNewLine
            Sql = Sql & " WHERE SD06001 = " & SD06KEY(port, KBN).S001 & vbNewLine               ' 工場コード
            Sql = Sql & "   AND SD06002 = " & SD06KEY(port, KBN).S002 & vbNewLine               ' 処理日
            Sql = Sql & "   AND SD06003 = " & SD06KEY(port, KBN).S003 & vbNewLine               ' 自動手動区分
            Sql = Sql & "   AND SD06004 = " & SQL_MOJICHK(SD06KEY(port, KBN).S004) & vbNewLine  ' 画面位置
            Sql = Sql & "   AND SD06005 = " & SD06KEY(port, KBN).L005 & vbNewLine               ' 設定順
            Sql = Sql & "   AND SD06006 = " & SD06KEY(port, KBN).L006 & vbNewLine               ' ライン順
            Sql = Sql & "   AND SD06007 = " & SD06KEY(port, KBN).L007 & vbNewLine               ' Ｍ／Ｃ順
            Sql = Sql & "   AND SD06008 = " & SD06KEY(port, KBN).L008 & vbNewLine               ' ライン№
            Sql = Sql & "   AND SD06009 = " & SQL_MOJICHK(SD06KEY(port, KBN).S009) & vbNewLine  ' Ｍ／Ｃコード
            Sql = Sql & "   AND SD06010 = " & iPart & vbNewLine                                 ' 項目順
            Sql = Sql & "   AND SD06012 = " & SQL_MOJICHK(SD06KEY(port, KBN).S012) & vbNewLine  ' 製品コード
            Sql = Sql & " ORDER BY SD06001, SD06002, SD06003, SD06004, SD06005, "               ' 工場コード, 処理日, 自動手動区分, 画面位置, 設定順
            Sql = Sql & "          SD06006, SD06007, SD06008, SD06010 "                         ' ライン順, M/C順, ライン№, 項目順
                
            rst.Open Sql, cnn, adOpenKeyset, adLockOptimistic
            
            If rst.EOF = False Then
                    If SkipFLG = False Then
                        rst!SD06020 = Nz(Me(fieldNmAdd & "測定値1" & alp & iPart), 0)
                        rst!SD06021 = Nz(Me(fieldNmAdd & "測定値2" & alp & iPart), 0)
                        rst!SD06022 = Nz(Me(fieldNmAdd & "測定値3" & alp & iPart), 0)
                        rst!SD06023 = Nz(Me(fieldNmAdd & "測定値4" & alp & iPart), 0)
                        rst!SD06024 = Nz(Me(fieldNmAdd & "測定値5" & alp & iPart), 0)
                        
                        rst!SD06025 = 1
                        rst!SD06029 = 0
                    Else
                        rst!SD06025 = 0
                        rst!SD06029 = 1
                    End If
                    
                    rst!変更日 = Now()
                
                rst.Update
            End If
                
            rst.Close
        End If
    Next
    
    cnn.CommitTrans
    
    KeisokuCnt(port, 1) = KeisokuCnt(port, 1) - 1
    If SkipFLG = False Then
        KeisokuCnt(port, 2) = KeisokuCnt(port, 2) + 1
    End If
    Me("未測定" & alp) = KeisokuCnt(port, 1)
    Me("測定済" & alp) = KeisokuCnt(port, 2)
    
    DATAWRITE = True
    
Exit_DATAWRITE:
    Call CloseADORecordset(rst)
    Call CloseADOConnection(cnn)
    Exit Function

Err_DATAWRITE:
    cnn.RollbackTrans
    MsgBox Err.Description, , "DATAWRITE"
    Call エラーログ記録(Me.name, "DATAWRITE", Err.Number, Err.Description, "モード：" & Me.モード & ", Sql：" & Sql)
    Resume Exit_DATAWRITE
End Function

' フォーカスを初期化して現在の測定項目の背景色を変える
' port -> 1:A / 2:B / 3:C / 4:D
Private Function FOCUS(port As Long)
On Error Resume Next

    Dim X As Long   ' 横位置(項目)
    Dim Y As Long   ' 縦位置(測定値)
    Dim alp As String
    Dim koumokuNum As Long
    Dim EndFlg As Boolean
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    Me.P1.SetFocus
    
    EndFlg = True
    For X = 1 To koumokuNum
        ' 未測定の項目があればフラグを立てる
        If KeisokuFLG(port, X) = False Then
            EndFlg = False
        End If
        
        If Nz(SD06DAT(port, X).L019, 1) <> 0 Then
            For Y = 1 To Nz(SD06DAT(port, X).L019, 1)
                Me("測定値" & Y & alp & X).BackColor = RGB(255, 255, 255)
                Me("測定値" & Y & alp & X).ForeColor = RGB(0, 0, 0)
                Me("測定値" & Y & alp & X).Enabled = False
                Me("測定値" & Y & alp & X).Locked = True
            Next
        Else
            For Y = 1 To 5
                Me("測定値" & Y & alp & X).BackColor = RGB(255, 255, 255)
                Me("測定値" & Y & alp & X).ForeColor = RGB(0, 0, 0)
                Me("測定値" & Y & alp & X).Enabled = False
                Me("測定値" & Y & alp & X).Locked = True
            Next
        End If
    Next
    
    ' 計測中の場合のみ
    If EndFlg = False Then
        X = ((Nz(Me("回数" & alp), 0) - 1) Mod koumokuNum) + 1
        Y = Int((Nz(Me("回数" & alp), 0) - 1) / koumokuNum) + 1
        
        Me("測定値" & Y & alp & X).BackColor = RGB(0, 255, 255)
        If allowInput Then
            Me("測定値" & Y & alp & X).Enabled = True
            Me("測定値" & Y & alp & X).Locked = False
            Me("測定値" & Y & alp & X).SetFocus
            Call GETACTIVE
        End If
    End If
End Function

' port -> 1:A / 2:B / 3:C / 4:D
Public Function CHK_DATA(port As Long, impData As Double) As Boolean
On Error GoTo Err_CHK_DATA
    
    Dim X As Long   ' 横位置(項目)
    Dim Y As Long   ' 縦位置(測定値)
    Dim alp As String
    Dim koumokuNum As Long
    Dim i As Long
    
    CHK_DATA = False
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    X = ((Nz(Me("回数" & alp), 0) - 1) Mod koumokuNum) + 1
    Y = Int((Nz(Me("回数" & alp), 0) - 1) / koumokuNum) + 1
    
    ' (規格 + 公差マイナス) ～ (規格 + 公差プラス)の範囲内か
    If ((SD06DAT(port, X).D016 + SD06DAT(port, X).D018) <= impData) And (impData <= (SD06DAT(port, X).D016 + SD06DAT(port, X).D017)) Then
        CHK_DATA = True
    Else
        Me("測定値" & Y & alp & X).ForeColor = RGB(255, 0, 0)
        
        ' 2024/12/09
        Call Beep
    End If
    
Exit_CHK_DATA:
    Exit Function

Err_CHK_DATA:
    MsgBox Err.Description, , "CHK_DATA"
    Call エラーログ記録(Me.name, "CHK_DATA", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_CHK_DATA
End Function

' 測定値が正常の場合の次の処理
' port -> 1:A / 2:B / 3:C / 4:D
Public Function NextRead(port As Long) As Boolean
On Error GoTo Err_NextRead
    
    Dim X As Long   ' 横位置(項目)
    Dim Y As Long   ' 縦位置(測定値)
    Dim alp As String
    Dim koumokuNum As Long
    Dim i As Long
    
    NextRead = False
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    X = ((Nz(Me("回数" & alp), 0) - 1) Mod koumokuNum) + 1
    Y = Int((Nz(Me("回数" & alp), 0) - 1) / koumokuNum) + 1
    
    ' 計測回数分の測定が完了したらこの項目は完了とする
    If SD06DAT(port, X).L019 <= Y Then
        KeisokuFLG(port, X) = True
        
        ' ポート内の全項目の測定が完了したら書き込む
        For i = 1 To koumokuNum
            ' 完了していない項目がある場合
            If KeisokuFLG(port, i) = False Then
                Call COUNTUP(port)
                Call FOCUS(port)
                Exit Function
            End If
        Next
        
        Call DATAWRITE(port)
        
        ' 前回結果タブに転記
        Call ZENKAI_SET(port)
        
        ' 次のデータの読み込み
        Call DATAREAD(port)
    Else
        Call COUNTUP(port)
        Call FOCUS(port)
    End If
    
Exit_NextRead:
    Exit Function

Err_NextRead:
    MsgBox Err.Description, , "NextRead"
    Call エラーログ記録(Me.name, "NextRead", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_NextRead
End Function

' 測定回数をカウントアップする(次の測定項目は何回目かを計算)
' port -> 1:A / 2:B / 3:C / 4:D
Private Function COUNTUP(port As Long)
On Error Resume Next

    Dim X As Long   ' 横位置(項目)
    Dim Y As Long   ' 縦位置(測定値)
    Dim alp As String
    Dim koumokuNum As Long
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    Me.P1.SetFocus
    
    Me("回数" & alp) = Nz(Me("回数" & alp), 0) + 1
    
    X = ((Nz(Me("回数" & alp), 0) - 1) Mod koumokuNum) + 1
    Y = Int((Nz(Me("回数" & alp), 0) - 1) / koumokuNum) + 1
    Do Until Y > 5
        Do Until X > koumokuNum
            ' 未完了の項目まで進んだら抜ける
            If KeisokuFLG(port, X) = False And Trim(Nz(SD06DAT(port, X).S001, "")) <> "" Then
                Exit Function
            End If
            
            Me("回数" & alp) = Nz(Me("回数" & alp), 0) + 1
            
            X = X + 1
        Loop
        
        X = 1
        Y = Y + 1
    Loop
End Function

' 測定値の更新前処理(手入力用)
Public Function SOKUTEI_BU() As Boolean
On Error GoTo Err_SOKUTEI_BU
    
    Dim alp As String
    Dim port As Long
    
    SOKUTEI_BU = False
    
    alp = Mid(ActCon, 5, 1)
    
    If alp = "A" Then
        port = 1
    ElseIf alp = "B" Then
        port = 2
    ElseIf alp = "C" Then
        port = 3
    ElseIf alp = "D" Then
        port = 4
    End If
    
    If CHK_DATA(port, Nz(Me(ActCon), 0)) = False Then
        DoCmd.CancelEvent
        Exit Function
    End If
    
    SOKUTEI_BU = True

Exit_SOKUTEI_BU:
    Exit Function

Err_SOKUTEI_BU:
    MsgBox Err.Description, , "SOKUTEI_BU"
    Call エラーログ記録(Me.name, "SOKUTEI_BU", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_SOKUTEI_BU
End Function

' 測定値の更新後処理(手入力用)
Public Function SOKUTEI_AU() As Boolean
On Error GoTo Err_SOKUTEI_AU
    
    Dim alp As String
    Dim port As Long
    
    SOKUTEI_AU = False
    
    alp = Mid(ActCon, 5, 1)
    
    If alp = "A" Then
        port = 1
    ElseIf alp = "B" Then
        port = 2
    ElseIf alp = "C" Then
        port = 3
    ElseIf alp = "D" Then
        port = 4
    End If
    
    Call NextRead(port)
    
    SOKUTEI_AU = True

Exit_SOKUTEI_AU:
    Exit Function

Err_SOKUTEI_AU:
    MsgBox Err.Description, , "SOKUTEI_AU"
    Call エラーログ記録(Me.name, "SOKUTEI_AU", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_SOKUTEI_AU
End Function

' 測定値の更新後処理(前回結果用)
Public Function ZSOKUTEI_BU() As Boolean
On Error GoTo Err_ZSOKUTEI_BU
    
    Dim alp As String
    Dim koumokuNum As Long
    Dim port As Long
    
    ZSOKUTEI_BU = False
    
    alp = Mid(ActCon, 6, 1)
    koumokuNum = Right(ActCon, 1)
    
    If alp = "A" Then
        port = 1
    ElseIf alp = "B" Then
        port = 2
    ElseIf alp = "C" Then
        port = 3
    ElseIf alp = "D" Then
        port = 4
    End If
    
    If Nz(Me("Z項目名" & alp & koumokuNum), "") = "" Then
        Exit Function
    End If
    
    ' (規格 + 公差マイナス) ～ (規格 + 公差プラス)の範囲内か
    If ((ZSD06DAT(port, koumokuNum).D016 + ZSD06DAT(port, koumokuNum).D018) <= Nz(Me(ActCon), 0)) And (Nz(Me(ActCon), 0) <= (ZSD06DAT(port, koumokuNum).D016 + ZSD06DAT(port, koumokuNum).D017)) Then
        Me(ActCon).BackColor = RGB(255, 255, 255)
        Me(ActCon).ForeColor = RGB(0, 0, 0)
        
        Call ZKEISAN(port, koumokuNum)
        
        ZSOKUTEI_BU = True
    Else
        Me(ActCon).ForeColor = RGB(255, 0, 0)
        
        DoCmd.CancelEvent
        Exit Function
    End If

Exit_ZSOKUTEI_BU:
    Exit Function

Err_ZSOKUTEI_BU:
    MsgBox Err.Description, , "ZSOKUTEI_BU"
    Call エラーログ記録(Me.name, "ZSOKUTEI_BU", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_ZSOKUTEI_BU
End Function

Public Function SOKUTEI_ENT() As Boolean
On Error GoTo Err_SOKUTEI_ENT
    
    Dim alp As String
    Dim port As Long
    
    SOKUTEI_ENT = False
    
    Call GETACTIVE
    
    Me(ActCon).BackColor = RGB(0, 255, 255)
    Me(ActCon).ForeColor = RGB(0, 0, 0)
    
    SOKUTEI_ENT = True

Exit_SOKUTEI_ENT:
    Exit Function

Err_SOKUTEI_ENT:
    MsgBox Err.Description, , "SOKUTEI_ENT"
    Call エラーログ記録(Me.name, "SOKUTEI_ENT", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_SOKUTEI_ENT
End Function

Public Function SOKUTEI_EXT() As Boolean
On Error GoTo Err_SOKUTEI_EXT
    
    Dim alp As String
    Dim port As Long
    
    SOKUTEI_EXT = False
    
    Me(ActCon).BackColor = RGB(255, 255, 255)
    Me(ActCon).ForeColor = RGB(0, 0, 0)
    
    SOKUTEI_EXT = True

Exit_SOKUTEI_EXT:
    Exit Function

Err_SOKUTEI_EXT:
    MsgBox Err.Description, , "SOKUTEI_EXT"
    Call エラーログ記録(Me.name, "SOKUTEI_EXT", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_SOKUTEI_EXT
End Function

' 前回結果タブに転記する
' port      -> 1:A / 2:B / 3:C / 4:D
' SkipFLG   -> スキップするデータか
Private Function ZENKAI_SET(port As Long, Optional SkipFLG As Boolean)
On Error GoTo Err_ZENKAI_SET
    
    Dim iPart As Long
    Dim alp As String
    Dim koumokuNum As Long
    Dim i As Long
    
    Dim tmpMAX As Double
    Dim tmpMIN As Double
    Dim tmpAVE As Double
    Dim tmpR As Double
    Dim tmpSTD As Double
    Dim tmpCpk As Double
    Dim tmp規格上限 As Double
    Dim tmp規格下限 As Double
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    Me("Z今回" & alp) = Me("今回" & alp)
    Me("Z製品名" & alp) = Me("製品名" & alp)
    
    For iPart = 1 To koumokuNum
        ' 初期化
        Call CLR_SD06(ZSD06DAT(port, iPart))
        
        Me("Z項目名" & alp & iPart) = Null
        Me("Z規格" & alp & iPart) = Null
        Me("Z公差プラス" & alp & iPart) = Null
        Me("Z公差マイナス" & alp & iPart) = Null
        For i = 1 To 5
            Me("Z測定値" & i & alp & iPart) = Null
            Me("Z測定値" & i & alp & iPart).Enabled = False
            Me("Z測定値" & i & alp & iPart).Locked = False
        Next
        Me("CPK" & alp & iPart) = Null
        
        ' スキップされたデータは転記しない
        If SkipFLG = False Then
            With ZSD06DAT(port, iPart)
                .S001 = SD06DAT(port, iPart).S001
                .S002 = SD06DAT(port, iPart).S002
                .S003 = SD06DAT(port, iPart).S003
                .S004 = SD06DAT(port, iPart).S004
                .L005 = SD06DAT(port, iPart).L005
                .L006 = SD06DAT(port, iPart).L006
                .L007 = SD06DAT(port, iPart).L007
                .L008 = SD06DAT(port, iPart).L008
                .S009 = SD06DAT(port, iPart).S009
                .L010 = SD06DAT(port, iPart).L010
                .S011 = SD06DAT(port, iPart).S011
                .S012 = SD06DAT(port, iPart).S012
                .S013 = SD06DAT(port, iPart).S013
                .S014 = SD06DAT(port, iPart).S014
                .S015 = SD06DAT(port, iPart).S015
                .D016 = SD06DAT(port, iPart).D016
                .D017 = SD06DAT(port, iPart).D017
                .D018 = SD06DAT(port, iPart).D018
                .L019 = SD06DAT(port, iPart).L019
                .D020 = SD06DAT(port, iPart).D020
                .D021 = SD06DAT(port, iPart).D021
                .D022 = SD06DAT(port, iPart).D022
                .D023 = SD06DAT(port, iPart).D023
                .D024 = SD06DAT(port, iPart).D024
                .S025 = SD06DAT(port, iPart).S025
                .S026 = SD06DAT(port, iPart).S026
                .S027 = SD06DAT(port, iPart).S027
                .S028 = SD06DAT(port, iPart).S028
            End With
            
            If Nz(Me("項目名" & alp & iPart), "") <> "" Then
                Me("Z項目名" & alp & iPart) = Me("項目名" & alp & iPart)
                Me("Z規格" & alp & iPart) = Me("規格" & alp & iPart)
                Me("Z公差プラス" & alp & iPart) = Me("公差プラス" & alp & iPart)
                Me("Z公差マイナス" & alp & iPart) = Me("公差マイナス" & alp & iPart)
                For i = 1 To Nz(SD06DAT(port, iPart).L019, 1)
                    Me("Z測定値" & i & alp & iPart) = Me("測定値" & i & alp & iPart)
                    If allowEdit Then
                        Me("Z測定値" & i & alp & iPart).Enabled = True
                        Me("Z測定値" & i & alp & iPart).Locked = False
                    Else
                        Me("Z測定値" & i & alp & iPart).Locked = True
                    End If
                Next
            End If
            
            Me("MSG_" & alp) = Null
            
            If allowEdit Then
                Me("修正登録" & alp).Visible = True
            Else
                Me("修正登録" & alp).Visible = False
            End If
            
            Call ZKEISAN(port, iPart)
        Else
            Me("修正登録" & alp).Visible = False
            
            Me("MSG_" & alp) = "スキップ"
        End If
    Next
    
Exit_ZENKAI_SET:
    Exit Function

Err_ZENKAI_SET:
    MsgBox Err.Description, , "ZENKAI_SET"
    Call エラーログ記録(Me.name, "ZENKAI_SET", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_ZENKAI_SET
End Function

' port  -> 1:A / 2:B / 3:C / 4:D
' iPart -> 1:項目1 / 2:項目2 / 3:項目3 / 4:項目4
Private Function ZKEISAN(port As Long, iPart As Long)
On Error GoTo Err_ZKEISAN
    
    Dim alp As String
    Dim i As Long
    Dim tmpMAX As Double
    Dim tmpMIN As Double
    Dim tmpAVE As Double
    Dim tmpR As Double
    Dim tmpSTD As String
    Dim tmpCpk As String
    Dim tmpCpk1 As Double
    Dim tmpCpk2 As Double
    Dim tmp規格上限 As Double
    Dim tmp規格下限 As Double
    Dim arr As Variant
    
    alp = portInfo1(port)
    
    If Nz(Me("Z項目名" & alp & iPart), "") <> "" Then
        ReDim arr(1 To Nz(ZSD06DAT(port, iPart).L019, 1)) As Double
        For i = 1 To Nz(ZSD06DAT(port, iPart).L019, 1)
            arr(i) = Nz(Me("Z測定値" & i & alp & iPart), 0)
        Next
        
        ' 最大値
        tmpMAX = getMaxValue(arr)
        Me("最大値" & alp & iPart) = tmpMAX
        
        ' 最小値
        tmpMIN = getMinValue(arr)
        Me("最小値" & alp & iPart) = tmpMIN
        
        ' 平均値
        tmpAVE = getAverage(arr)
        Me("平均値" & alp & iPart) = tmpAVE
        
        ' 差
        tmpR = tmpMAX - tmpMIN
        Me("差" & alp & iPart) = tmpR
        
        ' 計測回数が1回の場合は計算しない
        If Nz(ZSD06DAT(port, iPart).L019, 1) = 1 Then
            ' 標準偏差
            tmpSTD = " /  "
            
            ' CPK
            tmpCpk = " /  "
        Else
            ' 標準偏差
            tmpSTD = getStandardDeviation(arr)
            
            ' CPK
            ' 規格値が0の時は表示しない
            If Nz(Me("Z規格" & alp & iPart), 0) = 0 Then
                tmpCpk = " /  "
            Else
                If IsNumeric(Nz(Me("Z規格" & alp & iPart), 0)) Then
                    tmp規格上限 = CDbl(Nz(Me("Z規格" & alp & iPart), 0)) + CDbl(Nz(Me("Z公差プラス" & alp & iPart), 0))
                    tmp規格下限 = CDbl(Nz(Me("Z規格" & alp & iPart), 0)) + CDbl(Nz(Me("Z公差マイナス" & alp & iPart), 0))
                    
                    If tmp規格上限 = tmp規格下限 Or CDbl(tmpSTD) = 0 Then
                        tmpCpk = " /  "
                    Else
                        ' （USL―X）/3×σまたは（X―LSL）/3×σのいずれかの小さい値
                        tmpCpk1 = (tmp規格上限 - CDbl(tmpAVE)) / (3 * CDbl(tmpSTD))
                        tmpCpk2 = (CDbl(tmpAVE) - tmp規格下限) / (3 * CDbl(tmpSTD))
                        
                        If tmpCpk1 <= tmpCpk2 Then
                            tmpCpk = Format(tmpCpk1, "0.000")
                        Else
                            tmpCpk = Format(tmpCpk2, "0.000")
                        End If
                    End If
                Else
                    tmpCpk = "NaN"
                End If
            End If
        End If
        
        Me("標準偏差" & alp & iPart) = tmpSTD
        Me("CPK1" & alp & iPart) = tmpCpk1
        Me("CPK2" & alp & iPart) = tmpCpk2
        Me("CPK" & alp & iPart) = tmpCpk
    End If
    
Exit_ZKEISAN:
    Exit Function

Err_ZKEISAN:
    MsgBox Err.Description, , "ZKEISAN"
    Call エラーログ記録(Me.name, "ZKEISAN", Err.Number, Err.Description, "alp：" & alp & ", iPart：" & iPart & ", M/C：" & Me("Z今回" & alp) & ", 製品CD：" & ZSD06DAT(port, iPart).S012 & ", 製品名：" & Trim(Nz(Me("Z製品名" & alp), "")) & ", 項目名：" & Me("Z項目名" & alp & iPart))
    Resume Exit_ZKEISAN
End Function

Public Function KOUMOKU_DC()
On Error GoTo Err_KOUMOKU_DC
    
    Dim alp As String
    Dim iPart As Long
    Dim port As Long
    Dim MSG As String
    Dim i As Long
    Dim ZENKAIFLG As Boolean
    Dim name As String
    
    If Not ActCon Like "*項目名*" Then
        Exit Function
    End If
    
    If Nz(Me(ActCon), "") = "" Then
        Exit Function
    End If
    
    If Left(ActCon, 1) <> "Z" Then
        ZENKAIFLG = False
        
        name = ActCon
    Else
        ZENKAIFLG = True
        
        ' 最初のZを取らないと番号の計算が狂うため
        name = Mid(ActCon, 2)
    End If
    
    alp = Mid(name, 4, 1)
    iPart = Right(name, 1)
    
    If alp = "A" Then
        port = 1
    ElseIf alp = "B" Then
        port = 2
    ElseIf alp = "C" Then
        port = 3
    ElseIf alp = "D" Then
        port = 4
    End If
    
    MSG = ""
    
    ' 今回
    If ZENKAIFLG = False Then
        ' 製品
        MSG = MSG & "製品：" & SD06DAT(port, iPart).S012 & " " & SD06DAT(port, iPart).S013 & vbNewLine
        
        ' 項目名
        MSG = MSG & "製品：" & SD06DAT(port, iPart).S015 & vbNewLine
        
        ' 計測回数
        MSG = MSG & "計測回数：" & SD06DAT(port, iPart).L019
    Else
        
    End If
    
    MsgBox MSG, vbOKOnly, "項目情報"
    
Exit_KOUMOKU_DC:
    Exit Function

Err_KOUMOKU_DC:
    Resume Exit_KOUMOKU_DC
End Function

Public Function CPK_DC()
On Error GoTo Err_CPK_DC
    
    Dim alp As String
    Dim iPart As Long
    Dim port As Long
    Dim MSG As String
    Dim tmpCpk As Double
    Dim tmpCpk1 As Double
    Dim tmpCpk2 As Double
    Dim tmp規格上限 As Double
    Dim tmp規格下限 As Double
    Dim i As Long
    Dim name As String
    
    name = Replace(ActCon, "_前回", "")
    
    alp = Mid(name, 4, 1)
    iPart = Right(name, 1)
    
    If alp = "A" Then
        port = 1
    ElseIf alp = "B" Then
        port = 2
    ElseIf alp = "C" Then
        port = 3
    ElseIf alp = "D" Then
        port = 4
    End If
    
    If Nz(Me("Z項目名" & alp & iPart), "") = "" Then
        Exit Function
    End If
    
    If IsNumeric(Nz(Me("Z規格" & alp & iPart), 0)) = False Then
        Exit Function
    End If
    tmp規格上限 = CDbl(Nz(Me("Z規格" & alp & iPart), 0)) + CDbl(Nz(Me("Z公差プラス" & alp & iPart), 0))
    tmp規格下限 = CDbl(Nz(Me("Z規格" & alp & iPart), 0)) + CDbl(Nz(Me("Z公差マイナス" & alp & iPart), 0))
    
    MSG = ""
    
    MSG = MSG & "M/C№："
    MSG = MSG & Me("Z今回" & alp) & vbNewLine & vbNewLine
    
    MSG = MSG & "製品名："
    MSG = MSG & Me("Z製品名" & alp) & vbNewLine & vbNewLine
    
    For i = 1 To Nz(ZSD06DAT(port, iPart).L019, 1)
        MSG = MSG & "測定値" & i & "："
        MSG = MSG & Format(Me("Z測定値" & i & alp & iPart), "0.000") & vbNewLine & vbNewLine
    Next
    
    MSG = MSG & "規格上限："
    MSG = MSG & Format(tmp規格上限, "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "規格下限："
    MSG = MSG & Format(tmp規格下限, "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "最大値："
    MSG = MSG & Format(Me("最大値" & alp & iPart), "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "最小値："
    MSG = MSG & Format(Me("最小値" & alp & iPart), "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "平均値："
    MSG = MSG & Format(Me("平均値" & alp & iPart), "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "差："
    MSG = MSG & Format(Me("差" & alp & iPart), "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "標準偏差："
    MSG = MSG & Format(Me("標準偏差" & alp & iPart), "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "CPK1：【(規格上限 - 平均値) / 3 × 標準偏差】" & vbNewLine
    tmpCpk1 = Nz(Me("CPK1" & alp & iPart), 0)
    MSG = MSG & Format(tmpCpk1, "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "CPK2：【(平均値 - 規格下限) / 3 × 標準偏差】" & vbNewLine
    tmpCpk2 = Nz(Me("CPK2" & alp & iPart), 0)
    MSG = MSG & Format(tmpCpk2, "0.000") & vbNewLine & vbNewLine
    
    MSG = MSG & "CPK：【CPK1とCPK2のいずれかの小さい値】" & vbNewLine
    If tmpCpk1 <= tmpCpk2 Then
        tmpCpk = tmpCpk1
    Else
        tmpCpk = tmpCpk2
    End If
    MSG = MSG & Format(tmpCpk, "0.000") & vbNewLine & vbNewLine
    
    MsgBox MSG, vbOKOnly, "前回測定結果"
    
Exit_CPK_DC:
    Exit Function

Err_CPK_DC:
    Resume Exit_CPK_DC
End Function

Private Sub P12_Click()
On Error GoTo Err_P12_Click
    
    If MsgBox("終了してもよろしいですか？", vbYesNo + vbInformation + vbDefaultButton2, "確認") = vbNo Then
        Exit Sub
    End If
    
    Call 閉じる

Exit_P12_Click:
    Exit Sub

Err_P12_Click:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "P12_Click"
    Call エラーログ記録(Me.name, "P12_Click", Err.Number, Err.Description)
    Resume Exit_P12_Click
End Sub

' port      -> 1:A / 2:B / 3:C / 4:D
Private Sub Skip(port As Long)
On Error GoTo Err_Skip
    
    Dim X As Long   ' 横位置(項目)
    Dim Y As Long   ' 縦位置(測定値)
    Dim alp As String
    Dim koumokuNum As Long
    
    alp = portInfo1(port)
    koumokuNum = portInfo2(port)
    
    If Me("測定FLG_" & alp) = False Then
        Exit Sub
    End If
    
    If MsgBox("スキップしてもよろしいですか？", vbYesNo + vbInformation + vbDefaultButton2, "確認") = vbNo Then
        Exit Sub
    End If
    
    Me.P1.SetFocus
    
    X = ((Nz(Me("回数" & alp), 0) - 1) Mod koumokuNum) + 1
    Y = Int((Nz(Me("回数" & alp), 0) - 1) / koumokuNum) + 1
    
    KeisokuFLG(port, X) = True
    
    Call DATAWRITE(port, , True)
    
    ' 前回結果タブに転記
    Call ZENKAI_SET(port, True)
    
    ' 次のデータの読み込み
    Call DATAREAD(port)

Exit_Skip:
    Exit Sub

Err_Skip:
    DoCmd.Hourglass False
    MsgBox Err.Description, , "Skip"
    Call エラーログ記録(Me.name, "Skip", Err.Number, Err.Description, "モード：" & Me.モード)
    Resume Exit_Skip
End Sub

Private Sub スキップA_Click()
    Call Skip(1)
End Sub

Private Sub スキップB_Click()
    Call Skip(2)
End Sub

Private Sub スキップC_Click()
    Call Skip(3)
End Sub

Private Sub 修正登録A_Click()
    If DATAWRITE(1, True) Then
        MsgBox "修正登録しました", vbOKOnly, "確認"
    End If
End Sub

Private Sub 修正登録B_Click()
    If DATAWRITE(2, True) Then
        MsgBox "修正登録しました", vbOKOnly, "確認"
    End If
End Sub

Private Sub 修正登録C_Click()
    If DATAWRITE(3, True) Then
        MsgBox "修正登録しました", vbOKOnly, "確認"
    End If
End Sub

