Attribute VB_Name = "Module1"
' *****************************************************************************
'           FileMakerの勤務日誌から抽出したファイルから
'           超過勤務表を作るツール
' Date      2020/11/12  K.Endo
' In        .txtファイル    : FileMaker業務日誌からExportした実働時間のファイル
' Out       .xlsファイル×n : メンバ分の超勤表Excelファイル
' Memo      1. カンマ区切りのテキストファイルを選択する
'           2. UnicodeのtxtをSJISのcsvに変換する(PowerShell実行)
'           3. csvファイルからインポートする
'           4. 実働時間と勤務形態から超勤時間を求める
'           5. メンバごとに超勤ファイルを作成して4を転記、出力する
'
' History   <#001> Gドライブでファイル名を指定できるように、出力ファイル名に
'                  タイムスタンプを付加するのをやめる
' *****************************************************************************
' -----------------------------------------
' 構造体定義
' -----------------------------------------
' 読み込むtxtファイルの定義　1行分
Type REC_WORK_HOUR
    workDay As Date         ' 日にち
    name As String          ' 氏名
    strTypeOfWork As String ' 勤務形態（文字）
    startTime As Date       ' 実働時間（自）
    endTime As Date         ' 実働時間（至） 24時を超えた場合は-24した数字
    memo As String          ' メモ
End Type

' 書き込むデータの定義＋インポートデータ
Type REC_OVER_WORK
    workDay As Integer      ' 1～31固定
    strTypeOfWork As String ' 勤務形態
    startTime1 As Date      ' 勤務命令時間（上段）自
    endTime1 As Date        ' 勤務命令時間（上段）至
    startTime2 As Date      ' 勤務命令時間（下段）自
    endTime2 As Date        ' 勤務命令時間（下段）至
    memo As String          ' 業務内容
    ' インポートデータ
    startWorkTime As Date   ' 実働時間（自）
    endWorkTime As Date     ' 実働時間（至）
End Type

' 超過勤務時間構造体（上下段あるので2セット）
Type REC_OVER_TIME
    startTime1 As Date      ' 超過時間１（自）
    endTime1 As Date        ' 超過時間１（至）
    startTime2 As Date      ' 超過時間２（自）
    endTime2 As Date        ' 超過時間２（至）
End Type

' -----------------------------------------
' 定数定義
' -----------------------------------------
' 勤務形態(勤務日誌の定義）
Const TYPE_WORK_NIKKIN = "日勤"
Const TYPE_WORK_NITIYA = "日夜"
Const TYPE_WORK_NAGA = "長日"
Const TYPE_WORK_SYUKU = "宿直"
Const TYPE_WORK_HIBAN = "非番"
Const TYPE_WORK_HAYAA = "早A"
Const TYPE_WORK_HAYAB = "早B"
Const TYPE_WORK_DAIKYU = "代休"
Const TYPE_WORK_NENKYU = "年休"
Const TYPE_WORK_TOKKYU = "特休"
Const TYPE_WORK_4 = "4"
Const TYPE_WORK_TYOKIN = "超勤"
Const TYPE_WORK_SYUTTYO = "出張"
Const TYPE_WORK_KIBIKI = "忌引"

' 勤務形態(超勤表の定義）
Const TYPE_OVER_NIKKIN = "日勤"
Const TYPE_OVER_YAKIN = "夜勤"
Const TYPE_OVER_HIBAN = "非番"
Const TYPE_OVER_KYUJITU = "休日"
Const TYPE_OVER_SYUKU = "宿直"
Const TYPE_OVER_DAIKYU = "代休"
Const TYPE_OVER_CLEAR = "特休"  ' 特休、年休は空欄にする（超勤表には定義がない）

' Excelのセル配列（m_aryCell）のindex
Const CELLARY_IDX_YM = 1        ' 年月
Const CELLARY_IDX_BMN = 2       ' 部局
Const CELLARY_IDX_KA = 3        ' 課・室
Const CELLARY_IDX_NAME = 4      ' 氏名
Const CELLARY_IDX_HOLIDAY = 5   ' 祝日
Const CELLARY_IDX_TYPE = 6      ' 勤務形態
Const CELLARY_IDX_YOTEI_ST = 7  ' 勤務予定時間開始
Const CELLARY_IDX_YOTEI_ED = 8  ' 勤務予定時間終了
Const CELLARY_IDX_MEIREI_ST = 9 ' 勤務命令時間１開始
Const CELLARY_IDX_MEIREI_ED = 10 ' 勤務命令時間１終了
Const CELLARY_IDX_MEMO = 11     ' 業務内容

' メンバ一人分の構造体
Type REC_ARY_OVERS
    name As String              ' 名前
    aryOver(31) As REC_OVER_WORK ' 1か月分の超勤表
End Type

' メンバ情報（シートのTメンバテーブルから取得する）
Type MEMBER_LIST
    name As String              ' 氏名
    lastName As String          ' 姓
End Type

' -----------------------------------------
' メンバ変数定義
' -----------------------------------------
                                ' メンバリスト
Public m_lstMember() As MEMBER_LIST
                                ' 勤務実績の構造体配列（ファイルからインポートしたデータ）
Public m_aryWork() As REC_WORK_HOUR
                                ' メンバごとの超勤配列
Public m_aryOver() As REC_ARY_OVERS
                                ' 出力年月
Public m_outDate As Date
                                ' 出力先
Public m_outPath As String
                                ' 超勤表の原本ファイルパス
Public m_strFileName As String
                                ' 超勤表のシート名
Public m_strSheetName As String
                                ' 超勤表Excelファイルのセル位置
Public m_aryCell(11) As String

' *****************************************************************************
' Name      Init
' Param     Nothing
' Result    Nothing
' Memo      初期処理
' *****************************************************************************
Sub Init()
    Dim iCnt As Integer
    Dim iRow As Integer
    Dim iMemberCnt As Integer
    Dim aryStrName() As Variant   ' Arrayを使うのでVariantじゃないとだめ
    
    ' -----------------------------------------
    ' 入力パラメータ
    ' -----------------------------------------
    m_outDate = Sheets("ツール").Range("出力年月").Value
    m_outPath = Sheets("ツール").Range("出力先").Value
    
    ' -----------------------------------------
    ' 超勤表の定義
    ' -----------------------------------------
    m_strFileName = Sheets("ツール").Range("原本ファイルパス").Value
    m_strSheetName = Sheets("ツール").Range("シート名").Value
    
    ' 書き込みセル位置の定義名(VBAでは配列constが定義できない)
    aryStrName = Array("", _
                        "年月", _
                        "部局", _
                        "課・室", _
                        "氏名", _
                        "祝日１", _
                        "勤務形態１", _
                        "勤務予定時間１開始", _
                        "勤務予定時間１終了", _
                        "勤務命令時間１開始", _
                        "勤務命令時間１終了", _
                        "業務内容１")
    
    
    ' 名前のついたセルから値を取り出す
    For iCnt = 1 To UBound(m_aryCell)
        m_aryCell(iCnt) = Sheets("ツール").Range(aryStrName(iCnt)).Value
    Next
    
    
    ' -----------------------------------------
    ' メンバ一覧
    ' -----------------------------------------
                                            ' メンバ数取得(見出しの分を除く)
    iMemberCnt = Sheets("メンバ一覧").Range("T氏名").Rows.Count
    
    ReDim m_lstMember(iMemberCnt - 1)       ' メンバ分配列を拡張
    
    ' メンバリストをシートから読み込む
    For iCnt = 0 To iMemberCnt - 1
                                            ' 氏名
        m_lstMember(iCnt).name = Sheets("メンバ一覧").Range("T氏名").Cells(iCnt + 1, 2) ' +1は見出しの分
                                            ' 姓
        m_lstMember(iCnt).lastName = Sheets("メンバ一覧").Range("T氏名").Cells(iCnt + 1, 3)
    Next
    
    ReDim m_aryOver(iMemberCnt - 1)         ' 配列をメンバの人数分拡張
    
    
End Sub

' *****************************************************************************
' Name      ImportFile
' Param     Nothing
' Result    Boolean : T or F
' Memo      txtファイルを読み込んで勤務実績構造体につめる
' *****************************************************************************
Function ImportFile() As Boolean
On Error GoTo ErrHandler
    Dim iCnt As Integer
    Dim ofn As String
    Dim iFileNumber As Integer
    Dim strDay As String
    Dim strStart As String
    Dim strEnd As String
    Dim iHour As Integer
    
    ImportFile = True
    
    iCnt = 0
    iFileNumber = FreeFile
    
    ' txtファイルを選択
    ofn = Application.GetOpenFilename(FileFilter:="テキストファイル (*.txt), *.txt", _
                                    Title:="インポートファイル選択")
'    ofn = Application.GetOpenFilename(FileFilter:="CSVファイル (*.csv), *.csv", _
'                                    Title:="インポートファイル選択")

    ' 戻りはバリアント型。正しく開ければファイルパスが返ってくる
    If ofn <> "False" Then
        ' いったんPowerShellでエンコード。ここでcsvになる
        strOutPath = UnicodeToSJIS(ofn)
        ' 変換後のcsvファイルオープン
        Open strOutPath For Input As #iFileNumber
        'Open ofn For Input As #iFileNumber
    Else
        ImportFile = False
        Exit Function
    End If
    
    ' csvファイルを読み込み
    Do While Not EOF(iFileNumber)
        iCnt = iCnt + 1
        ReDim Preserve m_aryWork(iCnt)      ' 配列拡張
                                            ' 勤務実績1行分を配列にセット
        Input #iFileNumber, strDay, _
                            m_aryWork(iCnt - 1).name, _
                            m_aryWork(iCnt - 1).strTypeOfWork, _
                            strStart, _
                            strEnd, _
                            m_aryWork(iCnt - 1).memo
                            
        ' 日付、時間はいったん変数に入れてから型変換する
        If IsDate(strDay) Then
            m_aryWork(iCnt - 1).workDay = CDate(strDay)
        End If
        ' 実働時間（自）
        If IsDate(strStart) Then
            m_aryWork(iCnt - 1).startTime = CDate(strStart)
        End If
        ' 実働時間（至）
        If IsDate(strEnd) Then
            m_aryWork(iCnt - 1).endTime = CDate(strEnd)
        ElseIf strEnd <> "" Then
            ' 24時を超えた表記の場合、終わり時間を23:59にして、GetOverTimeでまるめる
            strEnd = "23:59:00"
            m_aryWork(iCnt - 1).endTime = CDate(strEnd)
        End If
        
    Loop
    
    GoTo Finally
        
ErrHandler:
    ImportFile = False
                                            ' エラーがあったらメッセージ
    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
    Resume Finally
    
Finally:
    If ofn <> "False" Then
        Close #iFileNumber
    End If

End Function

' *****************************************************************************
' Name      SetOvers
' Param     Nothing
' Result    Nothing
' Memo     インポートしたデータから超勤構造体配列を作成する（配列[メンバ][1..31]）につめる
' *****************************************************************************
Sub SetOvers()
    Dim iCnt As Integer
    Dim iDay As Integer
    Dim iMember As Integer
    Dim strType As String
    Dim startDefTime As Date
    Dim endDefTime As Date
    Dim recOver As REC_OVER_TIME

    For iCnt = LBound(m_aryWork) To UBound(m_aryWork)
        ' 日にち
        iDay = Day(m_aryWork(iCnt).workDay)
        ' 名前からメンバリストのindex取得
        iMember = GetMemberIdx(m_aryWork(iCnt).name)
        ' リストにない人は飛ばす
        If iMember < 0 Then
            GoTo continue
        End If
        
        With m_aryOver(iMember).aryOver(iDay)
            ' コピーする項目
                                            ' 日にち（1-31のdayのみ）
            .workDay = iDay
                                            ' 実働時間（自）
            .startWorkTime = m_aryWork(iCnt).startTime
                                            ' 実働時間（至）
            .endWorkTime = m_aryWork(iCnt).endTime
                                            ' メモ
            .memo = m_aryWork(iCnt).memo
            
        
            ' 計算する項目
                                            ' 勤務形態（FM用から超勤表用にコンバート）
            strType = GetOutTypeOfWork(m_aryWork(iCnt).strTypeOfWork)
            .strTypeOfWork = strType
                                            ' デフォルト勤務時間
            Call GetDefaultTime(m_aryWork(iCnt).strTypeOfWork, startDefTime, endDefTime)
                                            ' 超勤時間
                                            ' デフォ時間と実働時間から超勤時間を設定
            Call GetOverTime(startDefTime, endDefTime, .startWorkTime, .endWorkTime, recOver)
            
            .startTime1 = recOver.startTime1
            .endTime1 = recOver.endTime1
            .startTime2 = recOver.startTime2
            .endTime2 = recOver.endTime2
        End With
        
continue:
    Next

End Sub
' *****************************************************************************
' Name      GetMemberIdx
' Param     strName     : 名前
' Result    Integer     : 名前が入っている配列のindex
' Memo     メンバの名前が入っている配列のindexを返す
'           -1: 見つからなかった場合
' *****************************************************************************
Function GetMemberIdx(strName As String) As Integer
    Dim iCnt As Integer
    
    GetMemberIdx = -1
    
    For iCnt = LBound(m_lstMember) To UBound(m_lstMember)
        If m_lstMember(iCnt).name = strName Then
            GetMemberIdx = iCnt
            Exit Function
        End If
    Next
End Function

' *****************************************************************************
' Name      CreateOverFiles
' Param     Nothing
' Result    Nothing
' Memo     超勤構造体配列から各メンバごとのExcelファイルを出力する
' *****************************************************************************
Sub CreateOverFiles()
On Error GoTo ErrHandler
    Dim iCnt As Integer
    Dim strFileName As String
    Dim strDate As String
    Dim strTimeStamp As String
    Dim wb As Workbook
    Dim objBook As Workbook
    Dim objFileSystem As New Scripting.FileSystemObject
       
    ' 指定した出力年月
    strDate = Format(m_outDate, "yyyymm")
    ' 出力日時
    ' <#001> DEL
    'strTimeStamp = Format(Now, "yyyymmddhhnnss")
    
    ' 出力先の存在チェック
    If Not objFileSystem.FolderExists(m_outPath) Then
        MsgBox "出力先パスが存在しません。正しく入力してください。"
        Exit Sub
    End If
    
    ' 原本ファイルを人数分コピー
    For iCnt = LBound(m_lstMember) To UBound(m_lstMember)
        ' <#001> MOD start
        ' ex) 「名前202011_20201127xxxxxx.xls」
        'strFileName = m_outPath & m_lstMember(iCnt).lastName & strDate & "_" & strTimeStamp & ".xlsx"
        
        ' ex) 「名前202011.xls」Gドライブでファイル名を指定するためにタイムスタンプの付加をやめる
        strFileName = m_outPath & m_lstMember(iCnt).lastName & strDate & ".xlsx"
        ' <#001> MOD end
        
        ' 原本ファイルの存在チェック
        If Dir(m_strFileName) = "" Then
            MsgBox "超勤表原本のファイルパスを確認してください。"
            Exit Sub
        End If
        
        ' コピー
        FileCopy m_strFileName, strFileName
        
        ' ファイルオープン
        If Dir(strFileName) <> "" Then
            Set objBook = Workbooks.Open(strFileName)
        Else
            MsgBox "ファイル 【" & strFileName & "】 が存在しません。", vbExclamation
            Exit Sub
        End If
      
        ' ファイルの書き込み
        If Not WriteOutFile(iCnt, objBook) Then
            MsgBox "シート 【" & m_strSheetName & "】 が存在しません。", vbExclamation
            GoTo Finally
        End If
        
        ' 保存
        objBook.Save
        
        ' 閉じる
        objBook.Close
    Next

    GoTo Finally
    
ErrHandler:
                                            ' エラーがあったらメッセージ
    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
    Resume Finally
    
Finally:
    ' 最後に処理していたブックが閉じていなかったら閉じる
    For Each wb In Workbooks
        If wb.name = Dir(strFileName) Then
            wb.Close
        End If
    Next
End Sub

' *****************************************************************************
' Name      WriteOutFile
' Param     iMemberIdx  : m_lstMemberのindex
'           objBook     : 書き込むExcelBook
' Result    Boolean     : T: 正常 or F: 異常
' Memo      ファイルに書き込むメイン処理（書き込むファイルはOpen済み）
' *****************************************************************************
Function WriteOutFile(iMemberIdx As Integer, objBook As Workbook) As Boolean
On Error GoTo ErrHandler
    Dim iDay As Integer
    Dim iNextPageOffset As Integer
    Dim iOffset As Integer
    Dim iOffset2 As Integer
    Dim objSheet As Worksheet
    
    WriteOutFile = True
    
    ' シートがなかったら何もしない
    If Not SheetExists(m_strSheetName, objBook) Then
        WriteOutFile = False
        Exit Function
    End If

    Set objSheet = objBook.Sheets(m_strSheetName)
    
    ' 年月
    objSheet.Range(m_aryCell(CELLARY_IDX_YM)).Value = m_outDate
    ' ★部局
    'objSheet.Range(m_aryCell(CELLARY_IDX_BMN)).Value =
    ' ★課・室
    'objSheet.Range(m_aryCell(CELLARY_IDX_KA)).Value =
    ' 氏名
    objSheet.Range(m_aryCell(CELLARY_IDX_NAME)).Value = m_lstMember(iMemberIdx).name
    
    ' for 1日～31日ループ
    For iDay = 1 To 31
        ' 2ページ目は先頭位置を変更
        If iDay > 15 Then
            iNextPageOffset = 10
        Else
            iNextPageOffset = 0
        End If
        
        ' 1日につき2行
        iOffset = iNextPageOffset + ((iDay - 1) * 2 - 1)
        iOffset2 = iNextPageOffset + iDay * 2 - 2
        
        ' 1日目だけはOffsetを指定
        If iOffset < 0 Then
            iOffset = 0
        End If
        
        With m_aryOver(iMemberIdx).aryOver(iDay)
            ' 月末日をこえたら
            If Month(DateSerial(Year(m_outDate), Month(m_outDate), iDay)) <> Month(m_outDate) Then
                ' 勤務形態をクリア
                objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = ""
                GoTo continue
            End If
            ' 祝日
            objSheet.Range(m_aryCell(CELLARY_IDX_HOLIDAY)).Offset(iOffset, 0).Value = GetHoliday(DateSerial(Year(m_outDate), Month(m_outDate), iDay))
            
            ' 勤務形態
            Select Case .strTypeOfWork
                Case TYPE_OVER_CLEAR
                    ' 特休・年休の場合はクリア
                    objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = ""
                Case ""
                    ' 未設定の場合は休日
                    objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = TYPE_OVER_KYUJITU
                Case Else
                    ' その他
                    objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = .strTypeOfWork
            End Select
                                      
            ' ---------上段---------
            If (.startTime1 <> "0:00") Or (.endTime1 <> "0:00") Then
                ' *** 勤務予定時間は数式で勤務命令時間をコピーする ***
                ' 勤務予定時間　開始
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ST)).Offset(iOffset2, 0).Value = .startTime1
                ' 勤務予定時間　終了
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ED)).Offset(iOffset2, 0).Value = .endTime1
                
                ' 勤務命令時間　開始
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ST)).Offset(iOffset2, 0).Value = GetRoundTime(.startTime1)
                ' 勤務命令時間　終了
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ED)).Offset(iOffset2, 0).Value = GetRoundTime(.endTime1)
            End If
            
            ' ---------下段---------
            If (.startTime2 <> "0:00") Or (.endTime2 <> "0:00") Then
                ' *** 勤務予定時間は数式で勤務命令時間をコピーする ***
                ' 勤務予定時間　開始
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ST)).Offset(iOffset2 + 1, 0).Value = .startTime2
                ' 勤務予定時間　終了
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ED)).Offset(iOffset2 + 1, 0).Value = .endTime2
                
                ' 勤務命令時間　開始
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ST)).Offset(iOffset2 + 1, 0).Value = GetRoundTime(.startTime2)
                ' 勤務命令時間　終了
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ED)).Offset(iOffset2 + 1, 0).Value = GetRoundTime(.endTime2)
            End If
            
            ' 業務内容
            objSheet.Range(m_aryCell(CELLARY_IDX_MEMO)).Offset(iOffset2, 0).Value = .memo

        End With
continue:
    Next

    Exit Function
ErrHandler:
                                            ' エラーがあったらメッセージ
    MsgBox "【WriteOutFile】 " & Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Function

' *****************************************************************************
' Name      GetOutTypeOfWork
' Param     strType :   勤務形態(業務日誌）
' Result    Integer :   勤務形態(超勤表の定義）
' Memo     業務日誌の勤務形態（文字列）から超勤表の勤務形態（文字列）を返す
' *****************************************************************************
Function GetOutTypeOfWork(strType As String) As String
    Dim strRet As String
    
    ' FMの勤務形態から超勤表の定義を返す
    Select Case strType
        Case TYPE_WORK_NIKKIN, _
            TYPE_WORK_HAYAA, _
            TYPE_WORK_HAYAB, _
            TYPE_WORK_4, _
            TYPE_WORK_SYUTTYO
            strRet = TYPE_OVER_NIKKIN
        Case TYPE_WORK_HIBAN
            strRet = TYPE_OVER_HIBAN
        Case TYPE_WORK_NITIYA, _
            TYPE_WORK_NAGA
            strRet = TYPE_OVER_YAKIN
        Case TYPE_WORK_SYUKU
            strRet = TYPE_OVER_SYUKU
        Case TYPE_WORK_DAIKYU
            strRet = TYPE_OVER_DAIKYU
        Case TYPE_WORK_TYOKIN, _
            TYPE_WORK_KIBIKI
            strRet = TYPE_OVER_KYUJITU
        Case TYPE_WORK_NENKYU, _
            TYPE_WORK_TOKKYU
            ' 特休、年休は空にする
            strRet = TYPE_OVER_CLEAR
        Case Else
            strRet = ""
    End Select
    
    GetOutTypeOfWork = strRet
    
End Function

' *****************************************************************************
' Name      GetDefaultTime
' Param     strType     :   勤務形態(FMの定義）
'           startTime   :   デフォルト勤務時間（自）out
'           endTime     :   デフォルト勤務時間（至）out
' Result    Nothing
' Memo     勤務形態（超勤表）からデフォルト勤務時間を返す
' *****************************************************************************
Sub GetDefaultTime(strType As String, ByRef startTime As Date, ByRef endTime As Date)
    
    Select Case strType
        ' 日勤、出張
        Case TYPE_WORK_NIKKIN, _
            TYPE_WORK_SYUTTYO
            startTime = TimeValue("8:30")
            endTime = TimeValue("17:15")
        ' 日夜、長日
        Case TYPE_WORK_NITIYA, _
            TYPE_WORK_NAGA
            startTime = TimeValue("8:30")
            endTime = TimeValue("21:30")
        ' 非番、4
        Case TYPE_WORK_HIBAN, _
            TYPE_WORK_4
            startTime = TimeValue("8:30")
            endTime = TimeValue("12:30")
        ' 早A
        Case TYPE_WORK_HAYAA
            startTime = TimeValue("7:45")
            endTime = TimeValue("16:30")
        ' 早B
        Case TYPE_WORK_HAYAB
            startTime = TimeValue("8:00")
            endTime = TimeValue("16:45")
        ' 代休、宿直、年休、特休、超勤、忌引
        Case TYPE_WORK_DAIKYU, _
            TYPE_WORK_SYUKU, _
            TYPE_WORK_NENKYU, _
            TYPE_WORK_TOKKYU, _
            TYPE_WORK_TYOKIN, _
            TYPE_WORK_KIBIKI
            startTime = TimeValue("0:00")
            endTime = TimeValue("0:00")
        ' その他
        Case Else
            startTime = TimeValue("0:00")
            endTime = TimeValue("0:00")
    End Select
End Sub

' *****************************************************************************
' Name      GetOverTime
' Param     defStartTime   :   デフォルト勤務時間（自）
'           defEndTime     :   デフォルト勤務時間（至）
'           workStartTime  :   実働勤務時間（自）
'           workEndTime    :   実働勤務時間（至）
'           recOver        :   超過勤務時間構造体　out
' Result    Nothing
' Memo     デフォルト勤務時間と実働時間から超過勤務時間を返す
' *****************************************************************************
Sub GetOverTime(defStartTime, defEndTime, _
                workStartTime, workEndTime As Date, _
                ByRef recOver As REC_OVER_TIME)
    ' 実働時間なしの場合は00:00を埋めて返す
    If ((Not IsDate(workStartTime)) And (Not IsDate(workEndTime))) Or _
        ((workStartTime = "0:00:00") And (workEndTime = "0:00:00")) Then
        recOver.startTime1 = "00:00"
        recOver.endTime1 = "00:00"
        recOver.startTime2 = "00:00"
        recOver.endTime2 = "00:00"
        Exit Sub
    End If
        
    ' 超過勤務１
    If defStartTime <= workStartTime Then
        recOver.startTime1 = "00:00"
        recOver.endTime1 = "00:00"
    Else
        recOver.startTime1 = workStartTime
        recOver.endTime1 = defStartTime
    End If
           
    ' 超過勤務２
    If defEndTime >= workEndTime Then
        recOver.startTime2 = "00:00"
        recOver.endTime2 = "00:00"
    Else
        recOver.startTime2 = defEndTime
        recOver.endTime2 = workEndTime
    End If

    ' 超過勤務１がなくて２だけの場合は１（上段）に２の内容をつめる
    If (recOver.startTime1 = "00:00") And (recOver.startTime2 <> "00:00") Then
        recOver.startTime1 = recOver.startTime2
        recOver.endTime1 = recOver.endTime2
        recOver.startTime2 = "00:00"
        recOver.endTime2 = "00:00"
    End If
    
End Sub

' *****************************************************************************
' Name      GetHoliday
' Param     dt      : 日付
' Result    String  : "祝日", "休日", ""
' Memo     対象の日付が祝日か休日か、それ以外かを返す
' *****************************************************************************
Function GetHoliday(dt As Date) As String
    Dim str As String
    
        str = 日本の休み(dt)
        
        Select Case str
            Case "祝", "休", "振"
                GetHoliday = "祝日"
            Case Else
                GetHoliday = ""
        End Select
            
            
End Function

' *****************************************************************************
' Name      GetRoundTime
' Param     dt      : 時間
' Result    String  : 時間を文字列にしたもの（23:59だけは24:00にする）
' Memo      23:59を24:00にして返す
' *****************************************************************************
Function GetRoundTime(dt As Date) As String
    If dt = "23:59:00" Then
        GetRoundTime = "24:00:00"
    Else
        GetRoundTime = dt
    End If
End Function

' *****************************************************************************
' Name      SheetExists
' Param     SheetName   : シート名
'           wb          : Workbook
' Result    Boolean     : T: あり F: なし
' Memo      シートの存在チェック
' *****************************************************************************
Function SheetExists(SheetName As String, Optional wb As Excel.Workbook)
   Dim s As Excel.Worksheet
   If wb Is Nothing Then Set wb = ThisWorkbook
   On Error Resume Next
   Set s = wb.Sheets(SheetName)
   On Error GoTo 0
   SheetExists = Not s Is Nothing
End Function

' *****************************************************************************
' Name      UnicodeToSJIS
' Param     strInPath   : txtファイルのフルパス
' Result    String      : csvファイルのフルパス（拡張子が変わっただけ）
' Memo      Unicode(UTF16,BOMつき)のtxtファイルからSJISのcsvを作成する
' *****************************************************************************
Function UnicodeToSJIS(strInPath As String) As String
    Dim strOutPath As String
    Dim strCmd As String

                                            ' txtファイルパス
    strOutPath = Replace(strInPath, ".txt", ".csv")
                                            ' PowerShellのコマンド
    strCmd = "get-content -Encoding Unicode " & strInPath & "| Set-Content " & strOutPath
    
    Call RunPowerShell(strCmd, 0, True)     ' 実行
    
    UnicodeToSJIS = strOutPath
End Function

' *****************************************************************************
' Name      RunPowerShell
' Param     psCmd   : PowerShellコマンド(PowerShell仕様により260文字以内)
'           intVsbl : PowerShell画面表示(1)/非表示(0)
'           waitFlg : PowerShell実行終了を待つ(True)/待たない(False)
' Result
' Memo      PowerShellを実行する
' *****************************************************************************
Function RunPowerShell(psCmd As String, intVsbl As Integer, waitFlg As Boolean)
    Dim objWSH As Object
    Set objWSH = CreateObject("WScript.Shell")

    objWSH.Run "powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & psCmd, intVsbl, waitFlg

End Function
' *****************************************************************************
' 以下githubから
' https://gist.github.com/ooltcloud/30644a3ce497319e5d50
' *****************************************************************************
'------------------------------------------------------------------
' Excel VBA (標準モジュール)
' 日付を指定すると休みの属性に応じて、祝/休/振/日/土、を戻します。平時は空文字。
'------------------------------------------------------------------
' 諸注意
' 　2019年,2020年に対応しました。
' 　①春分の日と秋分の日は計算するのが面倒なので、Wikipedia から予想日をコピペしています。（汗
' 　　前年に官報で発表される日付と違っていた場合は､適宜修正してください｡
' 　②春分の日と秋分の日は 2000 年から設定していますが、他の祝日は 2013 年以前の状況を反映しません。(例:みどりの日)
' 　　あくまで今年 2014 年以降の表示用のプログラムです。過去の祝日を正確に反映したい場合は、修正が必要です。
' 　③祝日法が改正されたら､適宜修正してください｡ (未来は予想できません)
' 　④おまけで「非公式な休日」を足しています。会社とかの休みを反映したい場合は場合はここへ記述します。
' 　　以下のプログラムだとサンプルで盂蘭盆の日程とかが入っています｡不要な場合は当該ルーチン､および呼び出しを削除してください｡
' 　　また「世間は祝日だがうちの会社は休みじゃない」など、祝日を打ち消すロジックは用意していません（汗
'------------------------------------------------------------------
Public Function 日本の休み(vInDate As Date) As String

    Dim ret As String
    
    ' 祝日判定
    ret = 日本の祝日(vInDate)
    If ret <> "" Then
        日本の休み = "祝"
        Exit Function
    End If
        
    ' 国民の休日判定
    If is国民の休日(vInDate) = True Then
        日本の休み = "休"
        Exit Function
    End If
    
    ' 振替休日判定
    If is振替休日(vInDate) = True Then
        日本の休み = "振"
        Exit Function
    End If

    ' 非公式の休日 (盆とか県民の日とか会社の休日) / 必要に応じて
    ret = 非公式の休日(vInDate)
    If ret <> "" Then
        日本の休み = ret
        Exit Function

    End If

    ' 土日判定
    Select Case (Weekday(vInDate))
        Case 1: 日本の休み = "日"
        Case 7: 日本の休み = "土"
        Case Else: 日本の休み = ""
    End Select

End Function


'------------------------------------------------------------------
Private Function 日本の祝日(vInDate As Date) As String

    Dim strRet As String
    strRet = ""
    
    ' 2000年から2030年のまでの春分/秋分の一覧表
    ' (Wkipediaより http://ja.wikipedia.org/wiki/%E7%A7%8B%E5%88%86%E3%81%AE%E6%97%A5)
    Dim 春分 As Variant
    春分 = Array(20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 20, 21, 20, 20, 20)
    
    Dim 秋分 As Variant
    秋分 = Array(23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23)
    
    ' 要素分解
    Dim intyear As Integer:     intyear = Year(vInDate)
    Dim intday As Integer:      intday = Day(vInDate)
    Dim intMonth As Integer:    intMonth = Month(vInDate)
    Dim intWeekDay As Integer:  intWeekDay = Weekday(vInDate)
    Dim intWeek As Integer:     intWeek = (intday - 1) \ 7 + 1
    
    ' 月日固定
    If (intMonth = 1) And (intday = 1) Then strRet = "元日"
    If (intMonth = 2) And (intday = 11) Then strRet = "建国記念の日"
    If (intMonth = 2) And (intday = 23) And (intyear >= 2020) Then strRet = "天皇誕生日"
    If (intMonth = 4) And (intday = 29) Then strRet = "昭和の日"
    If (intMonth = 5) And (intday = 3) Then strRet = "憲法記念日"
    If (intMonth = 5) And (intday = 4) Then strRet = "みどりの日"
    If (intMonth = 5) And (intday = 5) Then strRet = "こどもの日"
    If (intMonth = 8) And (intday = 11) And (intyear >= 2016) And (intyear <> 2020) Then strRet = "山の日"
    If (intMonth = 11) And (intday = 3) Then strRet = "文化の日"
    If (intMonth = 11) And (intday = 23) Then strRet = "勤労感謝の日"
    If (intMonth = 12) And (intday = 23) And (intyear <= 2018) Then strRet = "天皇誕生日"
    
    ' 春分/秋分
    If (2000 <= intyear And intyear <= 2030) Then
        If (intMonth = 3 And intday = 春分(intyear - 2000)) Then strRet = "春分の日"
        If (intMonth = 9 And intday = 秋分(intyear - 2000)) Then strRet = "秋分の日"
    End If
    
    ' ハッピーマンデー(特定週の月曜日固定)
    If (intWeekDay = 2) Then    ' 月曜
        If (intMonth = 1) And (intWeek = 2) Then strRet = "成人の日"
        If (intMonth = 7) And (intWeek = 3) And (intyear <> 2020) Then strRet = "海の日"
        If (intMonth = 9) And (intWeek = 3) Then strRet = "敬老の日"
        If (intMonth = 10) And (intWeek = 2) And (intyear <= 2019) Then strRet = "体育の日"
        If (intMonth = 10) And (intWeek = 2) And (intyear >= 2021) Then strRet = "スポーツの日"
        
    End If
    
    ' 2019年特例
    If (intyear = 2019) Then
        If (intMonth = 5) And (intday = 1) Then strRet = "天皇の即位の日"
        If (intMonth = 10) And (intday = 22) Then strRet = "即位礼正殿の儀の行われる日"
    End If

    ' 2020年特例
    If (intyear = 2020) Then
        If (intMonth = 7) And (intday = 23) Then strRet = "海の日"
        If (intMonth = 7) And (intday = 24) Then strRet = "スポーツの日"
        If (intMonth = 8) And (intday = 10) Then strRet = "山の日"
    End If
    
    日本の祝日 = strRet

End Function


'------------------------------------------------------------------
Private Function is振替休日(vInDate As Date) As Boolean

    ' 自身が祝日であれば、振替ではない
    If 日本の祝日(vInDate) <> "" Then
        is振替休日 = False
        Exit Function
    End If

    ' 自身が祝日でなければ、祝日が日曜まで連続しているかを確認
    For n = 1 To 7
        Dim d As Date
        d = DateAdd("d", -n, vInDate)   ' n日前
        
        ' 日曜に至るまでに祝日が途切れれば振替でない
        If (日本の祝日(d) = "") Then
            is振替休日 = False
            Exit Function
            
        Else
            If (Weekday(d) = 1) Then
                ' 祝日であり日曜
                is振替休日 = True
                Exit Function
            
            Else
                ' LOOP(さらに前日を確認)
            
            End If
        End If
    Next

End Function


'------------------------------------------------------------------
Private Function is国民の休日(vInDate As Date) As Boolean

    ' 自身が祝日/日曜であれば、国民の休日ではない
    If (日本の祝日(vInDate) <> "") Or (Weekday(vInDate) = 1) Then
        is国民の休日 = False
        Exit Function

    End If

    If 日本の祝日(DateAdd("d", -1, vInDate)) <> "" And _
        日本の祝日(DateAdd("d", 1, vInDate)) <> "" Then
    
        is国民の休日 = True
        Exit Function

    End If
    
    is国民の休日 = False

End Function


'------------------------------------------------------------------
Private Function 非公式の休日(vInDate As Date) As String
    
    Dim strRet As String
    strRet = ""

    Dim intday As Integer:      intday = Day(vInDate)
    Dim intMonth As Integer:    intMonth = Month(vInDate)

    ' 盂蘭盆を 8/13～15 と仮定した場合
    'If (intMonth = 8) And (13 <= intday And intday <= 15) Then strRet = "盆"

    ' 群馬県民の日
    'If (intMonth = 10) And (intday = 28) Then strRet = "群"
    
    非公式の休日 = strRet

End Function



