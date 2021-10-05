' *****************************************************************************
'           定期研修の参加状況一覧シートに新規シートを追加するツール
' Date      2020/10/20  K.Endo
' Memo      年度が固定になっているので後でメンテ
'           2021/6/14 システム日付から年度を取得するように修正
' *****************************************************************************
' -------------------------------
'   定数
' -------------------------------
' 対象者シートのタイトル（月初）
Const TITLE_FIRST As String = "令和@year年度　医療機器定期研修　研修対象者一覧（@_@現在）"
' 実績シートのタイトル（月末）
Const TITLE_LAST As String = "令和@year年度　医療機器定期研修　受講状況一覧（@_@現在）"

' 置換文字列
Const REPLACE_STRING = "@_@"
Const REPLACE_YEAR = "@year"

' *****************************************************************************
' Control   Closeボタン
' Event     ボタンクリック
' *****************************************************************************
Private Sub btnCancel_Click()
    ' ダイアログを閉じる
    Unload frmSheetAdd
End Sub

' *****************************************************************************
' Control   ファイルを開くボタン
' Event     ボタンクリック
' *****************************************************************************
Private Sub btnFileOpen_Click()
    Dim strFileName As String
    
    ' カレントディレクトリをこのブックのパスにする
    ChDir ThisWorkbook.path
    
    ' ファイル選択ダイアログ呼び出し
    strFileName = Application.GetOpenFilename("ブック, *.xls?")
    If strFileName <> "False" Then
        txtFilePath.Text = strFileName
    End If
 
End Sub

' *****************************************************************************
' Control   OKボタン
' Event     ボタンクリック
' *****************************************************************************
Private Sub btnOK_Click()
    ' メイン処理実行
    Call Main
End Sub

' *****************************************************************************
' Control   フォーム
' Event     Initialize
' *****************************************************************************
Private Sub UserForm_Initialize()
    
    Dim iCnt
    
    ' コンボボックスに月リスト追加
    For iCnt = 1 To 12
        cmbMonth.AddItem CStr(iCnt) + "月"
    Next iCnt
    
    ' 初期値は先頭
    cmbMonth.ListIndex = 0
    
    ' 対象者ラジオボタンをON
    bTarget.Value = True

End Sub

' *****************************************************************************
' Name  Main
' Param  Nothing
' Memo  メイン処理
' *****************************************************************************
Private Sub Main()
    Dim strFileName As String
    Dim strPath As String
    Dim trgBook As Workbook
    Dim iIndex As Integer
    
    On Error GoTo ErrorHandler
    ' 対象ファイルを開く
    strPath = txtFilePath.Text
    If Dir(strPath) <> "" Then
        Set trgBook = Workbooks.Open(strPath)
    Else
        MsgBox "ファイルが存在しません"
    End If
    
    ' ファイルのフルパスからファイル名を取り出しておく（閉じるとき用）
    strFileName = GetFileName(strPath)
        
    ' シートを追加する
    iIndex = AddSheet(trgBook)
    
    If iIndex > 0 Then
        ' シートのタイトルや行列の変更
        Call UpdateSheet(iIndex, trgBook)
        ' 変更を保存
        trgBook.Save
    
        MsgBox "シートを追加しました"
    End If
    
    GoTo Finally
ErrorHandler:
    MsgBox "[No:" & Err.Number & "]" & Err.Description, vbCritical & vbOKOnly, "エラー"
    
Finally:
    ' ファイルが開かれていたら閉じる
    Call WorkBookClose(strFileName)

End Sub

' *****************************************************************************
' Name      AddSheet
' Param     wb      : ブック
' Result    Integer : 追加したシートのindex
' Memo      シート追加処理
' *****************************************************************************
Private Function AddSheet(wb As Workbook) As Integer

    Dim iSheetCount As Integer
    Dim iSheetIndex As Integer
    Dim iNewIndex As Integer
    
    Dim strSheetName As String
    Dim strNewName As String
    Dim strRet As String
    
    AddSheet = 0
    
    ' コピー元のシートindex取得（count = 最後のシートindex）
    iSheetIndex = wb.Sheets.Count
    
    ' 最後のシートから見ていって、関係ないシートを読み飛ばす
    Do While iSheetIndex >= 0
        ' 非表示のシートは読み飛ばす
        If (Not wb.Sheets(iSheetIndex).Visible) Or _
            (wb.Sheets(iSheetIndex).Name = "所属コ－ド") Or _
            (wb.Sheets(iSheetIndex).Name = "所属マスタ") Then
            iSheetIndex = iSheetIndex - 1
        Else
            Exit Do
        End If
    Loop
            
    iNewIndex = iSheetIndex + 1
       
    ' シート名取得
    strSheetName = wb.Sheets(iSheetIndex).Name
    
    ' 対象者or実績の文字位置を検索(1オリジン)
    iRet = InStr(strSheetName, "対象者")
    If iRet = 0 Then
        iRet = InStr(strSheetName, "実績")
    End If
        
    If iRet = 0 Then
        Exit Function
    End If
    
    ' シートをコピーする
    Call wb.Sheets(iSheetIndex).Copy(Before:=Sheets(iNewIndex))
    
    ' 「人工心肺(」まで取り出す
    strNewName = Left(strSheetName, iRet - 1)
    
    ' 「対象者」or「実績」を足す
    If bTarget Then
        strNewName = strNewName + "対象者"
    Else
        strNewName = strNewName + "実績"
    End If
    
    ' 「N月)」を足す
    strNewName = strNewName + cmbMonth.Text + ")"
        
    ' 既に使われている名前の場合、うしろに「_2」をつけておく
    If IsUsedName(wb, strNewName) Then
        strNewName = strNewName & "_2"
    End If
        
    wb.Sheets(iNewIndex).Name = strNewName
   
    AddSheet = iNewIndex
End Function
 
' *****************************************************************************
' Name      UpdateSheet
' Param     iIndex  : 追加したシートのindex
'           wb      : ブック
' Result    Nothing
' Memo      新規追加したシートのタイトルなどを変更する
' *****************************************************************************
Sub UpdateSheet(iIndex As Integer, wb As Workbook)
    Dim ws As Worksheet
    Dim strTitle As String
    Dim strDay As String
    Dim strWk As String
    Dim strNendo As String
    Dim iYear As Integer
    Dim iMonth As Integer
    Dim iDay As Integer
    
    Set ws = wb.Sheets(iIndex)
    ' 年度取得
    strNendo = StrConv(GetNendo(Now), vbWide)
    
    ' ------------タイトル設定------------
    ' 対象者の場合はN月1日
    If bTarget Then
        strDay = cmbMonth.Text & "1日"
        strTitle = Replace(TITLE_FIRST, REPLACE_STRING, strDay)
    ' 実績の場合はN月末日
    Else
        strWk = cmbMonth.ListIndex + 1
        iMonth = CInt(strWk)
        iMonth = iMonth + 1
        ' システム日付から年を取り出して翌月-1日する
        iYear = Year(Date)
        wkDate = CDate(CStr(iYear) & "/" & Format(iMonth, MM) & "/01")
        wkDate = wkDate - 1
        
        strDay = cmbMonth.Text & CStr(Day(wkDate)) & "日"
        strTitle = Replace(TITLE_LAST, REPLACE_STRING, strDay)
    End If
    
    ' 年度も置換
    strTitle = Replace(strTitle, REPLACE_YEAR, strNendo)
    
    ' セルにセット
    ws.Range("A1").Value = strTitle
    
    ' ------------行列の表示切替------------
    ' 対象者は行を非表示、実績は表示する
    ws.Rows(4).Hidden = bTarget
    ws.Rows(5).Hidden = bTarget
    ws.Rows(6).Hidden = bTarget
    
End Sub

' *****************************************************************************
' Name      IsUsedName
' Param     wb          : ブック
'           strNewName  : シート名
' Result    Boolean     : T: 使われている/ F: 未使用
' Memo      指定したシート名が使われているかチェック
' *****************************************************************************
Function IsUsedName(wb, strNewName) As Boolean
    Dim iCnt As Integer
    IsUsedName = False
    
    For iCnt = 1 To wb.Sheets.Count - 1
        If wb.Sheets(iCnt).Name = strNewName Then
            IsUsedName = True
            Exit Function
        End If
    Next iCnt
    
End Function


' *****************************************************************************
' Name      WorkBookClose
' Param     File        : ブック名
' Result    Nothing
' Memo      指定したブックが開いていたら閉じる
' *****************************************************************************
Sub WorkBookClose(File)
    '変数を宣言し、閉じたいブック名を代入する
    Dim wb As Workbook
 
    For Each wb In Workbooks
        If wb.Name = File Then   'wbの中身が指定のブック名なら
        
            Application.DisplayAlerts = False 'システムメッセージを一旦OFFにし
            wb.Close 'wbを閉じる
            Application.DisplayAlerts = True 'システムメッセージを再びONへ
            Exit Sub
            
        End If
    Next wb '次のブックを変数wbへ
        
End Sub

' *****************************************************************************
' Name      GetFileName
' Param     strPath : フルパス
' Result    String  : ファイル名
' Memo      ファイルのフルパスからファイル名を取り出す
' *****************************************************************************
Function GetFileName(strPath) As String
    Dim strFileName As String
    
    ' いちばん最後の\の位置より右がファイル名
    Dim pos As Long
    pos = InStrRev(strPath, "\")
    
    If (0 < pos) Then
        strFileName = Right(strPath, Len(strPath) - pos)
    Else
        strFileName = ""
    End If
    
    GetFileName = strFileName
End Function

' *****************************************************************************
' Name      GetNendo
' Param     dt      : 日付
' Result    Long    : 年度
' Memo      日付から和暦の年度を返す
' *****************************************************************************
Function GetNendo(ByVal dt As Date) As Long
Dim iNendo As Integer

    iNendo = Format(dt, "e")
    If Month(dt) >= 4 Then
        GetNendo = iNendo
    Else
        GetNendo = iNendo - 1
    End If
End Function
