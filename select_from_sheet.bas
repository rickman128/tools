Option Explicit

Sub sql_01()
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim File_Name, Sql As String
  Dim CurRow As Integer
    File_Name = ThisWorkbook.FullName
    
    CurRow = 1
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & "DBQ=" & File_Name & "; ReadOnly=True;"
    cn.Open
    ' 1
    'Sql = "SELECT 内容_研修タイトル, 計算_研修種類 FROM [シリーズ$] WHERE 計算_研修種類='定期研修' "
    ' 2
    'Sql = "SELECT DISTINCT 研修実施Code FROM [参加者$] WHERE 職員_職種 not in ('医師', '歯科医師')"
    ' 3
    'Sql = "SELECT SRS.研修タイトルCode, MST.研修実施Code, SRS.内容_研修タイトル " & _
        "FROM [シリーズ$] AS SRS " & _
        "INNER JOIN [実施マスタ$] AS MST " & _
        "ON SRS.研修タイトルCode = MST.研修タイトルCode "

    ' 【Fix】 研修シリーズと総参加者数、医師・歯科医師以外の参加人数
    Sql = "SELECT SRS.研修タイトルCode, SRS.内容_研修タイトル, SRS.集計_総参加者数, NOT_DR.CNT " & _
        " FROM ( " & _
        "SELECT MST.研修タイトルCode, Count(ALL_DR.研修実施Code) AS CNT " & _
        "FROM [実施マスタ$] AS MST " & _
        "INNER JOIN (SELECT 研修実施Code FROM [参加者$] WHERE 職員_職種 not in ('医師', '歯科医師')) AS ALL_DR " & _
        "ON MST.研修実施Code = ALL_DR.研修実施Code " & _
        "GROUP BY MST.研修タイトルCode " & _
        " ) NOT_DR " & _
        " INNER JOIN [シリーズ$] AS SRS " & _
        " ON NOT_DR.研修タイトルCode = SRS.研修タイトルCode "


    Debug.Print Sql
    
    rs.Open Sql, cn, adOpenStatic
    Do Until rs.EOF
        Sheets("SQL実行").Cells(CurRow, 1).Value = rs!研修タイトルCode
        Sheets("SQL実行").Cells(CurRow, 2).Value = rs!内容_研修タイトル
        Sheets("SQL実行").Cells(CurRow, 3).Value = rs!集計_総参加者数
        Sheets("SQL実行").Cells(CurRow, 4).Value = rs!CNT
        rs.MoveNext
        CurRow = CurRow + 1
    Loop
    rs.Close
    cn.Close
  Set rs = Nothing
  Set cn = Nothing
End Sub
