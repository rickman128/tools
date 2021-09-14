$folder = "印刷物"

Write-Host "$folder フォルダ直下のPDFファイルをすべて印刷します。"
Write-Host "※ファイル間のスリープ処理：2秒`r`n"

#ファイル名の昇順で印刷実行
Dir $folder | Sort Name | ForEach{

    #ファイル名表示後、実行
    Write-Host $_.Name
    Start-Process $_.FullName -Verb Print | Stop-Process

    #2秒スリープ
    Start-Sleep -s 2
}

Write-Host "`r`n完了！`r`n"
Read-Host "×ボタンで終了"