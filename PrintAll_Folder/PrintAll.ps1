$folder = "�����"

Write-Host "$folder �t�H���_������PDF�t�@�C�������ׂĈ�����܂��B"
Write-Host "���t�@�C���Ԃ̃X���[�v�����F2�b`r`n"

#�t�@�C�����̏����ň�����s
Dir $folder | Sort Name | ForEach{

    #�t�@�C�����\����A���s
    Write-Host $_.Name
    Start-Process $_.FullName -Verb Print | Stop-Process

    #2�b�X���[�v
    Start-Sleep -s 2
}

Write-Host "`r`n�����I`r`n"
Read-Host "�~�{�^���ŏI��"