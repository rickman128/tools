Attribute VB_Name = "Module1"
' *****************************************************************************
'           FileMaker�̋Ζ��������璊�o�����t�@�C������
'           ���ߋΖ��\�����c�[��
' Date      2020/11/12  K.Endo
' In        .txt�t�@�C��    : FileMaker�Ɩ���������Export�����������Ԃ̃t�@�C��
' Out       .xls�t�@�C���~n : �����o���̒��Ε\Excel�t�@�C��
' Memo      1. �J���}��؂�̃e�L�X�g�t�@�C����I������
'           2. Unicode��txt��SJIS��csv�ɕϊ�����(PowerShell���s)
'           3. csv�t�@�C������C���|�[�g����
'           4. �������ԂƋΖ��`�Ԃ��璴�Ύ��Ԃ����߂�
'           5. �����o���Ƃɒ��΃t�@�C�����쐬����4��]�L�A�o�͂���
'
' History   <#001> G�h���C�u�Ńt�@�C�������w��ł���悤�ɁA�o�̓t�@�C������
'                  �^�C���X�^���v��t������̂���߂�
' *****************************************************************************
' -----------------------------------------
' �\���̒�`
' -----------------------------------------
' �ǂݍ���txt�t�@�C���̒�`�@1�s��
Type REC_WORK_HOUR
    workDay As Date         ' ���ɂ�
    name As String          ' ����
    strTypeOfWork As String ' �Ζ��`�ԁi�����j
    startTime As Date       ' �������ԁi���j
    endTime As Date         ' �������ԁi���j 24���𒴂����ꍇ��-24��������
    memo As String          ' ����
End Type

' �������ރf�[�^�̒�`�{�C���|�[�g�f�[�^
Type REC_OVER_WORK
    workDay As Integer      ' 1�`31�Œ�
    strTypeOfWork As String ' �Ζ��`��
    startTime1 As Date      ' �Ζ����ߎ��ԁi��i�j��
    endTime1 As Date        ' �Ζ����ߎ��ԁi��i�j��
    startTime2 As Date      ' �Ζ����ߎ��ԁi���i�j��
    endTime2 As Date        ' �Ζ����ߎ��ԁi���i�j��
    memo As String          ' �Ɩ����e
    ' �C���|�[�g�f�[�^
    startWorkTime As Date   ' �������ԁi���j
    endWorkTime As Date     ' �������ԁi���j
End Type

' ���ߋΖ����ԍ\���́i�㉺�i����̂�2�Z�b�g�j
Type REC_OVER_TIME
    startTime1 As Date      ' ���ߎ��ԂP�i���j
    endTime1 As Date        ' ���ߎ��ԂP�i���j
    startTime2 As Date      ' ���ߎ��ԂQ�i���j
    endTime2 As Date        ' ���ߎ��ԂQ�i���j
End Type

' -----------------------------------------
' �萔��`
' -----------------------------------------
' �Ζ��`��(�Ζ������̒�`�j
Const TYPE_WORK_NIKKIN = "����"
Const TYPE_WORK_NITIYA = "����"
Const TYPE_WORK_NAGA = "����"
Const TYPE_WORK_SYUKU = "�h��"
Const TYPE_WORK_HIBAN = "���"
Const TYPE_WORK_HAYAA = "��A"
Const TYPE_WORK_HAYAB = "��B"
Const TYPE_WORK_DAIKYU = "��x"
Const TYPE_WORK_NENKYU = "�N�x"
Const TYPE_WORK_TOKKYU = "���x"
Const TYPE_WORK_4 = "4"
Const TYPE_WORK_TYOKIN = "����"
Const TYPE_WORK_SYUTTYO = "�o��"
Const TYPE_WORK_KIBIKI = "����"

' �Ζ��`��(���Ε\�̒�`�j
Const TYPE_OVER_NIKKIN = "����"
Const TYPE_OVER_YAKIN = "���"
Const TYPE_OVER_HIBAN = "���"
Const TYPE_OVER_KYUJITU = "�x��"
Const TYPE_OVER_SYUKU = "�h��"
Const TYPE_OVER_DAIKYU = "��x"
Const TYPE_OVER_CLEAR = "���x"  ' ���x�A�N�x�͋󗓂ɂ���i���Ε\�ɂ͒�`���Ȃ��j

' Excel�̃Z���z��im_aryCell�j��index
Const CELLARY_IDX_YM = 1        ' �N��
Const CELLARY_IDX_BMN = 2       ' ����
Const CELLARY_IDX_KA = 3        ' �ہE��
Const CELLARY_IDX_NAME = 4      ' ����
Const CELLARY_IDX_HOLIDAY = 5   ' �j��
Const CELLARY_IDX_TYPE = 6      ' �Ζ��`��
Const CELLARY_IDX_YOTEI_ST = 7  ' �Ζ��\�莞�ԊJ�n
Const CELLARY_IDX_YOTEI_ED = 8  ' �Ζ��\�莞�ԏI��
Const CELLARY_IDX_MEIREI_ST = 9 ' �Ζ����ߎ��ԂP�J�n
Const CELLARY_IDX_MEIREI_ED = 10 ' �Ζ����ߎ��ԂP�I��
Const CELLARY_IDX_MEMO = 11     ' �Ɩ����e

' �����o��l���̍\����
Type REC_ARY_OVERS
    name As String              ' ���O
    aryOver(31) As REC_OVER_WORK ' 1�������̒��Ε\
End Type

' �����o���i�V�[�g��T�����o�e�[�u������擾����j
Type MEMBER_LIST
    name As String              ' ����
    lastName As String          ' ��
End Type

' -----------------------------------------
' �����o�ϐ���`
' -----------------------------------------
                                ' �����o���X�g
Public m_lstMember() As MEMBER_LIST
                                ' �Ζ����т̍\���̔z��i�t�@�C������C���|�[�g�����f�[�^�j
Public m_aryWork() As REC_WORK_HOUR
                                ' �����o���Ƃ̒��Δz��
Public m_aryOver() As REC_ARY_OVERS
                                ' �o�͔N��
Public m_outDate As Date
                                ' �o�͐�
Public m_outPath As String
                                ' ���Ε\�̌��{�t�@�C���p�X
Public m_strFileName As String
                                ' ���Ε\�̃V�[�g��
Public m_strSheetName As String
                                ' ���Ε\Excel�t�@�C���̃Z���ʒu
Public m_aryCell(11) As String

' *****************************************************************************
' Name      Init
' Param     Nothing
' Result    Nothing
' Memo      ��������
' *****************************************************************************
Sub Init()
    Dim iCnt As Integer
    Dim iRow As Integer
    Dim iMemberCnt As Integer
    Dim aryStrName() As Variant   ' Array���g���̂�Variant����Ȃ��Ƃ���
    
    ' -----------------------------------------
    ' ���̓p�����[�^
    ' -----------------------------------------
    m_outDate = Sheets("�c�[��").Range("�o�͔N��").Value
    m_outPath = Sheets("�c�[��").Range("�o�͐�").Value
    
    ' -----------------------------------------
    ' ���Ε\�̒�`
    ' -----------------------------------------
    m_strFileName = Sheets("�c�[��").Range("���{�t�@�C���p�X").Value
    m_strSheetName = Sheets("�c�[��").Range("�V�[�g��").Value
    
    ' �������݃Z���ʒu�̒�`��(VBA�ł͔z��const����`�ł��Ȃ�)
    aryStrName = Array("", _
                        "�N��", _
                        "����", _
                        "�ہE��", _
                        "����", _
                        "�j���P", _
                        "�Ζ��`�ԂP", _
                        "�Ζ��\�莞�ԂP�J�n", _
                        "�Ζ��\�莞�ԂP�I��", _
                        "�Ζ����ߎ��ԂP�J�n", _
                        "�Ζ����ߎ��ԂP�I��", _
                        "�Ɩ����e�P")
    
    
    ' ���O�̂����Z������l�����o��
    For iCnt = 1 To UBound(m_aryCell)
        m_aryCell(iCnt) = Sheets("�c�[��").Range(aryStrName(iCnt)).Value
    Next
    
    
    ' -----------------------------------------
    ' �����o�ꗗ
    ' -----------------------------------------
                                            ' �����o���擾(���o���̕�������)
    iMemberCnt = Sheets("�����o�ꗗ").Range("T����").Rows.Count
    
    ReDim m_lstMember(iMemberCnt - 1)       ' �����o���z����g��
    
    ' �����o���X�g���V�[�g����ǂݍ���
    For iCnt = 0 To iMemberCnt - 1
                                            ' ����
        m_lstMember(iCnt).name = Sheets("�����o�ꗗ").Range("T����").Cells(iCnt + 1, 2) ' +1�͌��o���̕�
                                            ' ��
        m_lstMember(iCnt).lastName = Sheets("�����o�ꗗ").Range("T����").Cells(iCnt + 1, 3)
    Next
    
    ReDim m_aryOver(iMemberCnt - 1)         ' �z��������o�̐l�����g��
    
    
End Sub

' *****************************************************************************
' Name      ImportFile
' Param     Nothing
' Result    Boolean : T or F
' Memo      txt�t�@�C����ǂݍ���ŋΖ����э\���̂ɂ߂�
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
    
    ' txt�t�@�C����I��
    ofn = Application.GetOpenFilename(FileFilter:="�e�L�X�g�t�@�C�� (*.txt), *.txt", _
                                    Title:="�C���|�[�g�t�@�C���I��")
'    ofn = Application.GetOpenFilename(FileFilter:="CSV�t�@�C�� (*.csv), *.csv", _
'                                    Title:="�C���|�[�g�t�@�C���I��")

    ' �߂�̓o���A���g�^�B�������J����΃t�@�C���p�X���Ԃ��Ă���
    If ofn <> "False" Then
        ' ��������PowerShell�ŃG���R�[�h�B������csv�ɂȂ�
        strOutPath = UnicodeToSJIS(ofn)
        ' �ϊ����csv�t�@�C���I�[�v��
        Open strOutPath For Input As #iFileNumber
        'Open ofn For Input As #iFileNumber
    Else
        ImportFile = False
        Exit Function
    End If
    
    ' csv�t�@�C����ǂݍ���
    Do While Not EOF(iFileNumber)
        iCnt = iCnt + 1
        ReDim Preserve m_aryWork(iCnt)      ' �z��g��
                                            ' �Ζ�����1�s����z��ɃZ�b�g
        Input #iFileNumber, strDay, _
                            m_aryWork(iCnt - 1).name, _
                            m_aryWork(iCnt - 1).strTypeOfWork, _
                            strStart, _
                            strEnd, _
                            m_aryWork(iCnt - 1).memo
                            
        ' ���t�A���Ԃ͂�������ϐ��ɓ���Ă���^�ϊ�����
        If IsDate(strDay) Then
            m_aryWork(iCnt - 1).workDay = CDate(strDay)
        End If
        ' �������ԁi���j
        If IsDate(strStart) Then
            m_aryWork(iCnt - 1).startTime = CDate(strStart)
        End If
        ' �������ԁi���j
        If IsDate(strEnd) Then
            m_aryWork(iCnt - 1).endTime = CDate(strEnd)
        ElseIf strEnd <> "" Then
            ' 24���𒴂����\�L�̏ꍇ�A�I��莞�Ԃ�23:59�ɂ��āAGetOverTime�ł܂�߂�
            strEnd = "23:59:00"
            m_aryWork(iCnt - 1).endTime = CDate(strEnd)
        End If
        
    Loop
    
    GoTo Finally
        
ErrHandler:
    ImportFile = False
                                            ' �G���[���������烁�b�Z�[�W
    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "�G���["
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
' Memo     �C���|�[�g�����f�[�^���璴�΍\���̔z����쐬����i�z��[�����o][1..31]�j�ɂ߂�
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
        ' ���ɂ�
        iDay = Day(m_aryWork(iCnt).workDay)
        ' ���O���烁���o���X�g��index�擾
        iMember = GetMemberIdx(m_aryWork(iCnt).name)
        ' ���X�g�ɂȂ��l�͔�΂�
        If iMember < 0 Then
            GoTo continue
        End If
        
        With m_aryOver(iMember).aryOver(iDay)
            ' �R�s�[���鍀��
                                            ' ���ɂ��i1-31��day�̂݁j
            .workDay = iDay
                                            ' �������ԁi���j
            .startWorkTime = m_aryWork(iCnt).startTime
                                            ' �������ԁi���j
            .endWorkTime = m_aryWork(iCnt).endTime
                                            ' ����
            .memo = m_aryWork(iCnt).memo
            
        
            ' �v�Z���鍀��
                                            ' �Ζ��`�ԁiFM�p���璴�Ε\�p�ɃR���o�[�g�j
            strType = GetOutTypeOfWork(m_aryWork(iCnt).strTypeOfWork)
            .strTypeOfWork = strType
                                            ' �f�t�H���g�Ζ�����
            Call GetDefaultTime(m_aryWork(iCnt).strTypeOfWork, startDefTime, endDefTime)
                                            ' ���Ύ���
                                            ' �f�t�H���ԂƎ������Ԃ��璴�Ύ��Ԃ�ݒ�
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
' Param     strName     : ���O
' Result    Integer     : ���O�������Ă���z���index
' Memo     �����o�̖��O�������Ă���z���index��Ԃ�
'           -1: ������Ȃ������ꍇ
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
' Memo     ���΍\���̔z�񂩂�e�����o���Ƃ�Excel�t�@�C�����o�͂���
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
       
    ' �w�肵���o�͔N��
    strDate = Format(m_outDate, "yyyymm")
    ' �o�͓���
    ' <#001> DEL
    'strTimeStamp = Format(Now, "yyyymmddhhnnss")
    
    ' �o�͐�̑��݃`�F�b�N
    If Not objFileSystem.FolderExists(m_outPath) Then
        MsgBox "�o�͐�p�X�����݂��܂���B���������͂��Ă��������B"
        Exit Sub
    End If
    
    ' ���{�t�@�C����l�����R�s�[
    For iCnt = LBound(m_lstMember) To UBound(m_lstMember)
        ' <#001> MOD start
        ' ex) �u���O202011_20201127xxxxxx.xls�v
        'strFileName = m_outPath & m_lstMember(iCnt).lastName & strDate & "_" & strTimeStamp & ".xlsx"
        
        ' ex) �u���O202011.xls�vG�h���C�u�Ńt�@�C�������w�肷�邽�߂Ƀ^�C���X�^���v�̕t������߂�
        strFileName = m_outPath & m_lstMember(iCnt).lastName & strDate & ".xlsx"
        ' <#001> MOD end
        
        ' ���{�t�@�C���̑��݃`�F�b�N
        If Dir(m_strFileName) = "" Then
            MsgBox "���Ε\���{�̃t�@�C���p�X���m�F���Ă��������B"
            Exit Sub
        End If
        
        ' �R�s�[
        FileCopy m_strFileName, strFileName
        
        ' �t�@�C���I�[�v��
        If Dir(strFileName) <> "" Then
            Set objBook = Workbooks.Open(strFileName)
        Else
            MsgBox "�t�@�C�� �y" & strFileName & "�z �����݂��܂���B", vbExclamation
            Exit Sub
        End If
      
        ' �t�@�C���̏�������
        If Not WriteOutFile(iCnt, objBook) Then
            MsgBox "�V�[�g �y" & m_strSheetName & "�z �����݂��܂���B", vbExclamation
            GoTo Finally
        End If
        
        ' �ۑ�
        objBook.Save
        
        ' ����
        objBook.Close
    Next

    GoTo Finally
    
ErrHandler:
                                            ' �G���[���������烁�b�Z�[�W
    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "�G���["
    Resume Finally
    
Finally:
    ' �Ō�ɏ������Ă����u�b�N�����Ă��Ȃ����������
    For Each wb In Workbooks
        If wb.name = Dir(strFileName) Then
            wb.Close
        End If
    Next
End Sub

' *****************************************************************************
' Name      WriteOutFile
' Param     iMemberIdx  : m_lstMember��index
'           objBook     : ��������ExcelBook
' Result    Boolean     : T: ���� or F: �ُ�
' Memo      �t�@�C���ɏ������ރ��C�������i�������ރt�@�C����Open�ς݁j
' *****************************************************************************
Function WriteOutFile(iMemberIdx As Integer, objBook As Workbook) As Boolean
On Error GoTo ErrHandler
    Dim iDay As Integer
    Dim iNextPageOffset As Integer
    Dim iOffset As Integer
    Dim iOffset2 As Integer
    Dim objSheet As Worksheet
    
    WriteOutFile = True
    
    ' �V�[�g���Ȃ������牽�����Ȃ�
    If Not SheetExists(m_strSheetName, objBook) Then
        WriteOutFile = False
        Exit Function
    End If

    Set objSheet = objBook.Sheets(m_strSheetName)
    
    ' �N��
    objSheet.Range(m_aryCell(CELLARY_IDX_YM)).Value = m_outDate
    ' ������
    'objSheet.Range(m_aryCell(CELLARY_IDX_BMN)).Value =
    ' ���ہE��
    'objSheet.Range(m_aryCell(CELLARY_IDX_KA)).Value =
    ' ����
    objSheet.Range(m_aryCell(CELLARY_IDX_NAME)).Value = m_lstMember(iMemberIdx).name
    
    ' for 1���`31�����[�v
    For iDay = 1 To 31
        ' 2�y�[�W�ڂ͐擪�ʒu��ύX
        If iDay > 15 Then
            iNextPageOffset = 10
        Else
            iNextPageOffset = 0
        End If
        
        ' 1���ɂ�2�s
        iOffset = iNextPageOffset + ((iDay - 1) * 2 - 1)
        iOffset2 = iNextPageOffset + iDay * 2 - 2
        
        ' 1���ڂ�����Offset���w��
        If iOffset < 0 Then
            iOffset = 0
        End If
        
        With m_aryOver(iMemberIdx).aryOver(iDay)
            ' ����������������
            If Month(DateSerial(Year(m_outDate), Month(m_outDate), iDay)) <> Month(m_outDate) Then
                ' �Ζ��`�Ԃ��N���A
                objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = ""
                GoTo continue
            End If
            ' �j��
            objSheet.Range(m_aryCell(CELLARY_IDX_HOLIDAY)).Offset(iOffset, 0).Value = GetHoliday(DateSerial(Year(m_outDate), Month(m_outDate), iDay))
            
            ' �Ζ��`��
            Select Case .strTypeOfWork
                Case TYPE_OVER_CLEAR
                    ' ���x�E�N�x�̏ꍇ�̓N���A
                    objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = ""
                Case ""
                    ' ���ݒ�̏ꍇ�͋x��
                    objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = TYPE_OVER_KYUJITU
                Case Else
                    ' ���̑�
                    objSheet.Range(m_aryCell(CELLARY_IDX_TYPE)).Offset(iOffset, 0).Value = .strTypeOfWork
            End Select
                                      
            ' ---------��i---------
            If (.startTime1 <> "0:00") Or (.endTime1 <> "0:00") Then
                ' *** �Ζ��\�莞�Ԃ͐����ŋΖ����ߎ��Ԃ��R�s�[���� ***
                ' �Ζ��\�莞�ԁ@�J�n
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ST)).Offset(iOffset2, 0).Value = .startTime1
                ' �Ζ��\�莞�ԁ@�I��
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ED)).Offset(iOffset2, 0).Value = .endTime1
                
                ' �Ζ����ߎ��ԁ@�J�n
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ST)).Offset(iOffset2, 0).Value = GetRoundTime(.startTime1)
                ' �Ζ����ߎ��ԁ@�I��
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ED)).Offset(iOffset2, 0).Value = GetRoundTime(.endTime1)
            End If
            
            ' ---------���i---------
            If (.startTime2 <> "0:00") Or (.endTime2 <> "0:00") Then
                ' *** �Ζ��\�莞�Ԃ͐����ŋΖ����ߎ��Ԃ��R�s�[���� ***
                ' �Ζ��\�莞�ԁ@�J�n
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ST)).Offset(iOffset2 + 1, 0).Value = .startTime2
                ' �Ζ��\�莞�ԁ@�I��
                'objSheet.Range(m_aryCell(CELLARY_IDX_YOTEI_ED)).Offset(iOffset2 + 1, 0).Value = .endTime2
                
                ' �Ζ����ߎ��ԁ@�J�n
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ST)).Offset(iOffset2 + 1, 0).Value = GetRoundTime(.startTime2)
                ' �Ζ����ߎ��ԁ@�I��
                objSheet.Range(m_aryCell(CELLARY_IDX_MEIREI_ED)).Offset(iOffset2 + 1, 0).Value = GetRoundTime(.endTime2)
            End If
            
            ' �Ɩ����e
            objSheet.Range(m_aryCell(CELLARY_IDX_MEMO)).Offset(iOffset2, 0).Value = .memo

        End With
continue:
    Next

    Exit Function
ErrHandler:
                                            ' �G���[���������烁�b�Z�[�W
    MsgBox "�yWriteOutFile�z " & Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "�G���["
End Function

' *****************************************************************************
' Name      GetOutTypeOfWork
' Param     strType :   �Ζ��`��(�Ɩ������j
' Result    Integer :   �Ζ��`��(���Ε\�̒�`�j
' Memo     �Ɩ������̋Ζ��`�ԁi������j���璴�Ε\�̋Ζ��`�ԁi������j��Ԃ�
' *****************************************************************************
Function GetOutTypeOfWork(strType As String) As String
    Dim strRet As String
    
    ' FM�̋Ζ��`�Ԃ��璴�Ε\�̒�`��Ԃ�
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
            ' ���x�A�N�x�͋�ɂ���
            strRet = TYPE_OVER_CLEAR
        Case Else
            strRet = ""
    End Select
    
    GetOutTypeOfWork = strRet
    
End Function

' *****************************************************************************
' Name      GetDefaultTime
' Param     strType     :   �Ζ��`��(FM�̒�`�j
'           startTime   :   �f�t�H���g�Ζ����ԁi���jout
'           endTime     :   �f�t�H���g�Ζ����ԁi���jout
' Result    Nothing
' Memo     �Ζ��`�ԁi���Ε\�j����f�t�H���g�Ζ����Ԃ�Ԃ�
' *****************************************************************************
Sub GetDefaultTime(strType As String, ByRef startTime As Date, ByRef endTime As Date)
    
    Select Case strType
        ' ���΁A�o��
        Case TYPE_WORK_NIKKIN, _
            TYPE_WORK_SYUTTYO
            startTime = TimeValue("8:30")
            endTime = TimeValue("17:15")
        ' ����A����
        Case TYPE_WORK_NITIYA, _
            TYPE_WORK_NAGA
            startTime = TimeValue("8:30")
            endTime = TimeValue("21:30")
        ' ��ԁA4
        Case TYPE_WORK_HIBAN, _
            TYPE_WORK_4
            startTime = TimeValue("8:30")
            endTime = TimeValue("12:30")
        ' ��A
        Case TYPE_WORK_HAYAA
            startTime = TimeValue("7:45")
            endTime = TimeValue("16:30")
        ' ��B
        Case TYPE_WORK_HAYAB
            startTime = TimeValue("8:00")
            endTime = TimeValue("16:45")
        ' ��x�A�h���A�N�x�A���x�A���΁A����
        Case TYPE_WORK_DAIKYU, _
            TYPE_WORK_SYUKU, _
            TYPE_WORK_NENKYU, _
            TYPE_WORK_TOKKYU, _
            TYPE_WORK_TYOKIN, _
            TYPE_WORK_KIBIKI
            startTime = TimeValue("0:00")
            endTime = TimeValue("0:00")
        ' ���̑�
        Case Else
            startTime = TimeValue("0:00")
            endTime = TimeValue("0:00")
    End Select
End Sub

' *****************************************************************************
' Name      GetOverTime
' Param     defStartTime   :   �f�t�H���g�Ζ����ԁi���j
'           defEndTime     :   �f�t�H���g�Ζ����ԁi���j
'           workStartTime  :   �����Ζ����ԁi���j
'           workEndTime    :   �����Ζ����ԁi���j
'           recOver        :   ���ߋΖ����ԍ\���́@out
' Result    Nothing
' Memo     �f�t�H���g�Ζ����ԂƎ������Ԃ��璴�ߋΖ����Ԃ�Ԃ�
' *****************************************************************************
Sub GetOverTime(defStartTime, defEndTime, _
                workStartTime, workEndTime As Date, _
                ByRef recOver As REC_OVER_TIME)
    ' �������ԂȂ��̏ꍇ��00:00�𖄂߂ĕԂ�
    If ((Not IsDate(workStartTime)) And (Not IsDate(workEndTime))) Or _
        ((workStartTime = "0:00:00") And (workEndTime = "0:00:00")) Then
        recOver.startTime1 = "00:00"
        recOver.endTime1 = "00:00"
        recOver.startTime2 = "00:00"
        recOver.endTime2 = "00:00"
        Exit Sub
    End If
        
    ' ���ߋΖ��P
    If defStartTime <= workStartTime Then
        recOver.startTime1 = "00:00"
        recOver.endTime1 = "00:00"
    Else
        recOver.startTime1 = workStartTime
        recOver.endTime1 = defStartTime
    End If
           
    ' ���ߋΖ��Q
    If defEndTime >= workEndTime Then
        recOver.startTime2 = "00:00"
        recOver.endTime2 = "00:00"
    Else
        recOver.startTime2 = defEndTime
        recOver.endTime2 = workEndTime
    End If

    ' ���ߋΖ��P���Ȃ��ĂQ�����̏ꍇ�͂P�i��i�j�ɂQ�̓��e���߂�
    If (recOver.startTime1 = "00:00") And (recOver.startTime2 <> "00:00") Then
        recOver.startTime1 = recOver.startTime2
        recOver.endTime1 = recOver.endTime2
        recOver.startTime2 = "00:00"
        recOver.endTime2 = "00:00"
    End If
    
End Sub

' *****************************************************************************
' Name      GetHoliday
' Param     dt      : ���t
' Result    String  : "�j��", "�x��", ""
' Memo     �Ώۂ̓��t���j�����x�����A����ȊO����Ԃ�
' *****************************************************************************
Function GetHoliday(dt As Date) As String
    Dim str As String
    
        str = ���{�̋x��(dt)
        
        Select Case str
            Case "�j", "�x", "�U"
                GetHoliday = "�j��"
            Case Else
                GetHoliday = ""
        End Select
            
            
End Function

' *****************************************************************************
' Name      GetRoundTime
' Param     dt      : ����
' Result    String  : ���Ԃ𕶎���ɂ������́i23:59������24:00�ɂ���j
' Memo      23:59��24:00�ɂ��ĕԂ�
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
' Param     SheetName   : �V�[�g��
'           wb          : Workbook
' Result    Boolean     : T: ���� F: �Ȃ�
' Memo      �V�[�g�̑��݃`�F�b�N
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
' Param     strInPath   : txt�t�@�C���̃t���p�X
' Result    String      : csv�t�@�C���̃t���p�X�i�g���q���ς���������j
' Memo      Unicode(UTF16,BOM��)��txt�t�@�C������SJIS��csv���쐬����
' *****************************************************************************
Function UnicodeToSJIS(strInPath As String) As String
    Dim strOutPath As String
    Dim strCmd As String

                                            ' txt�t�@�C���p�X
    strOutPath = Replace(strInPath, ".txt", ".csv")
                                            ' PowerShell�̃R�}���h
    strCmd = "get-content -Encoding Unicode " & strInPath & "| Set-Content " & strOutPath
    
    Call RunPowerShell(strCmd, 0, True)     ' ���s
    
    UnicodeToSJIS = strOutPath
End Function

' *****************************************************************************
' Name      RunPowerShell
' Param     psCmd   : PowerShell�R�}���h(PowerShell�d�l�ɂ��260�����ȓ�)
'           intVsbl : PowerShell��ʕ\��(1)/��\��(0)
'           waitFlg : PowerShell���s�I����҂�(True)/�҂��Ȃ�(False)
' Result
' Memo      PowerShell�����s����
' *****************************************************************************
Function RunPowerShell(psCmd As String, intVsbl As Integer, waitFlg As Boolean)
    Dim objWSH As Object
    Set objWSH = CreateObject("WScript.Shell")

    objWSH.Run "powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & psCmd, intVsbl, waitFlg

End Function
' *****************************************************************************
' �ȉ�github����
' https://gist.github.com/ooltcloud/30644a3ce497319e5d50
' *****************************************************************************
'------------------------------------------------------------------
' Excel VBA (�W�����W���[��)
' ���t���w�肷��Ƌx�݂̑����ɉ����āA�j/�x/�U/��/�y�A��߂��܂��B�����͋󕶎��B
'------------------------------------------------------------------
' ������
' �@2019�N,2020�N�ɑΉ����܂����B
' �@�@�t���̓��ƏH���̓��͌v�Z����̂��ʓ|�Ȃ̂ŁAWikipedia ����\�z�����R�s�y���Ă��܂��B�i��
' �@�@�O�N�Ɋ���Ŕ��\�������t�ƈ���Ă����ꍇ�ͤ�K�X�C�����Ă��������
' �@�A�t���̓��ƏH���̓��� 2000 �N����ݒ肵�Ă��܂����A���̏j���� 2013 �N�ȑO�̏󋵂𔽉f���܂���B(��:�݂ǂ�̓�)
' �@�@�����܂ō��N 2014 �N�ȍ~�̕\���p�̃v���O�����ł��B�ߋ��̏j���𐳊m�ɔ��f�������ꍇ�́A�C�����K�v�ł��B
' �@�B�j���@���������ꂽ�礓K�X�C�����Ă�������� (�����͗\�z�ł��܂���)
' �@�C���܂��Łu������ȋx���v�𑫂��Ă��܂��B��ЂƂ��̋x�݂𔽉f�������ꍇ�͏ꍇ�͂����֋L�q���܂��B
' �@�@�ȉ��̃v���O�������ƃT���v����᱗��~�̓����Ƃ��������Ă��܂���s�v�ȏꍇ�͓��Y���[�`�������ьĂяo�����폜���Ă��������
' �@�@�܂��u���Ԃ͏j�����������̉�Ђ͋x�݂���Ȃ��v�ȂǁA�j����ł��������W�b�N�͗p�ӂ��Ă��܂���i��
'------------------------------------------------------------------
Public Function ���{�̋x��(vInDate As Date) As String

    Dim ret As String
    
    ' �j������
    ret = ���{�̏j��(vInDate)
    If ret <> "" Then
        ���{�̋x�� = "�j"
        Exit Function
    End If
        
    ' �����̋x������
    If is�����̋x��(vInDate) = True Then
        ���{�̋x�� = "�x"
        Exit Function
    End If
    
    ' �U�֋x������
    If is�U�֋x��(vInDate) = True Then
        ���{�̋x�� = "�U"
        Exit Function
    End If

    ' ������̋x�� (�~�Ƃ������̓��Ƃ���Ђ̋x��) / �K�v�ɉ�����
    ret = ������̋x��(vInDate)
    If ret <> "" Then
        ���{�̋x�� = ret
        Exit Function

    End If

    ' �y������
    Select Case (Weekday(vInDate))
        Case 1: ���{�̋x�� = "��"
        Case 7: ���{�̋x�� = "�y"
        Case Else: ���{�̋x�� = ""
    End Select

End Function


'------------------------------------------------------------------
Private Function ���{�̏j��(vInDate As Date) As String

    Dim strRet As String
    strRet = ""
    
    ' 2000�N����2030�N�̂܂ł̏t��/�H���̈ꗗ�\
    ' (Wkipedia��� http://ja.wikipedia.org/wiki/%E7%A7%8B%E5%88%86%E3%81%AE%E6%97%A5)
    Dim �t�� As Variant
    �t�� = Array(20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 21, 21, 20, 20, 20, 21, 20, 20, 20)
    
    Dim �H�� As Variant
    �H�� = Array(23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23, 23, 22, 23, 23)
    
    ' �v�f����
    Dim intyear As Integer:     intyear = Year(vInDate)
    Dim intday As Integer:      intday = Day(vInDate)
    Dim intMonth As Integer:    intMonth = Month(vInDate)
    Dim intWeekDay As Integer:  intWeekDay = Weekday(vInDate)
    Dim intWeek As Integer:     intWeek = (intday - 1) \ 7 + 1
    
    ' �����Œ�
    If (intMonth = 1) And (intday = 1) Then strRet = "����"
    If (intMonth = 2) And (intday = 11) Then strRet = "�����L�O�̓�"
    If (intMonth = 2) And (intday = 23) And (intyear >= 2020) Then strRet = "�V�c�a����"
    If (intMonth = 4) And (intday = 29) Then strRet = "���a�̓�"
    If (intMonth = 5) And (intday = 3) Then strRet = "���@�L�O��"
    If (intMonth = 5) And (intday = 4) Then strRet = "�݂ǂ�̓�"
    If (intMonth = 5) And (intday = 5) Then strRet = "���ǂ��̓�"
    If (intMonth = 8) And (intday = 11) And (intyear >= 2016) And (intyear <> 2020) Then strRet = "�R�̓�"
    If (intMonth = 11) And (intday = 3) Then strRet = "�����̓�"
    If (intMonth = 11) And (intday = 23) Then strRet = "�ΘJ���ӂ̓�"
    If (intMonth = 12) And (intday = 23) And (intyear <= 2018) Then strRet = "�V�c�a����"
    
    ' �t��/�H��
    If (2000 <= intyear And intyear <= 2030) Then
        If (intMonth = 3 And intday = �t��(intyear - 2000)) Then strRet = "�t���̓�"
        If (intMonth = 9 And intday = �H��(intyear - 2000)) Then strRet = "�H���̓�"
    End If
    
    ' �n�b�s�[�}���f�[(����T�̌��j���Œ�)
    If (intWeekDay = 2) Then    ' ���j
        If (intMonth = 1) And (intWeek = 2) Then strRet = "���l�̓�"
        If (intMonth = 7) And (intWeek = 3) And (intyear <> 2020) Then strRet = "�C�̓�"
        If (intMonth = 9) And (intWeek = 3) Then strRet = "�h�V�̓�"
        If (intMonth = 10) And (intWeek = 2) And (intyear <= 2019) Then strRet = "�̈�̓�"
        If (intMonth = 10) And (intWeek = 2) And (intyear >= 2021) Then strRet = "�X�|�[�c�̓�"
        
    End If
    
    ' 2019�N����
    If (intyear = 2019) Then
        If (intMonth = 5) And (intday = 1) Then strRet = "�V�c�̑��ʂ̓�"
        If (intMonth = 10) And (intday = 22) Then strRet = "���ʗ琳�a�̋V�̍s�����"
    End If

    ' 2020�N����
    If (intyear = 2020) Then
        If (intMonth = 7) And (intday = 23) Then strRet = "�C�̓�"
        If (intMonth = 7) And (intday = 24) Then strRet = "�X�|�[�c�̓�"
        If (intMonth = 8) And (intday = 10) Then strRet = "�R�̓�"
    End If
    
    ���{�̏j�� = strRet

End Function


'------------------------------------------------------------------
Private Function is�U�֋x��(vInDate As Date) As Boolean

    ' ���g���j���ł���΁A�U�ւł͂Ȃ�
    If ���{�̏j��(vInDate) <> "" Then
        is�U�֋x�� = False
        Exit Function
    End If

    ' ���g���j���łȂ���΁A�j�������j�܂ŘA�����Ă��邩���m�F
    For n = 1 To 7
        Dim d As Date
        d = DateAdd("d", -n, vInDate)   ' n���O
        
        ' ���j�Ɏ���܂łɏj�����r�؂��ΐU�ւłȂ�
        If (���{�̏j��(d) = "") Then
            is�U�֋x�� = False
            Exit Function
            
        Else
            If (Weekday(d) = 1) Then
                ' �j���ł�����j
                is�U�֋x�� = True
                Exit Function
            
            Else
                ' LOOP(����ɑO�����m�F)
            
            End If
        End If
    Next

End Function


'------------------------------------------------------------------
Private Function is�����̋x��(vInDate As Date) As Boolean

    ' ���g���j��/���j�ł���΁A�����̋x���ł͂Ȃ�
    If (���{�̏j��(vInDate) <> "") Or (Weekday(vInDate) = 1) Then
        is�����̋x�� = False
        Exit Function

    End If

    If ���{�̏j��(DateAdd("d", -1, vInDate)) <> "" And _
        ���{�̏j��(DateAdd("d", 1, vInDate)) <> "" Then
    
        is�����̋x�� = True
        Exit Function

    End If
    
    is�����̋x�� = False

End Function


'------------------------------------------------------------------
Private Function ������̋x��(vInDate As Date) As String
    
    Dim strRet As String
    strRet = ""

    Dim intday As Integer:      intday = Day(vInDate)
    Dim intMonth As Integer:    intMonth = Month(vInDate)

    ' ᱗��~�� 8/13�`15 �Ɖ��肵���ꍇ
    'If (intMonth = 8) And (13 <= intday And intday <= 15) Then strRet = "�~"

    ' �Q�n�����̓�
    'If (intMonth = 10) And (intday = 28) Then strRet = "�Q"
    
    ������̋x�� = strRet

End Function



