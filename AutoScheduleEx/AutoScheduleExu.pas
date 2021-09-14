// ***********************************************************************
// *    Name    : AutoScheduleExu
// *    Create  : K.Endo    2020/12/03
// *    Memo    : �Ζ��������쐬�c�[��Ex
// *    History : AutoSchedule������2020�N���݂̎d�l�ōč쐬
// *    Memo    : Midas.dll���p�X�̒ʂ����Ƃ���ɐݒu����K�v������
// *              AutoSchedule�����ɐV�K�쐬
// *			  ���l���폜ver
// ***********************************************************************
unit AutoScheduleExu;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.StdCtrls, Vcl.WinXCalendars, Datasnap.DBClient, Vcl.ComCtrls,
  System.DateUtils, StrUtils,
  System.UITypes,
  Datasnap.Midas,
  ComObj,
  SettingExu,
  IniFiles, GraphUtil,
  Vcl.Samples.Calendar, Datasnap.Provider, System.ImageList, Vcl.ImgList,
  Vcl.Menus, Vcl.ExtCtrls, ssDBGridu;

const
	COUNT_MEMBER    = 12;					// �����o��
    COUNT_DAY       = 28;                   // �\���������
    COUNT_TERM      = 12;                   // �R���{�̃A�C�e����

type
    // �N���R���{���
    TYMCombo = record
        strTitle    : String;
        // ���t�͈͂̊J�n��
        dtStartDay  : TDateTime;
        iYear       : Integer;              // �N
        iMonth      : Integer;              // ��
        iDay        : Integer;              // ��
    end;

	// ������\����
	TDay = record
        dtDay       : TDateTime;            // ���ɂ�
        iDspDay     : Integer;              // �\��������ɂ�
		iWeekday   	: Integer;		        // �j��(DateUtils.DayMonday�`)
		bHoliday	: Boolean;		        // �x���t���O	T: �x��
        iDayICU     : Integer;              // ����ICU�̐l��No
        iNightICU   : Integer;              // ���ICU�̐l��No
                                            // �����o�̃V�t�g�l�z��
        aryMember   : array[0..COUNT_MEMBER] of Integer;
                                            // fix�t���O�z��(T: �V�t�g�m��/ F: �����_���ݒ��)
        aryFix      : array[0..COUNT_MEMBER] of Boolean;
	end;

	// �����o�����\����
	TMemberAttr = record
		iNo			: Integer;		// No
		strName		: String;		// ���O
        strRole     : String;       // �S��
		bSinge		: Boolean;		// �S�O�t���O	    T: �S�O�ɓ���
        bDayICU     : Boolean;      // ����ICU�t���O    T: ����ICU����
		bNight	    : Boolean;		// �h���t���O	    T: �h������
        iHolidayWork    : Integer;  // �x���o�Ή񐔁i�y���j���́��j
        iWeekendNight   : Integer;  // �T���h���񐔁i���y���j���̏h�j
	end;


  TMainFormf = class(TForm)
    DDSource: TDataSource;
    EGrid: TssDBGrid;
    BExec: TButton;
    BClearRandom: TButton;
    BClearAll: TButton;
    BExport: TButton;
    EComboYM: TComboBox;
    DDSet: TClientDataSet;
    ECal: TCalendar;
    Label1: TLabel;
    DlgSave: TSaveDialog;
    EGrid2: TDBGrid;
    DDSet2: TClientDataSet;
    DDSource2: TDataSource;
    BCheck: TButton;
    BExcel: TButton;
    BSettings: TButton;
    ImageList: TImageList;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    BImport: TButton;
    DlgOpen: TOpenDialog;
    PHidden: TPanel;
    Image1: TImage;
    LResult: TLabel;
    N2: TMenuItem;
    ICU1: TMenuItem;
    PFooter: TPanel;
    EGridSub: TssDBGrid;
    DDSourceSub: TDataSource;
    DDSetSub: TClientDataSet;
    BSummary: TButton;
    procedure evtFormCreate(Sender: TObject);
    procedure evtFormClose(Sender: TObject; var Action: TCloseAction);
    procedure evtEComboYMChange(Sender: TObject);
    procedure evtBExecClick(Sender: TObject);
    procedure evtBClearRandomClick(Sender: TObject);
    procedure evtBClearAllClick(Sender: TObject);
    procedure evtBExportClick(Sender: TObject);
    procedure MakeYMComboItem();
    procedure SetColumnCaption(iIndex : Integer);
    procedure SetColumnCaption2();
    function SetRandom(): Integer;
    procedure SetSummary();
    function SetStatus(iDay: Integer; iNo: Integer; iStatus: Integer): Boolean;
    procedure SetWeekDays();
    procedure SetColumnColor();
    procedure AryToDS();
    procedure ClearAryAll();
    procedure ClearAryRandom();
    procedure GetSatTouseki(iDay: Integer; var iNo: Integer);
    procedure PaintBar(iLevel: Integer);
    function GetSunMember(iDay: Integer): Integer;
    function GetInputSunMember(iDay: Integer): Integer;
    function GetDayICU(iDay: Integer): Integer;
    function GetNight(iDay: Integer): Integer;
    function GetInputNight(iDay: Integer): Integer;
    function GetDaikyu(iNo: Integer): Integer;
    function GetNextStatus(iStatus: Integer): Integer;
    function GetStatusStr(iStatus: Integer): String;
    function GetStatusID(strStatus: String): Integer;
    function GetMemberStr(iMemberNo: Integer): String;
    function GetWeekdayStr(iWeekday: Integer): String;
    function IsTakeOff(iDay: Integer; iNo: Integer): Boolean;
    function IsDonitiSyukujitu(iDay: Integer): Boolean;
    function IsSpecialHoliday(ADate: TDate; var AName: string): Boolean;
    function CountCathe(iDay: Integer): Integer;
    function ImportCSV(bDirect: Boolean): Boolean;
    procedure evtEGridDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure SetOutputFileName();
    procedure OutputCSV();
    procedure LogWrite(strLog: String);
    procedure evtEGridDblClick(Sender: TObject);
    procedure evtBCheckClick(Sender: TObject);
    procedure evtBExcelClick(Sender: TObject);
    procedure evtBSettingsClick(Sender: TObject);
    procedure evtN1Click(Sender: TObject);
    procedure evtPopupMenuPopup(Sender: TObject);
    procedure evtBImportClick(Sender: TObject);
    procedure evtPHiddenDblClick(Sender: TObject);
    procedure evtN2Click(Sender: TObject);
    procedure PHiddenClick(Sender: TObject);
    procedure evtICU1Click(Sender: TObject);
    procedure DDSouceSubDataChange(Sender: TObject; Field: TField);
    procedure evtBSummaryClick(Sender: TObject);


  private
    { Private �錾 }
                                            // �R���{�ɕ\������T�C�N���̔z��
    m_aryYM         : array[0..COUNT_TERM - 1] of TYMCombo;

    m_aryDay        : array[0..COUNT_DAY - 1] of TDay;
                                            // 0�͖��ݒ�p
    m_aryNightCnt   : array[0..COUNT_MEMBER] of Integer;
                                            // �����o�p
    m_aryMember     : array[1..COUNT_MEMBER] of TMemberAttr;

    procedure Init();
    procedure ReadIni();

  public
    { Public �錾 }
  end;

type
    // �V�t�g�ݒ�l�\����
    TStatusArr = record
        iNo         : Integer;      // No
        strName     : String;       // ����
    end;


// ***********************************************************************
// *    const
// ***********************************************************************
const
    DATE_ORIGIN     = '2021/02/21';         // �����\���̏���
    MAX_COUNT_NIGHT = 3;                    // �h���ő��
    MAX_COUNT_RETRY = 30;                   // �����_���ݒ�̃��g���C�ő��

	{ �Ζ����ݒ�l }
    ST_EMPTY    = 0;                        // ���ݒ�
	ST_NORMAL	= 1;                        // �Z
	ST_4HOUR	= 2;                        // �S
	ST_LONG		= 3;                        // ��
	ST_STAY		= 4;                        // �h
	ST_HIBAN	= 5;                        // ��
	ST_X		= 6;                        // �~
	ST_HOLIDAY	= 7;                        // �x
    ST_HAYAA    = 8;                        // ��A
    ST_HAYAB    = 9;                        // ��B
    ST_PM       = 10;                       // P

    ARY_STATUS  : array[ST_EMPTY..ST_PM] of TStatusArr = (
        (iNo: ST_EMPTY;     strName: ''),
        (iNo: ST_NORMAL;    strName: '�Z'),
        (iNo: ST_4HOUR;     strName: '�S'),
        (iNo: ST_LONG;      strName: '��'),
        (iNo: ST_STAY;      strName: '�h'),
        (iNo: ST_HIBAN;     strName: '��'),
        (iNo: ST_X;         strName: '�~'),
        (iNo: ST_HOLIDAY;   strName: '�x'),
        (iNo: ST_HAYAA;   strName: '��A'),
        (iNo: ST_HAYAB;   strName: '��B'),
        (iNo: ST_PM;   strName: '�o'));


    { �����o(�l�͕\����) }
    MEM_ALL = [1..COUNT_MEMBER];            // �����o�S���̗񋓌^
    MEM_A	= 1;				    // 
    MEM_B	= 2;				    // 
    MEM_C	= 3;				    // 
    MEM_D	= 4;				    // 
    MEM_E	= 5;				    // 
    MEM_F	= 6;				    // 
    MEM_G	= 7;				    // 
    MEM_H  	= 8;				    // 
    MEM_I	= 9;				    // 
    MEM_J	= 10;				    // 
    MEM_K	= 11;				    // 
    MEM_L    = 12;				    // 


    { ���C���f�[�^�Z�b�g�̃t�B�[���h�� }
    FLD_NO      = 'No';                     // No
    FLD_NAME    = 'Name';                   // ����
    FLD_ROLE    = 'Role';                   // �S��
    FLD_DAY     = 'Day';                    // Day1�`28
    FLD_NOIDX   = 'Idx_No';                 // �����Ŏg�p����index

    { �T�u�f�[�^�Z�b�g�̃t�B�[���h�� }
    FLD_HOLIDAYWORK = 'HolidayWork';        // �x���o�ΐ�
    FLD_WEEKENDNIGHT = 'WeekendNight';      // ���E�y���h����

    { ini�t�@�C�� - �Z�N�V���� }
    SEC_MEMBER   = 'Member';                // �����o�[

    { ini�t�@�C�� - �L�[ }
    KEY_Name        = 'Name';               // ����
    KEY_ROLE        = 'Role';               // �S��
    KEY_SINGE       = 'Singe';              // �S�O�t���O
    KEY_DAYICU      = 'DayICU';             // ����ICU�t���O
    KEY_NIGHTICU    = 'NightICU';           // ���ICU�t���O

    { �`�F�b�N���� }
    CHECK_RESULT_OK = 0;                    // OK
    CHECK_RESULT_ALERT = 1;                 // �x������
    CHECK_RESULT_ERROR = 2;                 // NG
    CHECK_RESULT_NONE = 9;                  // ���`�F�b�N


    { ���O }
    LOG_LINE    = '-----------------------------------';

    { ���b�Z�[�W }
    MSG_CONFIRM_RETRY = '���g���C���܂��������s���܂����B' + ''#13#10'' +
                        '������x�����_�����s���܂����H';
var
  MainFormf: TMainFormf;

implementation

{$R *.dfm}



// ***********************************************************************
// *    event  : �t�H�[����OnCreate�C�x���g
// ***********************************************************************
procedure TMainFormf.evtFormCreate(Sender: TObject);
begin
    LogWrite(LOG_LINE);
    LogWrite('��FormCreate');
    LogWrite(LOG_LINE);

    Init();                                 // ��������
end;

// ***********************************************************************
// *    event  : �t�H�[����OnClose�C�x���g
// ***********************************************************************
procedure TMainFormf.evtFormClose(Sender: TObject; var Action: TCloseAction);
begin
    LogWrite(LOG_LINE);
    LogWrite('��FormClose');
    LogWrite(LOG_LINE);

    // �I������
    EComboYM.Items.Clear();
    DDSet.Close();
    DDSetSub.Close();
    EGrid.Columns.Clear();
    EGridSub.Columns.Clear();
    DDSet.Fields.Clear();
    DDSetSub.Fields.Clear();
end;


// ***********************************************************************
// *    event  : �����_�����s�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBExecClick(Sender: TObject);
var
    iCnt    : Integer;
    iRetry  : Integer;
begin
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // �`�F�b�N���ʃo�[���N���A

    iCnt := 0;
    iRetry := 0;

    while iCnt < MAX_COUNT_RETRY do         // ���g���C�񐔂��J�E���g
    begin
        LogWrite(LOG_LINE);
        LogWrite('SetRandom: ' + IntToStr(iCnt + 1 + (MAX_COUNT_RETRY * iRetry)) + '���');
        LogWrite(LOG_LINE);

        if SetRandom() = 0 then             // �z��Ƀ����_���l��ݒ�
        begin
//            MessageDlg('����ɏI�����܂����B', mtInformation, [mbOk], 0);
            break;
        end
        else
        begin
            Inc(iCnt);
        end;

        if iCnt = MAX_COUNT_RETRY then      // ���g���C�ő�񐔂ɂȂ���
        begin
                                            // �܂����g���C���邩�m�F����
            if MessageDlg(MSG_CONFIRM_RETRY, mtConfirmation, mbYesNo, 0, mbYes) = mrYes then
            begin
                iCnt := 0;
                Inc(iRetry);
                continue;
            end;
        end;
    end;

    AryToDS();                              // �z��ɂ߂�����DataSet�ɃR�s�[����
end;

// ***********************************************************************
// *    event  : �����_���N���A�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBClearRandomClick(Sender: TObject);
begin
    ClearAryRandom();                       // ����͂����Z���ȊO���N���A����
    AryToDS();
end;

// ***********************************************************************
// *    event  : �S���N���A�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBClearAllClick(Sender: TObject);
begin
    ClearAryAll();
    AryToDS();
end;

// ***********************************************************************
// *    event  : �C���|�[�g�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBImportClick(Sender: TObject);
var
    bDirect : Boolean;
begin
    bDirect := False;

    { Ctrl�������Ȃ���N���b�N�Ŏ���̓C���|�[�g�ɂ��� }
    if GetAsynckeyState(VK_CONTROL) < 0 then
    begin
        bDirect := True;                    // ����̓t���OON
    end;

    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // �`�F�b�N���ʃo�[���N���A

    // ��
    // ���t�͈͂��ɑI�����Ă��炤
    if ImportCSV(bDirect) then
    begin
        AryToDS();
    end;

end;

// ***********************************************************************
// *    event  : �G�N�X�|�[�g�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBExportClick(Sender: TObject);
begin
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // �`�F�b�N���ʃo�[���N���A

    // CSV�o��
    OutputCSV();
end;

// ***********************************************************************
// *    event  : �ďW�v�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBSummaryClick(Sender: TObject);
begin
    SetSummary();
    AryToDS();
end;

// ***********************************************************************
// *    event  : �N���R���{��OnChange�C�x���g
// ***********************************************************************
procedure TMainFormf.evtEComboYMChange(Sender: TObject);
begin
    { �O���b�h�̒l���������񂷂ׂăN���A }
    ClearAryAll();
    AryToDS();

    { �O���b�h�̓��ɂ�����蒼�� }
    SetColumnCaption(EComboYM.ItemIndex);   // �J�����̃^�C�g���ݒ�
    SetWeekDays();                          // �j���Đݒ�
    SetColumnCaption2();                    // �j���J�����̃^�C�g���ݒ�
    SetColumnColor();                       // ���ɂ��J�����̔w�i�F�ݒ�

    { �t�@�C���o�͐ݒ� }
    SetOutputFileName();                    // �o�̓_�C�A���O�̃f�t�H���g�t�@�C�����ύX
end;

// ***********************************************************************
// *    event  : �O���b�h��OnDoubleClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtEGridDblClick(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
    iStatus : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // �N���b�N������index(0�I���W��)������t���擾(�Œ�񕪂�����)

    if iDay < 1 then                        // ���t��łȂ���΃X���[
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // �����oID�擾

                                            // W�N���b�N�����Z���̐ݒ�l��ύX����i�g�O���j
    iStatus := GetNextStatus(m_aryDay[iDay - 1].aryMember[iNo]);

    { �z��ɃZ�b�g���āA����͒l�Ƃ��ăt���O�����Ă� }
    m_aryDay[iDay - 1].aryMember[iNo] := iStatus;
    m_aryDay[iDay - 1].aryFix[iNo] := True;
    if (iStatus = ST_LONG) or (iStatus = ST_STAY) then
    begin
        m_aryDay[iDay - 1].iNightICU := iNo;
    end
    else
    begin
        m_aryDay[iDay - 1].iNightICU := 0;
    end;


    { �f�[�^�Z�b�g�X�V }
    EGrid.DataSource.DataSet.Edit;
    EGrid.SelectedField.AsString := GetStatusStr(iStatus);
    EGrid.DataSource.DataSet.Post;

end;

// ***********************************************************************
// *    event  : �O���b�h��OnDrawColumnCell�C�x���g
// ***********************************************************************
procedure TMainFormf.evtEGridDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
    iDay    : Integer;
    iNo     : Integer;
begin
                                            // ���ɂ��t�B�[���h�̏ꍇ
    if LeftStr(Column.FieldName, 3) = FLD_DAY then
    begin
                                            // �t�B�[���h��������ɂ��𒊏o(1�I���W��)
        iDay := StrToInt(StringReplace(Column.FieldName, FLD_DAY, '', [rfReplaceAll]));
        iNo := Column.Field.DataSet.RecNo;  // �s���烁���oID���擾

        { ICU�̐l�̃}�[�N����ɂ��� }
        if iNo = m_aryDay[iDay - 1].iDayICU then    // �z���0�I���W��
        begin
            EGrid.Canvas.Font.Color := clBlue;
        end;

        { ����͂����l�̓Z���F��ς��� }
        if m_aryDay[iDay - 1].aryFix[iNo] then
        begin
            EGrid.Canvas.Brush.Color := clWebLightYellow;
        end;
    end;

{
    // �Q�l
    // �t�B�[���h����Day����n�܂�A�\���l���u�h�v�̃Z���ɐF������
    if (LeftStr(Column.FieldName, 3) = 'Day') and
        (Column.Field.AsString = '�h') then
}


    //�`��
    EGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);

end;

// ***********************************************************************
// *    event   : �G�N�Z���o�̓{�^����OnClick�C�x���g
// *    memo    :
// ***********************************************************************
procedure TMainFormf.evtBExcelClick(Sender: TObject);
var
    MsExcel         : Variant;
    MsApplication   : Variant;
    WBook           : Variant;
    WSheet          : Variant;
    iCnt            : Integer;
    iRow            : Integer;
begin
    { Excel�N�� }
    MsExcel := CreateOleObject('Excel.Application');
    MsApplication := MsExcel.Application;
//    MsApplication.Visible := True;
    WBook := MsApplication.WorkBooks.Add;
    WSheet :=WBook.ActiveSheet;

    { �\�̃^�C�g�� }
    WSheet.Cells[1, 1].Value := DlgSave.FileName;
    WSheet.Cells[1, 1].Font.Size := 15;

    { �J�����^�C�g���o�� }
    WSheet.Cells[3, 1].Value := 'No';
    WSheet.Cells[3, 2].Value := '����';
    WSheet.Cells[3, 3].Value := '�S��';
                                            // ���ɂ�
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        WSheet.Cells[3, iCnt + 4].Value := IntToStr(m_aryDay[iCnt].iDspDay) + '��';
    end;

    { �f�[�^�o�� }
    try
        DDSet.DisableControls;

        DDSet.First;
        iRow := 4;
        while not DDSet.Eof do
        begin
                                            // NO
            WSheet.Cells[iRow, 1].Value := DDSet.FieldByName(FLD_NO).AsString;
                                            // ����
            WSheet.Cells[iRow, 2].Value := DDSet.FieldByName(FLD_NAME).AsString;
                                            // �S��
            WSheet.Cells[iRow, 3].Value := DDSet.FieldByName(FLD_ROLE).AsString;
                                            // ���ɂ�
            for iCnt := Low(m_aryDay) to High(m_aryDay) do
            begin
                WSheet.Cells[iRow, iCnt + 4].Value := DDSet.FieldByName(FLD_DAY + IntToStr(iCnt + 1)).AsString;
            end;

            DDSet.Next;
            Inc(iRow);
        end;
    finally
        DDSet.First;
        DDSet.EnableControls;
    end;

    MsApplication.Visible := True;

    //�ۑ��̊m�F���s��
    WBook.Saved := False;

end;

// ***********************************************************************
// *    event   : �ݒ�{�^����OnClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtBSettingsClick(Sender: TObject);
var
    dlg : TSettingf;
begin
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // �`�F�b�N���ʃo�[���N���A

    dlg := TSettingf.Create(Self);          // �ݒ�_�C�A���O

    dlg.ShowModal();                        // �_�C�A���O�\��

end;

// ***********************************************************************
// *    event   : �|�b�v�A�b�v��OnMenuPopup�C�x���g
// *    memo    :
// ***********************************************************************
procedure TMainFormf.evtPopupMenuPopup(Sender: TObject);
var
    iDay    : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // �N���b�N������index(0�I���W��)������t���擾(�Œ�񕪂�����)

    if iDay < 1 then                        // ���t��łȂ���΃X���[
    begin
        N1.Enabled := False;
    end
    else
    begin
        N1.Enabled := True;
    end;

end;

// ***********************************************************************
// *    event   : �|�b�v�A�b�v�̃N���A�{�^����OnClick�C�x���g
// *    memo    :
// ***********************************************************************
procedure TMainFormf.evtN1Click(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // �N���b�N������index(0�I���W��)������t���擾(�Œ�񕪂�����)

    if iDay < 1 then                        // ���t��łȂ���΃X���[
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // �����oID�擾

                                            // W�N���b�N�����Z���̐ݒ�l���N���A
    { �z��ɃZ�b�g���āA����͒l�Ƃ��ăt���O��OFF }
    m_aryDay[iDay - 1].aryMember[iNo] := ST_EMPTY;
    m_aryDay[iDay - 1].aryFix[iNo] := False;

    { �f�[�^�Z�b�g�X�V }
    EGrid.DataSource.DataSet.Edit;
    EGrid.SelectedField.AsString := '';
    EGrid.DataSource.DataSet.Post;
end;

// ***********************************************************************
// *    event   : �|�b�v�A�b�v�̌Œ�{�^����OnClick�C�x���g
// *    memo    : ���͂���Ă���l������͈����ɂ���
// ***********************************************************************
procedure TMainFormf.evtN2Click(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // �N���b�N������index(0�I���W��)������t���擾(�Œ�񕪂�����)

    if iDay < 1 then                        // ���t��łȂ���΃X���[
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // �����oID�擾

    { �����̓Z���łȂ���ΌŒ肷�� }
    if m_aryDay[iDay - 1].aryMember[iNo] <> ST_EMPTY then
    begin
        m_aryDay[iDay - 1].aryFix[iNo] := True; // ����͂Ƃ���
    end;

end;

// ***********************************************************************
// *    event   : �|�b�v�A�b�v��ICU�{�^����OnClick�C�x���g
// *    memo    : ICU�ݒ�/����
// ***********************************************************************
procedure TMainFormf.evtICU1Click(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // �N���b�N������index(0�I���W��)������t���擾(�Œ�񕪂�����)

    if iDay < 1 then                        // ���t��łȂ���΃X���[
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // �����oID�擾

    { �����̓Z���łȂ����ICU�ݒ�/�������� }
    if m_aryDay[iDay - 1].aryMember[iNo] <> ST_EMPTY then
    begin
        if m_aryDay[iDay - 1].iDayICU <> iNo then
        begin
            m_aryDay[iDay - 1].iDayICU := iNo; // ICU�ݒ�
        end
        else
        begin
            m_aryDay[iDay - 1].iDayICU := 0;   // ICU����
        end;
    end;

end;

// ���`�F�b�N���ڏC��
// ***********************************************************************
// *    event   : �`�F�b�N�{�^����OnClick�C�x���g
// *    memo    : �`�F�b�N����
// *                �@ �~�����ԓ���8�񂠂邩
// *                �A�h���Əh���̊Ԃ��P�������Ȃ��Ƃ��Ɍx��
// *                �B�}��E�֌��E�g�c�̂���2�l��������Ȃ��ƌx��
// *                �C����ICU��1���Ɉ�l���Ȃ��ƌx��
// ***********************************************************************
procedure TMainFormf.evtBCheckClick(Sender: TObject);
var
    iNo     : Integer;
    iCnt    : Integer;
    iCntX   : Integer;
    iCheckCnt   : Integer;
    iDay    : Integer;
    iResult : Integer;
    bErr    : Boolean;
begin
    iCheckCnt := 2;
    iResult := iCheckCnt;
    bErr := False;

    { �@�e�l�́~�����ԓ��ɂW�񂠂邩�`�F�b�N }
    for iNo := Low(m_aryMember) to High(m_aryMember) do
    begin
        iCntX := 0;
        for iCnt := Low(m_aryDay) to High(m_aryDay) do
        begin
            // �j���łȂ��~�̐����J�E���g
            if (not m_aryDay[iCnt].bHoliday) and (m_aryDay[iCnt].aryMember[iNo] = ST_X) then
            begin
                Inc(iCntX);
            end
            // �܂��͓y���ŁZ�����͈ȊO�̐l�ɂ��Ă���ꍇ�����ۂ͋x���Ȃ̂ŃJ�E���g
            else if (IsDonitiSyukujitu(iCnt)) and (m_aryDay[iCnt].aryMember[iNo] = ST_NORMAL) and
                (m_aryMember[iNo].strRole <> '����') then
            begin
                Inc(iCntX);
            end;
        end;
        if iCntX < 8 then
        begin
            ShowMessage('�~���W��Ȃ��F ' + GetMemberStr(iNo));
            bErr := True;
        end;
    end;

    if bErr then
    begin
        Dec(iResult);
    end;

    bErr := False;

    { �A�h���Əh���̊Ԃ��P�������Ȃ��Ƃ��Ɍx�� }
    for iNo := Low(m_aryMember) to High(m_aryMember) do
    begin
        iDay := 0;
        for iCnt := Low(m_aryDay) to High(m_aryDay) do
        begin
            // �h���̏ꍇ
            if iNo = m_aryDay[iCnt].iNightICU then
            begin
                if iDay = 0 then
                begin
                    iDay := iCnt;
                end
                else
                begin
                    if (iCnt = (iDay + 1)) or (iCnt = (iDay + 2)) then
                    begin
                        ShowMessage('�h���Əh���̊Ԃ��߂�����F ' + GetMemberStr(iNo));
                        bErr := True;
                        iDay := 0;
                        break;
                    end
                    else
                    begin
                        iDay := iCnt;
                    end;
                end;
            end;
        end;
    end;

    if bErr then
    begin
        Dec(iResult);
    end;

    bErr := False;
(*
    { �B�}��E�֌��E�g�c�̂���2�l��������Ȃ��ƌx�� }
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        if IsDonitiSyukujitu(iCnt) then
        begin
            continue;
        end;

        iCntX := 3;
        // �x�݁E�ߌ�x�`�F�b�N�@
        if IsTakeOff(iCnt, MEM_B) then
        begin
            Dec(iCntX);
        end;
        if IsTakeOff(iCnt, MEM_SEKIGUTI) then
        begin
            Dec(iCntX);
        end;
        if IsTakeOff(iCnt, MEM_K) then
        begin
            Dec(iCntX);
        end;
        if iCntX < 2 then
        begin
            ShowMessage('�A�֌�CE�A�̂���2�l�ȏオ���Ȃ�: ' + IntToStr(m_aryDay[iCnt].iDspDay) + '��');
            bErr := True;
            break;
        end;
    end;

    if bErr then
    begin
        Dec(iResult);
    end;

    bErr := False;
*)

    { �C����ICU��1���Ɉ�l���Ȃ��ƌx�� }
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        if m_aryDay[iCnt].iDayICU = 0 then
        begin
            ShowMessage('����ICU�����Ȃ�: ' + IntToStr(m_aryDay[iCnt].iDspDay) + '��');
            bErr := True;
            break;
        end;
    end;

    if bErr then
    begin
        Dec(iResult);
    end;

    bErr := False;

    { �D���̑��`�F�b�N }
    // ��


    { �`�F�b�N���ʁ@�i���Ȃ��j�΁����F���ԁi�N���[�����郌�x���j}

    if iResult = iCheckCnt then
    begin
        PaintBar(CHECK_RESULT_OK);
        LResult.Caption := '���Ȃ�';
    end
    else if iResult <= 0 then
    begin
        PaintBar(CHECK_RESULT_ERROR);
        LResult.Caption := '�N���[���K��';
    end
    else if iResult < iCheckCnt then
    begin
        PaintBar(CHECK_RESULT_ALERT);
        LResult.Caption := '����';
    end;

    Self.Repaint;                           // ���b�Z�[�W�̌�낪�`�悳��Ȃ����Ƃ�����̂ōĕ`��

    ShowMessage('�`�F�b�N���W�b�N�I��');

end;

// ***********************************************************************
// *    event  : �B���p�l����OnDoubleClick�C�x���g
// ***********************************************************************
procedure TMainFormf.evtPHiddenDblClick(Sender: TObject);
begin
    { �B���@�\�\�� }
    BExport.Visible := True;
end;

// ***********************************************************************
// *    Name    : Init
// *    Args    : Nothing
// *    Create  : K.Endo    2017/06/05
// *    Memo    : ��������
// ***********************************************************************
procedure TMainFormf.Init();
var
    bActive : Boolean;
    iCnt    : Integer;
begin

    PHidden.Color := clBtnFace;


    ReadIni();                              // �ݒ�t�@�C���ǂݍ���

    // --------------------------------
    // �N���R���{�̏�������
    // --------------------------------
    MakeYMComboItem();
    EComboYM.ItemIndex := 0;
    SetOutputFileName();                    // �t�@�C���o��Dlg�̃f�t�H���g�t�@�C�����ݒ�

    // --------------------------------
    //  ���C���f�[�^�Z�b�g�̏�������
    // --------------------------------
    DDSet.Close;
    DDSet.FieldDefs.Clear();
    // ��{�t�B�[���h�쐬
                                            // No
    DDSet.FieldDefs.Add(FLD_NO, ftInteger);
                                            // ����
    DDSet.FieldDefs.Add(FLD_NAME, ftString, 16);
                                            // �S��
    DDSet.FieldDefs.Add(FLD_ROLE, ftString, 4);

    // 28�����̃t�B�[���h���Œ�ō쐬
    for iCnt := 1 to COUNT_DAY do
    begin                                   // Day1�`28
        DDSet.FieldDefs.Add(FLD_DAY + IntToStr(iCnt), ftString, 6);
    end;
                                            // index�p�t�B�[���h
    DDSet.IndexDefs.Add(FLD_NOIDX, FLD_NO, [ixPrimary]);

    DDSet.CreateDataSet();
    DDSet.Close();

    for iCnt := 0 to DDSet.FieldDefs.Count - 1 do
    begin
        DDSet.FieldDefs[iCnt].CreateField(DDSet);
    end;

    // �C���f�b�N�X�̐ݒ�
    DDSet.IndexName := FLD_NOIDX;

    // --------------------------------
    //  �T�u�f�[�^�Z�b�g�̏�������
    // --------------------------------
    DDSetSub.Close;
    DDSetSub.FieldDefs.Clear();
                                            // No
    DDSetSub.FieldDefs.Add(FLD_NO, ftInteger);
                                            // �x���o��
    DDSetSub.FieldDefs.Add(FLD_HOLIDAYWORK, ftInteger);
                                            // ���E�x���h��
    DDSetSub.FieldDefs.Add(FLD_WEEKENDNIGHT, ftInteger);
                                            // index�p�t�B�[���h
    DDSetSub.IndexDefs.Add(FLD_NOIDX, FLD_NO, [ixPrimary]);

    DDSetSub.CreateDataSet();
    DDSetSub.Close();

    for iCnt := 0 to DDSetSub.FieldDefs.Count - 1 do
    begin
        DDSetSub.FieldDefs[iCnt].CreateField(DDSetSub);
    end;

    // �C���f�b�N�X�̐ݒ�
    DDSetSub.IndexName := FLD_NOIDX;

    // --------------------------------
    //  �����o�[�̍s�쐬
    // --------------------------------
    bActive := DDSet.Active;
    if not bActive then
    begin
        DDSet.Open();
        DDSetSub.Open();
    end;

    try
        for iCnt := 1 to COUNT_MEMBER do
        begin
            // ���C���O���b�h
            DDSet.AppendRecord([m_aryMember[iCnt].iNo,
                            m_aryMember[iCnt].strName,
                            m_aryMember[iCnt].strRole]);
            // �T�u�O���b�h
            DDSetSub.AppendRecord([0, 0]);

        end;
        DDSet.CheckBrowseMode();
        DDSetSub.CheckBrowseMode();
    finally
        if not bActive then
        begin
            DDSet.Close();
            DDSetSub.Close();
        end;
    end;


    DDSet.Open();
    DDSetSub.Open();

    // --------------------------------
    //  ���C���O���b�h�̏�������
    // --------------------------------
    EGrid.Columns[0].Title.Caption := 'No';
    EGrid.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGrid.Columns[0].ReadOnly := True;
    EGrid.Columns[0].Width := 24;

    EGrid.Columns[1].Title.Caption := '����';
    EGrid.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGrid.Columns[1].ReadOnly := True;

    EGrid.Columns[2].Title.Caption := '�S��';
    EGrid.Columns[2].Title.Alignment := TAlignment.taCenter;
    EGrid.Columns[2].Alignment := TAlignment.taCenter;
    EGrid.Columns[2].ReadOnly := True;

    // ���t�J����
    for iCnt := 1 to COUNT_DAY do
    begin
        EGrid.Columns[iCnt + 2].Title.Alignment := TAlignment.taCenter;
        EGrid.Columns[iCnt + 2].Alignment := TAlignment.taCenter;
        EGrid.Columns[iCnt + 2].Width := 26;
    end;

    SetColumnCaption(EComboYM.ItemIndex);   // �J�����̃^�C�g���ݒ�

    SetWeekDays();                          // �j���Đݒ�

    // --------------------------------
    //  �T�u�O���b�h�̏�������
    // --------------------------------
                                            // �X�N���[���o�[OFF
                                            // �i�{����ssNone�ŏc���Ȃ��ɂȂ邯��ssVertical�ŗ����Ȃ��ɂȂ�j
    TssDBGrid(EGridSub).ScrollBars := ssVertical;

    EGridSub.Columns[0].Title.Caption := 'No';
    EGridSub.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGridSub.Columns[0].Visible := False;

    EGridSub.Columns[1].Title.Caption := '�x���o��';
    EGridSub.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGridSub.Columns[1].Width := 54;

    EGridSub.Columns[2].Title.Caption := '��/�j���h��';
    EGridSub.Columns[2].Title.Alignment := TAlignment.taCenter;
    EGridSub.Columns[2].Width := 74;

    // --------------------------------
    //  �O���b�h�i�j���j�̏�������
    // --------------------------------
    DDSet2.Data := DDSet.Data;
    while not DDSet2.Eof do
    begin
        DDSet2.Edit;
        DDSet2.Delete;
    end;

    for iCnt := 0 to 2 do
    begin
        EGrid2.Columns[iCnt].Title.Caption := ' ';
        EGrid2.Columns[iCnt].Width := EGrid.Columns[iCnt].Width;
    end;

    // ���t�J����
    for iCnt := 1 to COUNT_DAY do
    begin
        EGrid2.Columns[iCnt + 2].Title.Alignment := TAlignment.taCenter;
        EGrid2.Columns[iCnt + 2].Width := 26;
    end;

    SetColumnCaption2();                    // �J�����̃^�C�g���ݒ�

    SetColumnColor();                       // �x���̗�ɐF�ݒ�(EGrid, EGrid2)

end;

// ***********************************************************************
// *    Name    : ReadIni
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/04
// *    Memo    : ini�t�@�C���ǂݍ��݁@�����o�̔z��ɂ߂�
// ***********************************************************************
procedure TMainFormf.ReadIni();
var
    iCnt    : Integer;
    ini     : TIniFile;
begin
                                            // exe�̂���p�X��ini�t�@�C��������
    ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));

    try
        for iCnt := Low(m_aryMember) to High(m_aryMember) do
        begin
            m_aryMember[iCnt].iNo := iCnt;  // No
                                            // ����
            m_aryMember[iCnt].strName := ini.ReadString(SEC_MEMBER + IntToStr(iCnt), KEY_NAME, '');
                                            // �S��
            m_aryMember[iCnt].strRole := ini.ReadString(SEC_MEMBER + IntToStr(iCnt), KEY_ROLE, '');
                                            // �S�O�t���O
            m_aryMember[iCnt].bSinge := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_SINGE, False);
                                            // ����ICU
            m_aryMember[iCnt].bDayICU := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_DAYICU, False);
                                            // ���ICU
            m_aryMember[iCnt].bNight := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_NIGHTICU, False);
        end;

    finally
        ini.Free;
    end;
end;

// ***********************************************************************
// *    Name    : MakeYMComboItem
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo  2017/06/06
// *    Memo    : �N���R���{�̃A�C�e���쐬
// ***********************************************************************
procedure TMainFormf.MakeYMComboItem();
var
    iCnt        : Integer;
    iYear       : Integer;
    iMonth      : Integer;
    iDay        : Integer;
    strTitle    : String;                   // �R���{�̃A�C�e���ɓ���閼�O
    strFormat   : String;
    dt          : TDateTime;
begin
    strFormat := '%d�N%d��%d��';

    EComboYM.Items.Clear();

    dt := StrToDate(DATE_ORIGIN);

    // ��������12�T�C�N�����̃A�C�e���𓮓I�ɍ쐬
    for iCnt := Low(m_aryYM) to High(m_aryYM) do
    begin
        // �����̔N�����擾 ex) 2017�N5��21��
        iYear := YearOf(dt);
        iMonth := MonthOf(dt);
        iDay := DayOf(dt);

        strTitle := Format(strFormat, [iYear, iMonth, iDay]);

        { �z��ɊJ�n����ۑ� }
        m_aryYM[iCnt].dtStartDay := dt;
        m_aryYM[iCnt].iYear := iYear;
        m_aryYM[iCnt].iMonth := iMonth;
        m_aryYM[iCnt].iDay := iDay;

        strTitle := strTitle + ' �` ';
        dt := IncDay(dt, COUNT_DAY - 1);    // 28����

        // �ŏI���̔N�����擾 ex) 2017�N6��17��
        iYear := YearOf(dt);
        iMonth := MonthOf(dt);
        iDay := DayOf(dt);

        strTitle := strTitle + Format(strFormat, [iYear, iMonth, iDay]);

        m_aryYM[iCnt].strTitle := strTitle; // ex) 2017�N5��21�� �` 2017�N6��17��
        EComboYM.Items.Add(strTitle);

        dt := IncDay(dt, 1);                // ���̃T�C�N��
    end;

end;

// ***********************************************************************
// *    Name    : SetColumnCaption
// *    Args    : iIndex    : ���t�R���{index
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/12
// *    Memo    : �O���b�h�̃J�����L���v�V������ݒ肷��
// ***********************************************************************
procedure TMainFormf.SetColumnCaption(iIndex : Integer);
var
    iCnt    : Integer;
    iDay    : Integer;
    dt      : TDateTime;
begin
    // ����
    dt := m_aryYM[iIndex].dtStartDay;

    // ���t�J����
    for iCnt := 1 to COUNT_DAY do
    begin
        iDay := DayOf(dt);
        EGrid.Columns[iCnt + 2].Title.Caption := IntToStr(iDay);
        dt := IncDay(dt);                   // ���̓���
    end;
end;

// ***********************************************************************
// *    Name    : SetColumnCaption2
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/12
// *    Memo    : �O���b�h�i�j���j�̃J�����L���v�V������ݒ肷��
// ***********************************************************************
procedure TMainFormf.SetColumnCaption2();
var
    iCnt    : Integer;
begin
    // ���t�J�����ɗj����ݒ�
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        EGrid2.Columns[iCnt + 3].Title.Caption := GetWeekdayStr(m_aryDay[iCnt].iWeekday);
    end;
end;

// ***********************************************************************
// *    Name    : SetWeekDays
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/06
// *    Memo    : �z��ɕ\���p���ɂ��E�j���E�x���ݒ�
// ***********************************************************************
procedure TMainFormf.SetWeekDays();
var
    iCnt    : Integer;
    strBuf  : String;
    dt      : TDateTime;
begin
                                            // �J�n�N�����擾
    dt := m_aryYM[EComboYM.ItemIndex].dtStartDay;

                                            // �����ł���������28�����ݒ肷��
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
                                            // ���ɂ�
        m_aryDay[iCnt].iDspDay := DayOf(dt);
                                            // TDateTime�̓��ɂ�
        m_aryDay[iCnt].dtDay := dt;
                                            // �j��
        m_aryDay[iCnt].iWeekDay := DayOfTheWeek(dt);
                                            // �y���̏ꍇ
        if m_aryDay[iCnt].iWeekDay in [DaySaturday, DaySunday] then
        begin
            m_aryDay[iCnt].bHoliday := False;
        end
                                            // �j���̏ꍇ
        else if IsSpecialHoliday(dt, strBuf) then
        begin
            m_aryDay[iCnt].bHoliday := True;
        end
        else                                // ����ȊO
        begin
            m_aryDay[iCnt].bHoliday := False;
        end;

        dt := IncDay(dt);                   // ������
    end;
end;

// ***********************************************************************
// *    Name    : SetColumnColor
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/08
// *    Memo    : �O���b�h�̓��ɂ��J�����̐F�ݒ�
// ***********************************************************************
procedure TMainFormf.SetColumnColor();
var
    iCnt    : Integer;
begin
    // �y���E�j���̔w�i�F���s���N�ɂ���
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        if IsDonitiSyukujitu(iCnt) then     // �y���j��

        begin
            EGrid.Columns[iCnt + 3].Color := clWebMistyRose;
            EGrid.Columns[iCnt + 3].Title.Font.Color := clRed;
            EGrid2.Columns[iCnt + 3].Title.Font.Color := clRed;
        end
        else
        begin
            EGrid.Columns[iCnt + 3].Color := clWindow;
            EGrid.Columns[iCnt + 3].Title.Font.Color := clWindowText;
            EGrid2.Columns[iCnt + 3].Title.Font.Color := clWindowText;
        end;
    end;
end;

// ***********************************************************************
// *    Name    : SetRandom
// *    Args    : Nothing
// *    Return  : Integer   : -1: �G���[
// *    Create  : K.Endo    2017/06/06
// *    Memo    : �����_���g�𖄂߂�
// ***********************************************************************
function TMainFormf.SetRandom(): Integer;
var
    iCnt        : Integer;
    iDay        : Integer;
    iNo         : Integer;
    iNo2        : Integer;
    iTouseki    : Integer;
begin
    Result := -1;
    ClearAryRandom();                       // �z��̎���͈ȊO���N���A

    { ���������Ȃ���\���̂̔z��ɂ�������߂� }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        // ----------------------------------------
        // �@�O���ɂ���ė����̃V�t�g�����܂�
        // ----------------------------------------
        if iDay = 0 then                    // 1���̏ꍇ�͑O���������`�F�b�N����
        begin
            // ���O�������Ɣ�r
        end
        else
        begin

            for iCnt := Low(m_aryDay[iDay - 1].aryMember) to High(m_aryDay[iDay - 1].aryMember) do
            begin
                // �O�����y�ȊO�̂Ƃ�
                if not (m_aryDay[iDay - 1].iWeekday in [DayFriday, DaySaturday, DaySunday]) then
                begin
                    // �O���ɏh���̐l�́A��������ɂȂ�
                    iNo := m_aryDay[iDay - 1].iNightICU;
                    SetStatus(iDay, iNo, ST_HIBAN);
                end;
            end;
        end;

        // ----------------------------------------
        // �A�y�j�E�j���͓��͂�1�l�o��
        // ----------------------------------------
        if (m_aryDay[iDay].iWeekDay = DaySaturday) or
            ((m_aryDay[iDay].iWeekDay <> DaySunday) and (m_aryDay[iDay].bHoliday)) then
        begin
                                            // ���͂œy�j�E�j���ɏo����l�������_����1�l���o
            GetSatTouseki(iDay, iTouseki);
            if iTouseki > 0 then            // ����͂Őݒ�ς݂̏ꍇ�A0���Ԃ����
            begin
                SetStatus(iDay, iTouseki, ST_NORMAL);
            end;
        end;

        // ----------------------------------------
        // �B���ݒ�̉ӏ��ɏh���𖄂߂�
        // ----------------------------------------
        // �h��
        iNo := GetInputNight(iDay);         // ����͂ŏh���������Ă���Ύ擾
        if iNo = 0 then                     // ����͂Ȃ�
        begin
            iNo := GetNight(iDay);          // �����_���őI��
            if iNo = -1 then
            begin
{$IFDEF DEBUG}
//            ShowMessage('�h���ł��郁���o�����܂���B�����𒆒f���܂��B');
{$ENDIF}
                Exit;
            end;
        end;
        m_aryDay[iDay].iNightICU := iNo;
        if SetStatus(iDay, iNo, ST_STAY) then
        begin
            Inc(m_aryNightCnt[iNo]);        // �h���񐔁{
        end;

        // ----------------------------------------
        // �C���j�͏o����l�����܂��Ă���
        // ----------------------------------------
        if m_aryDay[iDay].iWeekDay = DaySunday then
        begin
            iNo := GetInputSunMember(iDay); // ����͂œ��j�ɁZ�����Ă���l��Ԃ�
            if iNo = 0 then                 // ����͂Ȃ�
            begin
                iNo := GetSunMember(iDay);  // �����_���őI��
                if iNo = -1 then
                begin
{$IFDEF DEBUG}
//                ShowMessage('���j�o�΂ł��郁���o�����܂���B�����𒆒f���܂��B');
{$ENDIF}
                    Exit;
                end;
                SetStatus(iDay, iNo, ST_NORMAL);
            end;
        end;

        // ----------------------------------------
        // �D���ݒ�̉ӏ��ɓ���ICU�𖄂߂�i���j
        // ----------------------------------------
        // ����ICU
        iNo := GetDayICU(iDay);
        if iNo = -1 then
        begin
{$IFDEF DEBUG}
//            ShowMessage('����ICU�ł��郁���o�����܂���B�����𒆒f���܂��B');
{$ENDIF}
            Exit;
        end;

        m_aryDay[iDay].iDayICU := iNo;

        // ----------------------------------------
        // �E�c�������ݒ�̉ӏ��ɁZ�A�~�𖄂߂�
        // ----------------------------------------
        for iCnt := Low(m_aryDay[iDay].aryMember) to High(m_aryDay[iDay].aryMember) do
        begin
            // ���ݒ�̗��𖳏����ɖ��߂�
            if m_aryDay[iDay].aryMember[iCnt] = ST_EMPTY then
            begin
                // �y���j���́~
                if IsDonitiSyukujitu(iDay) then
                begin
                    SetStatus(iDay, iCnt, ST_X);
                end
                // ����ȊO�́Z
                else
                begin
                    SetStatus(iDay, iCnt, ST_NORMAL);
                end;
            end;
        end;

    end;

    // ----------------------------------------
    // �F���j�h���̑O���͂S�ɂȂ�
    // ----------------------------------------
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        if m_aryDay[iDay].iWeekday = DayFriday then
        begin
            iNo := m_aryDay[iDay].iNightICU;

            if iDay > 0 then                // �����łȂ����
            begin
                                            // �O�����S�ɂ���
                SetStatus(iDay - 1, iNo, ST_4HOUR);

                // ICU��ݒ肵����ɂS�ɏ㏑�����邽�߁A
                // ���̓���ICU�S�����ق��̐l�ɂ��炷�K�v������
                                            // ���̐l��ICU�S���̏ꍇ
                if iNo = m_aryDay[iDay - 1].iDayICU then
                begin
                    iNo2 := GetDayICU(iDay - 1);
                    m_aryDay[iDay - 1].iDayICU := iNo2;
                end;
            end;
        end;
    end;

    // ----------------------------------------
    // �G�y���ɋΖ������l�̑�x�𖄂߂�
    // ----------------------------------------

    for iNo := 1 to COUNT_MEMBER do
    begin
        iCnt := 0;
        { �x���o�΂̉񐔎擾 }
        for iDay := Low(m_aryDay) to High(m_aryDay) do
        begin
                                            // �y���j��
            if IsDonitiSyukujitu(iDay) then
            begin
                                            // �h
                if m_aryDay[iDay].aryMember[iNo] =ST_STAY then
                begin
                    Inc(iCnt);
                                            // ���E�x���h���{
                    Inc(m_aryMember[iNo].iWeekendNight);
                end
                                            // ��
                else if m_aryDay[iDay].aryMember[iNo] =ST_NORMAL then
                begin
                    // ���͋Ζ��̏ꍇ(����ICU�����́��͑�x�s�v)
                    // �����͒S���ȊO�̐l���Ɩ��ŏo�΂����Ƃ��̍l�����Ȃ�
                        if m_aryMember[iNo].strRole = '����' then
                    begin
                        Inc(iCnt);
                    end;
                end;

            end
                                            // ���j��
            else if m_aryDay[iDay].iWeekday = DayFriday then
            begin
                                            // �h
                if m_aryDay[iDay].aryMember[iNo] =ST_STAY then
                begin
                                            // ���E�x���h���{
                    Inc(m_aryMember[iNo].iWeekendNight);
                end;
            end;
        end;

        { �x���o�΂̉񐔕���x�Z�b�g }
        while iCnt > 0 do
        begin
            iDay := GetDaikyu(iNo);
            if iDay < 0 then
            begin
{$IFDEF DEBUG}
//                ShowMessage('��x�����Ȃ�( ɄD`)');
{$ENDIF}
                Exit;
            end;
            SetStatus(iDay, iNo, ST_X);
            Dec(iCnt);
        end;
    end;


    SetSummary();                           // �x���o�ΐ����ďW�v
    Result := 0;

    LogWrite(LOG_LINE);
    LogWrite('SetRandom: ����I��');
    LogWrite(LOG_LINE);

end;

// ***********************************************************************
// *    Name    : SetSummary
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2021/01/28
// *    Memo    : �W�v�l�������o�ϐ����ɃZ�b�g����iAryToDS�ŃO���b�h�ɕ\���j
// ***********************************************************************
procedure TMainFormf.SetSummary();
var
    iMember : Integer;
    iDay    : Integer;
    iCntDayTime : Integer;
    iCntNight   : Integer;
begin

    for iMember := 1 to COUNT_MEMBER do
    begin
        iCntDayTime := 0;
        iCntNight := 0;

        // ��������N���A���ďW�v���Ȃ���
        m_aryMember[iMember].iWeekendNight := 0;
        m_aryMember[iMember].iHolidayWork := 0;

        { �x���o�΁E��/�x���h���̉񐔎擾 }
        for iDay := Low(m_aryDay) to High(m_aryDay) do
        begin
                                            // �y���j��
            if IsDonitiSyukujitu(iDay) then
            begin
                                            // �h
                if m_aryDay[iDay].aryMember[iMember] =ST_STAY then
                begin
                                            // ���E�x���h���{
                    Inc(iCntNight);
                end
                                            // ������A
                else if m_aryDay[iDay].aryMember[iMember] in [ST_NORMAL, ST_HAYAA] then
                begin
                    Inc(iCntDayTime);
                end;
            end
                                            // ���j��
            else if m_aryDay[iDay].iWeekday = DayFriday then
            begin
                                            // �h
                if m_aryDay[iDay].aryMember[iMember] =ST_STAY then
                begin
                                            // ���E�x���h���{
                    Inc(iCntNight);
                end;
            end;
        end;

        m_aryMember[iMember].iWeekendNight := iCntNight;
        m_aryMember[iMember].iHolidayWork := iCntDayTime;
    end;

end;

// ***********************************************************************
// *    Name    : SetStatus
// *    Args    : iDay      : ���ɂ��z��index
// *              iNo       : �����oID
// *              iStatus   : �ݒ�l
// *    Return  : Boolean   : T: �ݒ肵��/ F: �ݒ肵�Ă��Ȃ�(����͍ς݂�����)
// *    Create  : K.Endo    2017/06/13
// *    Memo    : �z��ɐݒ�l�����߂�(����͂���Ă��Ȃ���΂��߂Ȃ�)
// ***********************************************************************
function TMainFormf.SetStatus(iDay: Integer; iNo: Integer; iStatus: Integer): Boolean;
begin
    Result := False;

    if not m_aryDay[iDay].aryFix[iNo] then  // ����͂���Ă��Ȃ����
    begin
                                            // �w�肳�ꂽ�ݒ�l�����߂�
        m_aryDay[iDay].aryMember[iNo] := iStatus;
        Result := True;
    end;
end;

// ***********************************************************************
// *    Name    : AryToDS
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/08
// *    Memo    : �z�񂩂�f�[�^�Z�b�g�ɃR�s�[����
// ***********************************************************************
procedure TMainFormf.AryToDS();
var
    iCnt    : Integer;
    iDay    : Integer;
    iStatus : Integer;
begin

    try
        DDSet.DisableControls;
        DDSet.First;

        DDSetSub.DisableControls;
        DDSetSub.First;

        // m_aryDay�\���̔z�񂩂�f�[�^�Z�b�g�ɃR�s�[����
        for iCnt := 1 to COUNT_MEMBER do
        begin
            for iDay := Low(m_aryDay) to High(m_aryDay) do
            begin
                iStatus := m_aryDay[iDay].aryMember[iCnt];

                DDSet.Edit;
                DDSet.FieldValues[FLD_DAY + IntToStr(iDay + 1)] := GetStatusStr(iStatus);
                DDSet.Post;
            end;

            DDSetSub.Edit;
            DDSetSub.FieldValues[FLD_HOLIDAYWORK] := m_aryMember[iCnt].iHolidayWork;
            DDSetSub.FieldValues[FLD_WEEKENDNIGHT] := m_aryMember[iCnt].iWeekendNight;
            DDSetSub.Post;

            DDSet.Next;
            DDSetSub.Next;
        end;
    finally
        DDSet.First;
        DDSetSub.First;
        DDSet.EnableControls;
        DDSetSub.EnableControls;
                                            // �X�N���[���o�[OFF
                                            // �f�[�^�Z�b�g�̃J�[�\���������ƃX�N���[���o�[���܂��\�������
                                            // ����ɁAScrollBars�v���p�e�B����x�ύX���Ȃ��ƃX�N���[���o�[�������Ȃ�
        TssDBGrid(EGridSub).ScrollBars := ssNone;
                                            // �i�{����ssNone�ŏc���Ȃ��ɂȂ邯��ssVertical�ŗ����Ȃ��ɂȂ�j
        TssDBGrid(EGridSub).ScrollBars := ssVertical;
    end;
end;

// ***********************************************************************
// *    Name    : ClearAll
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/08
// *    Memo    : �z������ׂăN���A����
// ***********************************************************************
procedure TMainFormf.ClearAryAll();
var
    iCnt    : Integer;
    iDay    : Integer;
begin
    { ���t�z��N���A }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        for iCnt := 1 to COUNT_MEMBER do
        begin
            m_aryDay[iDay].aryMember[iCnt] := ST_EMPTY;
            m_aryDay[iDay].aryFix[iCnt] := False;
        end;
        m_aryDay[iDay].iDayICU := 0;
        m_aryDay[iDay].iNightICU := 0;
    end;

    { �h���񐔔z��N���A }
    for iCnt := Low(m_aryNightCnt) to High(m_aryNightCnt) do
    begin
        m_aryNightCnt[iCnt] := 0;
        m_aryMember[iCnt].iHolidayWork := 0;
        m_aryMember[iCnt].iWeekendNight := 0;
    end;

    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // �`�F�b�N���ʃo�[���N���A
end;

// ***********************************************************************
// *    Name    : ClearAryRandom
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/13
// *    Memo    : �z�񂩂����͂����Z���ȊO���N���A����
// ***********************************************************************
procedure TMainFormf.ClearAryRandom();
var
    iCnt    : Integer;
    iDay    : Integer;
begin
    { ���t�z��N���A }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        for iCnt := 1 to COUNT_MEMBER do
        begin
            // ����͂����Ƃ���ȊO
            SetStatus(iDay, iCnt, ST_EMPTY);
        end;
        m_aryDay[iDay].iDayICU := 0;
        m_aryDay[iDay].iNightICU := 0;      // ���f������̂ł�������N���A
    end;

    { �h���񐔔z��N���A }
    for iCnt := Low(m_aryNightCnt) to High(m_aryNightCnt) do
    begin
        m_aryNightCnt[iCnt] := 0;
    end;

    { �h������͂���Ă����ꍇ�A�h���S�����ォ��⊮ }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        for iCnt := 1 to COUNT_MEMBER do
        begin
            if m_aryDay[iDay].aryMember[iCnt] = ST_STAY then
            begin
                m_aryDay[iDay].iNightICU := iCnt;
                Inc(m_aryNightCnt[iCnt]);   // �h���񐔁{
            end;
        end;
    end;

    SetSummary();                           // �x���o�ΐ����ďW�v
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // �`�F�b�N���ʃo�[���N���A
end;


// ***********************************************************************
// *    Name    : GetNextStatus
// *    Args    : iStatus   : �X�e�[�^�XID
// *    Return  : Integer   : ���̃X�e�[�^�XID(�g�O��)
// *    Create  : K.Endo    2017/06/13
// *    Memo    : �X�e�[�^�XID���玟�̃X�e�[�^�XID��Ԃ�
// ***********************************************************************
function TMainFormf.GetNextStatus(iStatus: Integer): Integer;
var
    iNext   : Integer;
begin
    if iStatus = High(ARY_STATUS) then
    begin
        iNext := ST_NORMAL;
    end
    else
    begin
        iNext := iStatus + 1;
    end;

    Result := ARY_STATUS[iNext].iNo;
end;

// ***********************************************************************
// *    Name    : GetStatusID
// *    Args    : String    : �\���p������(�Z�A��A�~�Ȃ�)
// *    Return  : iStatus   : �X�e�[�^�XID
// *    Create  : K.Endo    2017/10/27
// *    Memo    : �\���p�����񂩂�X�e�[�^�XID��Ԃ�
// ***********************************************************************
function TMainFormf.GetStatusID(strStatus: String): Integer;
var
    iCnt    : Integer;
begin
    Result := ST_EMPTY;

    if strStatus = '��' then                // �������L���ɒu��������(�C���|�[�g�Ή�)
    begin
        strStatus := '�Z';
    end
    else if strStatus = '�x��' then
    begin
        strStatus := '�x';
    end
    else if strStatus = '4' then
    begin
        strStatus := '�S';
    end;
    // ����A�A��B�AP���ϊ����K�v�H

    for iCnt := Low(ARY_STATUS) to High(ARY_STATUS) do
    begin
        if strStatus = ARY_STATUS[iCnt].strName then
        begin
            Result := ARY_STATUS[iCnt].iNo;
            Exit;
        end;
    end;

end;

// ***********************************************************************
// *    Name    : GetStatusStr
// *    Args    : iStatus   : �X�e�[�^�XID
// *    Return  : String    : �\���p������(�Z�A��A�~�Ȃ�)
// *    Create  : K.Endo    2017/06/08
// *    Memo    : �X�e�[�^�XID����\���p�������Ԃ�
// ***********************************************************************
function TMainFormf.GetStatusStr(iStatus: Integer): String;
var
    iCnt    : Integer;
begin
    Result := '';

    for iCnt := Low(ARY_STATUS) to High(ARY_STATUS) do
    begin
        if iStatus = ARY_STATUS[iCnt].iNo then
        begin
            Result := ARY_STATUS[iCnt].strName;
            Exit;
        end;
    end;

end;

// ***********************************************************************
// *    Name    : GetMemberStr
// *    Args    : iMemberNo : �����o�[No
// *    Return  : String    : ����
// *    Create  : K.Endo    2017/06/09
// *    Memo    : �����o�[No���玁����Ԃ�
// ***********************************************************************
function TMainFormf.GetMemberStr(iMemberNo: Integer): String;
var
    iCnt    : Integer;
begin
    Result := '';

    for iCnt := Low(m_aryMember) to High(m_aryMember) do
    begin
        if iMemberNo = m_aryMember[iCnt].iNo then
        begin
            Result := m_aryMember[iCnt].strName;
            Exit;
        end;
    end;

end;

// ***********************************************************************
// *    Name    : GetWeekdayStr
// *    Args    : iWeekday  : �j����`(DayMonday,�`�j
// *    Return  : String    : �j��
// *    Create  : K.Endo    2017/06/09
// *    Memo    : �j����`�l����j���̕������Ԃ�
// ***********************************************************************
function TMainFormf.GetWeekdayStr(iWeekday: Integer): String;
begin
    Result := '';

    case iWeekday of
        DayMonday:      Result := '��';
        DayTuesday:     Result := '��';
        DayWednesday:   Result := '��';
        DayThursday:    Result := '��';
        DayFriday:      Result := '��';
        DaySaturday:    Result := '�y';
        DaySunday:      Result := '��';
    end;

end;

// ***********************************************************************
// *    Name    : GetSunMember
// *    Args    : iDay      :
// *    Return  : Integer   : �����oNo
// *    Create  : K.Endo    2017/06/09
// *    Memo    : ���j�ɏo�郁���oNo��ԋp����
// ***********************************************************************
function TMainFormf.GetSunMember(iDay: Integer): Integer;
var
    iRandom : Integer;
    sTrg    : set of 0..COUNT_MEMBER;       // �Ώێ҃��X�g
begin
    iRandom := 0;

                                            // ���j�ɏo��l
    sTrg := MEM_ALL - [MEM_I];

    while not (iRandom in sTrg) do
    begin
        if sTrg = [] then
        begin
            LogWrite('GetSunMember: ���j�o�΂ł���l�����Ȃ�');
            LogWrite(IntToStr(m_aryDay[iDay].iDspDay) + '��');
            Result := -1;
            Exit;
        end;
                                            // �����_���擾
        iRandom := Random(COUNT_MEMBER + 1);

        if not (iRandom in sTrg) then       // �����_���l���Ώێ҃��X�g�ɂȂ�
        begin
            continue;
        end;

                                            // ���łɁ~�A�x�A�h�̐l�͏���
        if m_aryDay[iDay].aryMember[iRandom] in [ST_X, ST_HOLIDAY, ST_STAY] then
        begin
            sTrg := sTrg - [iRandom];       // �Ώێ҃��X�g����O��
            continue;
        end;

        // ���y���j���̏h���񐔂������l�͏��O����
        if m_aryMember[iRandom].iWeekendNight = MAX_COUNT_NIGHT then
        begin
            sTrg := sTrg - [iRandom];       // �Ώێ҃��X�g����O��
            continue;
        end;

    end;

    Result := iRandom;

end;

// ***********************************************************************
// *    Name    : GetInputSunMember
// *    Args    : iDay      : ���t�z��index
// *    Return  : Integer   : �����oNo(0: ����͂Ȃ�)
// *    Create  : K.Endo    2017/06/09
// *    Memo    : ����͂œ��j�ɁZ�������Ă��郁���oNo��ԋp����
// ***********************************************************************
function TMainFormf.GetInputSunMember(iDay: Integer): Integer;
var
    iCnt    : Integer;
begin
    Result := 0;
    for iCnt := Low(m_aryMember) to High(m_aryMember) do
    begin
        if m_aryDay[iDay].aryMember[iCnt] = ST_NORMAL then
        begin
            Result := iCnt;
            Exit;
        end;
    end;
end;

// ***********************************************************************
// *    Name    : GetSatTouseki
// *    Args    : iDay  : ���ɂ��z��index
// *              iNo   : ���̓����o�i�[��
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/09
// *    Memo    : �y�j�E�j���̓��̓����o�������_����1�l���o����
// ***********************************************************************
procedure TMainFormf.GetSatTouseki(iDay: Integer; var iNo: Integer);
var
    iCnt    : Integer;
    iRandom : Integer;
    sTrg    : set of 0..COUNT_MEMBER;       // �Ώێ҃��X�g
    procedure IgnoreMember(iMember: Integer; var iFixCnt: Integer);
    begin
        { ����͂œy�j�E�j���ɁZ�����Ă��郁���o�͏��O���� }
        if m_aryDay[iDay].aryMember[iMember] = ST_NORMAL then
        begin
            sTrg := sTrg - [iMember];
            Inc(iFixCnt);
        end;

        { ���y���j���̏h���񐔂����������o�����O���� }
        if m_aryMember[iMember].iWeekendNight = MAX_COUNT_NIGHT then
        begin
            sTrg := sTrg - [iMember];       // �Ώێ҃��X�g����O��
            Inc(iFixCnt);
        end;
    end;
begin
    iCnt := 0;
    iRandom := 0;
    iNo := 0;

                                            // �y�j�ɏo�铧�̓����o
    sTrg := [MEM_J, MEM_G, MEM_K, MEM_L];

    { ����͍ς݂܂��͋x���o�΂������l�����O���� }
    IgnoreMember(MEM_J, iCnt);
    IgnoreMember(MEM_G, iCnt);
    IgnoreMember(MEM_K, iCnt);
    IgnoreMember(MEM_L, iCnt);

    if iCnt = 1 then                        // ����͂�1�l���Ɋm�肵�Ă���
    begin
        Exit;
    end;

    // �����_���Ń����o�ݒ�
    while not (iRandom in sTrg) do
    begin
        iRandom := Random(COUNT_MEMBER + 1);
    end;
    iNo := iRandom;

end;

// ***********************************************************************
// *    Name    : GetNight
// *    Args    : iDay  : ���ɂ�
// *    Return  : �h�����郁���oNo
// *    Create  : K.Endo    2017/06/06
// *    Memo    : �h�����郁���oNo��ԋp����
// ***********************************************************************
function TMainFormf.GetNight(iDay: Integer): Integer;
var
    iRandom : Integer;
    iCnt    : Integer;
    sIgnore : set of 0..COUNT_MEMBER;       // ���O���X�g
begin
    iRandom := 0;
    sIgnore := [0];

    while iRandom in sIgnore do
    begin
        if sIgnore = [0..COUNT_MEMBER] then
        begin
            Result := -1;
            LogWrite('GetNight: �h���ł���l�����Ȃ��B ' + IntToStr(m_aryDay[iDay].iDspDay) + '��');
            Exit;
        end;
                                            // �����o��No�������_���őI��
        iRandom := Random(COUNT_MEMBER + 1);

        if iRandom in sIgnore then          // ���O���X�g��No�̏ꍇ�͂�蒼��
        begin
            continue;
        end;

        // �����͓��͂ŏo���l���h�����Ă���ۂ�
        // �Z�A��ԁA�~�̐l�̏ꍇ(���͂̐l�͓y�j�ɁZ�������Ƃ����邪�h�����Ȃ��j
{        if m_aryDay[iDay].aryMember[iRandom] in  [ST_NORMAL, ST_HIBAN, ST_4HOUR, ST_X, ST_HOLIDAY] then
        begin
            sIgnore := sIgnore + [iRandom];
            continue;
        end;
}
        // �h�����Ȃ��l���������蒼��
        if not m_aryMember[iRandom].bNight then
        begin
            sIgnore := sIgnore + [iRandom];
            continue;
        end;

        // ���\�Ȕ͈͂��H���p���Ă݂Ă��猩����
        // �O���`4���O�܂ł��h���̏ꍇ�͂�蒼���i���̏h���܂ł�4�����������j
        for iCnt := 1 to 4 do
        begin
            if (iDay - iCnt) >= Low(m_aryDay) then
            begin
                if m_aryDay[iDay - iCnt].iNightICU = iRandom then
                begin
                    sIgnore := sIgnore + [iRandom];
                    continue;
                end;
            end;
        end;

        // 3��h�����Ă���l�͑ΏۊO�ɂȂ�
        if m_aryNightCnt[iRandom] = MAX_COUNT_NIGHT then
        begin
            sIgnore := sIgnore + [iRandom];
            continue;
        end;

        // �����͏h���Ɋւ���l�̏������Ȃ��ۂ�
        { ���j���ɂ̏h����NG�̂��߂�蒼�� }
{        if iRandom = MEM_F then
        begin
            if m_aryDay[iDay].iWeekday = DayMonday then
            begin
                sIgnore := sIgnore + [iRandom];
                continue;
            end;
        end;
}
        {TODO: ���̑��h�������o�����肷�����}

    end;


    Result := iRandom;
end;

// ***********************************************************************
// *    Name    : GetInputNight
// *    Args    : iDay  : ���ɂ�
// *    Return  : �h�����郁���oNo (0: �h���Ȃ�)
// *    Create  : K.Endo    2017/06/15
// *    Memo    : ����͂ŏh���ɐݒ肳�ꂽ�����oNo��Ԃ�
// ***********************************************************************
function TMainFormf.GetInputNight(iDay: Integer): Integer;
var
    iCnt    : Integer;
begin
    Result := 0;
    for iCnt := Low(m_aryMember) to High(m_aryMember) do
    begin
        if m_aryDay[iDay].aryMember[iCnt] = ST_STAY then
        begin
            Result := iCnt;
            Exit;
        end;
    end;
end;

// ***********************************************************************
// *    Name    : GetDayICU
// *    Args    : iDay  : ���ɂ�
// *    Return  : ����ICU���郁���oNo
// *    Create  : K.Endo    2017/06/08
// *    Memo    : ����ICU���郁���oNo��ԋp����
// ***********************************************************************
function TMainFormf.GetDayICU(iDay: Integer): Integer;
var
    iRandom : Integer;
    sTrg    : set of 0..COUNT_MEMBER;       // �Ώێ҃��X�g
begin
    iRandom := 0;

    // ����ICU����l
    sTrg := MEM_ALL - [MEM_I];


    { ���O�������}

    { �j���̏��� }
    // �y���j���ȊO��
    if not IsDonitiSyukujitu(iDay) then
    begin
        // ��2�l�ȉ���1.5�l�ȉ� �݂�ȓ��͌Œ肶��Ȃ����ǂǂ�����ĎZ�o����H
        // �����́E���w�E�J�e�EPM�E�S�O�𔼓��P�ʂŐݒ肷��H
        // ���͂�2�l�ȉ��̂Ƃ���ICU�ɂ܂킳�Ȃ�
{        if (m_aryDay[iDay].aryMember[MEM_A] in [ST_4HOUR, ST_HIBAN, ST_X, ST_HOLIDAY]) or
            (m_aryDay[iDay].aryMember[MEM_J] in [ST_4HOUR, ST_HIBAN, ST_X, ST_HOLIDAY]) or
            (m_aryDay[iDay].aryMember[MEM_G] in [ST_4HOUR, ST_HIBAN, ST_X, ST_HOLIDAY]) then
        begin
            LogWrite('GetDayICu: ' + '���̓����o�����O ' + IntToStr(m_aryDay[iDay].iDspDay) + '��');
                                                // ���̓����o��S�����O
            sTrg := sTrg - [MEM_A, MEM_J, MEM_G];
        end;
}
    end;

    case m_aryDay[iDay].iWeekDay of
        DayTuesday:                         // �y�Ηj���z�S�O�̐l�����O����
        begin
            sTrg := sTrg - [MEM_B, MEM_D, MEM_L];
        end;
        DayWednesday:                       // �y���j���z
        begin
                                            // 2���A8���ȊO�͂����O����
            if not (MonthOf(m_aryDay[iDay].dtDay) in [2, 8]) then
            begin
                sTrg := sTrg - [MEM_G];
            end;
        end;
    end;


    { �c�����l����I�� }
    while not (iRandom in sTrg) do
    begin
        if sTrg = [] then
        begin
            LogWrite('GetDayICU: ����ICU�ł���l�����Ȃ� ' + IntToStr(m_aryDay[iDay].iDspDay) + '��');
            Result := -1;
            Exit;
        end;

        { ����ICU�����o�����肷����� }

                                            // �����o��No�������_���őI��
        iRandom := Random(COUNT_MEMBER + 1);

        if not (iRandom in sTrg) then       // �Ώێ҃��X�g��No�ȊO�͂�蒼��
        begin
            continue;
        end;

        { �y���j���Ŗ��ݒ�̐l�A�h���̐l�͏��O����(���̎��_�ŁZ�����Ă���l���ΏۂɂȂ�j }
        if IsDonitiSyukujitu(iDay) then
        begin
            if m_aryDay[iDay].aryMember[iRandom] in [ST_EMPTY, ST_STAY] then
            begin
                sTrg := sTrg - [iRandom];
                continue;
            end;
        end;

        { ��ԁA�~�A�x�̐l�̏ꍇ(�����_���̏ꍇ�͂܂��ݒ肳��Ă��Ȃ��B����͂��ꂽ�Ƃ��p) }
        if m_aryDay[iDay].aryMember[iRandom] in [ST_HIBAN, ST_4HOUR, ST_X, ST_HOLIDAY] then
        begin
            sTrg := sTrg - [iRandom];
            continue;
        end;
    end;

    // ex) 1��(��) xxxx �h
    LogWrite(IntToStr(m_aryDay[iDay].iDspDay) + '(' + GetWeekdayStr(m_aryDay[iDay].iWeekDay) + ') ' +
                GetMemberStr(iRandom) + ' ' +  GetStatusStr(m_aryDay[iDay].aryMember[iRandom]));

    Result := iRandom;
end;

// ***********************************************************************
// *    Name    : GetDaikyu
// *    Args    : iNo       : �����oNo
// *    Return  : Integer   : ��x���Ƃ����index
// *    Create  : K.Endo    2017/06/08
// *    Memo    : ��x���Ƃ�����index��Ԃ�
// ***********************************************************************
function TMainFormf.GetDaikyu(iNo: Integer): Integer;
var
    iRandom     : Integer;
//��    iHibanAke   : Integer;
    iCathe      : Integer;
    sTrg        : set of 0..COUNT_DAY - 1;  // �\��������t���̃��X�g

    // ***********************************************************************
    // *    Name    : GetHibanAke
    // *    Args    : Nothing
    // *    Return  : Integer   : ��Ԃ̗����̓��tindex
    // *    Create  : K.Endo    2017/06/15
    // *    Memo    : �Ώێ҃��X�g�����Ԃ̗����̓��tindex��Ԃ�
    // ***********************************************************************
    function FindHibanAke(): Integer;
    var
        iCnt    : Integer;
    begin
        Result := 0;
        for iCnt := Low(m_aryDay) to High(m_aryDay) do
        begin
                                            // ���
            if (m_aryDay[iCnt].aryMember[iNo] = ST_HIBAN) and
                                            // �������Ώۓ��t�͈͓�
                ((iCnt + 1) in sTrg) then
            begin
                Result := iCnt + 1;
                break;
            end;
        end;
    end;

begin
    Result := -1;
    sTrg := [0..COUNT_DAY - 1];

    while True do
    begin
        if sTrg = [] then                   // ��x���Ƃ������Ȃ�
        begin
            Exit;
        end;

        { �ł��邾����Ԃ̌�Ɂ~������ }
{������͓���Ȃ��d�l
        iHibanAke := FindHibanAke();        // �Ώۃ��X�g�̒������Ԃ̗�����Ԃ�
        if iHibanAke > 0 then               // ��Ԃ̗������Ώۃ��X�g���猩�������ꍇ
        begin
            iRandom := iHibanAke;           // ��Ԗ���
        end
        else
        begin
}
                                            // ���ɂ��z���index�������_���ŕԂ�(0�I���W��)
            iRandom := Random(High(m_aryDay) + 1);
//��        end;

        if not (iRandom in sTrg) then       // ���ɏ��O����Ă����
        begin
            continue;
        end;

        { �j���ˑ��̏��� }
        if IsDonitiSyukujitu(iRandom) then  // �y�y���j���z���������蒼��
        begin
            sTrg := sTrg - [iRandom];
            continue;
        end;

        Case m_aryDay[iRandom].iWeekday of
            DayMonday:                      // �y���j���z�A���x�߂Ȃ�
            begin
                if iNo in [MEM_B, MEM_L] then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;

            DayTuesday:                     // �y�Ηj���z�S���x��NG
            begin
                sTrg := sTrg - [iRandom];
                continue;
            end;

            DayWednesday:                   // �y���j���z2���A8���ȊO���x�߂Ȃ�
            begin
                if iNo = MEM_G then
                begin
                    if not (MonthOf(m_aryDay[iRandom].dtDay) in [2, 8]) then
                    begin
                        sTrg := sTrg - [iRandom];
                        continue;
                    end;
                end;
            end;
        end;
                                            // �y�ؗj���z�J�e3�l
        if m_aryDay[iRandom].iWeekday = DayThursday then
        begin
            iCathe := 3;
        end
        else                                // �y�ؗj���ȊO�z�J�e2�l
        begin
            iCathe := 2;
        end;

        { �����o�ˑ��̏��� }
        case iNo of
            MEM_A:                   // 
            begin
                                            // �Ƃ��x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if (IsTakeOff(iRandom, MEM_B) and
                    IsTakeOff(iRandom, MEM_E)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_B:                   // 
            begin
                                            // ���x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if IsTakeOff(iRandom, MEM_C) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // �Ƃ��x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if (IsTakeOff(iRandom, MEM_A) and
                    IsTakeOff(iRandom, MEM_E)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_C:                       // 
            begin
                                            // ���x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if IsTakeOff(iRandom, MEM_B) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_K:                     // 
            begin
                                            // ���x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if IsTakeOff(iRandom, MEM_L) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;

                                            // �Ƃ��x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if (IsTakeOff(iRandom, MEM_J) and
                    IsTakeOff(iRandom, MEM_I)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_E:                     // 
            begin
                                            // ���x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if IsTakeOff(iRandom, MEM_G) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // �Ƃ��x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if (IsTakeOff(iRandom, MEM_A) and
                    IsTakeOff(iRandom, MEM_B)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // �J�e����3or2�l���Ȃ��Ƃ��͋x�߂Ȃ�
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_J:                   // 
            begin
                                            // �Ƃ��x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if (IsTakeOff(iRandom, MEM_K) and
                    IsTakeOff(iRandom, MEM_I)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_G:                   // 
            begin
                                            // ���x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if IsTakeOff(iRandom, MEM_E) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // �J�e����3or2�l���Ȃ��Ƃ��͋x�߂Ȃ�
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_F:                   // 
            begin
                                            // �J�e����3or2�l���Ȃ��Ƃ��͋x�߂Ȃ�
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_L:                   // 
            begin
                                            // ���x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if IsTakeOff(iRandom, MEM_K) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_I:                    // 
            begin
                                            // �Ƃ��x�݁A�ߌ�x�̂Ƃ��͋x�߂Ȃ�
                if (IsTakeOff(iRandom, MEM_J) and
                    IsTakeOff(iRandom, MEM_K)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_H:                    // 
            begin
                                            // �J�e����3or2�l���Ȃ��Ƃ��͋x�߂Ȃ�
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
        end;

                                            // �Z�̓���ICU���ԂłȂ���Α�x�ɂł���
        if (m_aryDay[iRandom].aryMember[iNo] = ST_NORMAL) and
            (iNo <> m_aryDay[iRandom].iDayICU) then
        begin
            Result := iRandom;
            Exit;
        end
        else
        begin
            sTrg := sTrg - [iRandom];       // NG�ȓ��̓��X�g���珜�O
        end;
    end;
end;

// ***********************************************************************
// *    Name    : IsTakeOff
// *    Args    : iDay  : ���ɂ��z��index
// *              iNo   : �����oID
// *    Return  : Boolean   : T: iNo�̃����o���~�A�x�A�S�A��/ F: ����ȊO
// *    Create  : K.Endo    2017/06/12
// *    Memo    : �w�肳�ꂽ���Ɏw�肳�ꂽ�����o���x��or�ߌ�x���ǂ�����Ԃ�
// ***********************************************************************
function TMainFormf.IsTakeOff(iDay: Integer; iNo: Integer): Boolean;
begin
    if m_aryDay[iDay].aryMember[iNo] in [ST_X, ST_HOLIDAY, ST_4HOUR, ST_HIBAN] then
    begin
        Result := True;
    end
    else
    begin
        Result := False;
    end;
end;

// ***********************************************************************
// *    Name    : CountCathe
// *    Args    : iDay  : ���ɂ��z��index
// *    Return  : �J�e���̋x�݁E�ߌ�x�łȂ��l�̐l��

// *    Create  : K.Endo    2021/01/25
// *    Memo    : �J�e���̋x�݁E�ߌ�x�łȂ��l����Ԃ�
// ***********************************************************************
function TMainFormf.CountCathe(iDay: Integer): Integer;
var
    iCnt: Integer;
begin
    iCnt := 0;

    // 
    if not IsTakeOff(iDay, MEM_E) then
    begin
        Inc(iCnt);
    end;
    // 
    if not IsTakeOff(iDay, MEM_F) then
    begin
        Inc(iCnt);
    end;
    // 
    if not IsTakeOff(iDay, MEM_G) then
    begin
        Inc(iCnt);
    end;
    // 
    if not IsTakeOff(iDay, MEM_H) then
    begin
        Inc(iCnt);
    end;

    Result := iCnt;

end;

procedure TMainFormf.DDSouceSubDataChange(Sender: TObject; Field: TField);
begin

end;

// ***********************************************************************
// *    Name    : IsDonitiSyukujitu
// *    Args    : iDay  : ���ɂ��z��index
// *    Return  : Boolean   : T: �y���j�� / F: ����
// *    Create  : K.Endo    2017/06/16
// *    Memo    : �w�肳�ꂽ�����y���j����
// ***********************************************************************
function TMainFormf.IsDonitiSyukujitu(iDay: Integer): Boolean;
begin
                                            // �y��
    if (m_aryDay[iDay].iWeekday in [DaySaturday, DaySunday]) or
                                            // �܂��͏j��
        (m_aryDay[iDay].bHoliday) then
    begin
        Result := True;
    end
    else
    begin
        Result := False;
    end;

end;

// ***********************************************************************
// *    Name    : PaintBar
// *    Args    : iLevel  : �`�F�b�N���x��
// *    Return  : Nothing
// *    Create  : K.Endo    2017/10/31
// *    Memo    : �`�F�b�N���ʃo�[�̕\���ؑ�
// ***********************************************************************
procedure TMainFormf.PaintBar(iLevel: Integer);
var
    clStart, clEnd: TColor;
begin
    case iLevel of
        CHECK_RESULT_OK:
        begin
            clStart := clGreen;
            clEnd := clGreen;
        end;
        CHECK_RESULT_ALERT:
        begin
            clStart := clGreen;
            clEnd := clYellow;
        end;
        CHECK_RESULT_ERROR:
        begin
            clStart := clYellow;
            clEnd := clRed;
        end;
        CHECK_RESULT_NONE:
        begin
            image1.Visible := False;
            Exit;
        end;
        else
        begin
            clStart := clGreen;
            clEnd := clGreen;
        end;
    end;

    image1.Visible := True;
    GradientFillCanvas(image1.Canvas,
                        clStart,
                        clEnd,
                        Rect(0, 0, 225, 19),
                        TGradientDirection.gdHorizontal);

end;

procedure TMainFormf.PHiddenClick(Sender: TObject);
begin

end;

// ***********************************************************************
// *    Name    : SetOutputFileName
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/09
// *    Memo    : �t�@�C���o�̓_�C�A���O�̃f�t�H���g�t�@�C�����ύX
// ***********************************************************************
procedure TMainFormf.SetOutputFileName();
begin
    DlgSave.FileName := '�Ζ���_' + EComboYM.Text;
end;

// ***********************************************************************
// *    Name    : ImportCSV
// *    Args    : bDirect   : T: �����/ F: �ʏ�f�[�^
// *    Return  : Nothing
// *    Create  : K.Endo    2017/10/27
// *    Memo    : CSV�C���|�[�g
// ***********************************************************************
function TMainFormf.ImportCSV(bDirect: Boolean): Boolean;
var
    CSVFile : String;
    stl     : TStringList;
    str     : String;
    iMember : Integer;
    iDay    : Integer;
    strSplitted : TStringList;
begin
    Result := False;

    // �t�@�C���̎w��
    DlgOpen.Filter := 'CSV�t�@�C�� (*.csv)|*.csv';
    if DlgOpen.Execute then
    begin
        CSVFile := DlgOpen.FileName;
    end
    else
    begin
        Exit;
    end;

    stl := TStringList.Create();
    strSplitted := TStringList.Create();

    try
        stl.LoadFromFile(CSVFile);          // StringList�Ƀt�@�C���ǂݍ���

        // �z��ɂ߂�
        for iMember := 1 to stl.Count - 1 do// ���t�̍s�͔�΂��̂łP�����OK
        begin
            // 1�s�Ƃ肾��
            str := stl[iMember];
            // �J���}�ŕ�����
            strSplitted.DelimitedText := stl[iMember];

            for iDay := Low(m_aryDay) to High(m_aryDay) do
            begin
                // �����oNo�̃J�����ɁuICU�v�������Ă���s�ɓ���ICU�̐l�̃����oNo�������Ă���
                if strSplitted[0] = 'ICU' then
                begin
                    // ����ICU
                    m_aryDay[iDay].iDayICU := StrToInt(strSplitted[iDay + 3]);
                end
                else
                begin
                    // �X�e�[�^�X�����񂩂�ID�ɂ��Ĕz��ɂ߂�
                    m_aryDay[iDay].aryMember[iMember] := GetStatusID(strSplitted[iDay + 3]);

                    // ����̓C���|�[�g�̏ꍇ�͒l�������Ă���ӏ��̎���̓t���O��ON�ɂ���
                    if bDirect then
                    begin
                        if m_aryDay[iDay].aryMember[iMember] <> ST_EMPTY then
                        begin
                            m_aryDay[iDay].aryFix[iMember] := True;
                        end;
                    end;
                end;
            end;
        end;

    finally
        stl.Free;
        strSplitted.Free;

        Result := True;
    end;
end;

// ***********************************************************************
// *    Name    : OutputCSV
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/09
// *    Memo    : CSV�o��
// ***********************************************************************
procedure TMainFormf.OutputCSV();
var
    F       : TextFile;
    CSVFile : String;
    stl     : TStringList;
    i       : Integer;
    iDay    : Integer;
begin
    // �ۑ��ꏊ�̎w��
    if DlgSave.Execute then
    begin
        CSVFile := DlgSave.FileName;
    end
    else
    begin
        Exit;
    end;

    stl := TStringList.Create();

    try

        DDSet.DisableControls;              // ������h�~

        // �t�@�C���o��
        AssignFile(F, CSVFile);             // �t�@�C���Ǝ��t�@�C�����т���
        ReWrite(F);                         // �t�@�C����V�K�쐬���ĊJ��

        DDSet.First;
        // �^�C�g��(�t�B�[���h)�s�̏o��
        for i := 0 to DDSet.FieldCount - 1 do
        begin
            stl.Add(DDSet.Fields[i].FieldName);
        end;

        Writeln(F, stl.CommaText);
        stl.Clear;

        // ���X�g�o��
        while not DDSet.Eof do
        begin
            for i := 0 to DDSet.FieldCount - 1 do
            begin
                stl.Add(DDSet.Fields[i].AsString);
            end;

            Writeln(F, stl.CommaText);      // �e�L�X�g�t�@�C����1�s�o��
            stl.Clear;

            DDSet.Next;
        end;

        // ����ICU�̐l��No���o�͂���
        stl.Add('ICU');
        stl.Add('');                        // Name
        stl.Add('');                        // Role
        for iDay := Low(m_aryDay) to High(m_aryDay) do
        begin
            stl.Add(IntToStr(m_aryDay[iDay].iDayICU));
        end;

        Writeln(F, stl.CommaText);          // �e�L�X�g�t�@�C����1�s�o��
        stl.Clear;

        CloseFile(F);                       //�t�@�C�������

    finally
        DDSet.EnableControls;               // ������h�~
        stl.Free;
    end;
end;

// ***********************************************************************
// *    Name    : LogWrite
// *    Args    : strLog    : �o�̓��O
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/09
// *    Memo    : ���O�o��
// ***********************************************************************
procedure TMainFormf.LogWrite(strLog: String);
{$IFDEF DEBUG}
var
    txtFile: TextFile;
    str, path: string;
    strDateTime : String;
{$ENDIF DEBUG}
begin
{$IFDEF DEBUG}
    strDateTime := DateTimeToStr(Now());

    str := Format('%s %s', [strDateTime, strLog]);
    path := 'C:\end\03 �J��\AutoSchedule.log';
    AssignFile(txtFile, path);

    if FileExists(path) then
    begin
        Append(txtFile);
    end
    else
    begin
        Rewrite(txtFile);
    end;

    Writeln(txtFile, str);
    CloseFile(txtFile);
{$ENDIF DEBUG}
end;


// ***********************************************************************
// *    Name    : IsSpecialHoliday
// *    Args    : ADate
// *              AName
// *    Return  : Boolean   : T: �j��/ F: �j���łȂ�
// *    Create  : K.Endo    2017/06/06
// *    Memo    : �j�����ǂ������f����
// ***********************************************************************
function TMainFormf.IsSpecialHoliday(ADate: TDate; var AName: string): Boolean;
// ------------------------------------------------------------
// http://koyomi.vis.ne.jp/
// http://www.asahi-net.or.jp/~CI5M-NMR/misc/equinox.html#Rule
// ------------------------------------------------------------
// ADate���j�����ǂ�����Ԃ��B
// �j��=True,�j���ł͂Ȃ�=False
// AName �ɂ͏j���̖��O��Ԃ�
var
    DName: string;
    i:Integer;
    {FreqOfWeek Begin}
    function FreqOfWeek(AYear, AMonth: Word; AWeekNo, ADayOfWeeek: Byte): TDateTime;
    // AYear�NAMonth���̑�AWeekNo�uADayOfWeeek�j���v�̓��t��Ԃ�
    // ADayOfWeeek�@���j��=1..�y�j��=7
    var
        dDay: Word;
        dDoW: Word;
        dWeekNo: Byte;
    begin
        dDoW := DayOfWeek(EncodeDate(AYear, AMonth, 1));
        dWeekNo := AWeekNo;
        if ADayOfWeeek >= dDoW then
        begin
            dWeekNo := dWeekNo - 1;
        end;
        dDay := (dWeekNo * 7) + (ADayOfWeeek - dDoW) + 1;
        result := EncodeDate(AYear, AMonth, dDay);
    end;
    {FreqOfWeek End}
    {LeapYearCount Begin}
    function LeapYearCount(SYear, EYear: Word): Integer;
    // SYear����EYear���ɉ���[�N�����邩��Ԃ�
    var
        i: Integer;
        Cnt: Integer;
    begin
        Cnt := 0;
        for i := sYear to eYear do
        begin
            if (i mod 4) <> 0 then
              Continue;
            if IsLeapYear(i) then
              Inc(Cnt);
        end;
        result := Cnt;
    end;
    {LeapYearCount End}
    {VernalEquinox End}
    function VernalEquinox(AYear: Word): TDateTime;
    // Ayear�̏t���̓������߂�
    var
        dDay: Word;
    begin
        dDay := Trunc((21.147 + ((AYear - 1940) * 0.2421904) - (LeapYearCount(1940, AYear) - 1)));
        result := EncodeDate(AYear, 3, dDay);
    end;
    {VernalEquinox End}
    {AutumnalEquinox End}
    function AutumnalEquinox(AYear: Word): TDateTime;
    // Ayear�̏H���̓������߂�
    var
        dDay: Word;
    begin
        dDay := Trunc((23.5412 + ((AYear - 1940) * 0.2421904) - (LeapYearCount(1940, AYear) - 1)));
        result := EncodeDate(AYear, 9, dDay);
    end;
    {AutumnalEquinox End}
    {_IsSpecialHoliday Begin}
    function _IsSpecialHoliday(ADate: TDate; var AName: string): Boolean;
    // ADate���j�����ǂ�����Ԃ��B
    // �j��=True,�j���ł͂Ȃ�=False
    // AName �ɂ͏j���̖��O��Ԃ�
    // '�����̋x��'�͂����ł͎Z�o����Ȃ�
    var
        dYear, dMonth, dDay: Word;
    begin
        AName := '';
        result := False;
        DecodeDate(ADate, dYear, dMonth, dDay);
        case dMonth of
          1:
            begin
              // '����' 1948�`
              if (dYear >= 1948) and (dDay = 1) then
                begin
                  result := True;
                  AName := '����';
                  Exit;
                end;
              // '���l�̓��@' 1948�`1999
              if (dYear >= 1948) and (dYear <= 1999) and (dDay = 15) then
                begin
                  result := True;
                  AName := '���l�̓�';
                  Exit;
                end;
              // '���l�̓��A' 2000�`
              // ��Q���j��(�n�b�s�[�}���f�[)
              if (dYear >= 2000) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 2, 2) then
                    begin
                      result := True;
                      AName := '���l�̓�';
                      Exit;
                    end;
                end;
            end;
          2:
            begin
              // '�����L�O�̓�' 1966�`
              if (dYear >= 1966) and (dDay = 11) then
                begin
                  result := True;
                  AName := '�����L�O�̓�';
                  Exit;
                end;
              // �����a�V�c�̑�r�̗�(1989/02/24)
              if (dYear = 1989) and (dDay = 24) then
                begin
                  result := True;
                  AName := '���a�V�c�̑�r�̗�';
                  Exit;
                end;
            end;
          3:
            begin
              // '�t���̓�' 1949�`
              if (dYear >= 1949) then
                begin
                  if ADate = VernalEquinox(dYear) then
                    begin
                      result := True;
                      AName := '�t���̓�';
                      Exit;
                    end;
                end;
            end;
          4:
            begin
              // '�V�c�a����' 1948�`1988
              if (dYear >= 1948) and (dYear <= 1988) and (dDay = 29) then
                begin
                  result := True;
                  AName := '�V�c�a����';
                  Exit;
                end;
              // '�݂ǂ�̓��@' 1989�`2006
              if (dYear >= 1989) and (dYear <= 2006) and (dDay = 29) then
                begin
                  result := True;
                  AName := '�݂ǂ�̓�';
                  Exit;
                end;
              // '���a�̓�' 2007�`
              if (dYear >= 2007) and (dDay = 29) then
                begin
                  result := True;
                  AName := '���a�̓�';
                  Exit;
                end;
              // ���c���q���m�e���̌����̋V(1959/04/10)
              if (dYear = 1959) and (dDay = 10) then
                begin
                  result := True;
                  AName := '�c���q���m�e���̌����̋V';
                  Exit;
                end;
            end;
          5:
            begin
              // '���@�L�O��' 1948�`
              if (dYear >= 1948) and (dDay = 3) then
                begin
                  result := True;
                  AName := '���@�L�O��';
                  Exit;
                end;
              // '�݂ǂ�̓��A' 2007�`
              if (dYear >= 2007) and (dDay = 4) then
                begin
                  result := True;
                  AName := '�݂ǂ�̓�';
                  Exit;
                end;
              // '���ǂ��̓�' 1948�`
              if (dYear >= 1948) and (dDay = 5) then
                begin
                  result := True;
                  AName := '���ǂ��̓�';
                  Exit;
                end;
            end;
          6:
            begin
              // ���c���q���m�e���̌����̋V(1993/06/09)
              if (dYear = 1993) and (dDay = 9) then
                begin
                  result := True;
                  AName := '�c���q���m�e���̌����̋V';
                  Exit;
                end;
            end;
          7:
            begin
              // '�C�̓��@' 1995�`2002
              if (dYear >= 1995) and (dYear <= 2002) and (dDay = 20) then
                begin
                  result := True;
                  AName := '�C�̓�';
                  Exit;
                end;
              // '�C�̓��A' 2003�`
              // ��R���j��
              if (dYear >= 2003) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 3, 2) then
                    begin
                      result := True;
                      AName := '�C�̓�';
                      Exit;
                    end;
                end;
            end;
          8:
            begin
              // '�R�̓�' 2016�`
              if (dYear >= 2016) and (dDay = 11) then
                begin
                  result := True;
                  AName := '�R�̓�';
                  Exit;
                end;
            end;
          9:
            begin
              // '�h�V�̓��@' 1966�`2002
              if (dYear >= 1966) and (dYear <= 2002) and (dDay = 15) then
                begin
                  result := True;
                  AName := '�h�V�̓�';
                  Exit;
                end;
              // '�h�V�̓��A' 2003�`
              // ��R���j��
              if (dYear >= 2003) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 3, 2) then
                    begin
                      result := True;
                      AName := '�h�V�̓�';
                      Exit;
                    end;
                end;
              // '�H���̓�' 1948�`
              if (dYear >= 1948) then
                begin
                  if ADate = AutumnalEquinox(dYear) then
                    begin
                      result := True;
                      AName := '�H���̓�';
                      Exit;
                    end;
                end;
            end;
          10:
            begin
              // '�̈�̓��@' 1966�`1999
              if (dYear >= 1966) and (dYear <= 1999) and (dDay = 10) then
                begin
                  result := True;
                  AName := '�̈�̓�';
                  Exit;
                end;
              // '�̈�̓��A' 2000�`
              // ��Q���j��(�n�b�s�[�}���f�[)
              if (dYear >= 2000) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 2, 2) then
                    begin
                      result := True;
                      AName := '�̈�̓�';
                      Exit;
                    end;
                end;
            end;
          11:
            begin
              // '�����̓�' 1948�`
              if (dYear >= 1948) and (dDay = 3) then
                begin
                  result := True;
                  AName := '�����̓�';
                  Exit;
                end;
              // '�ΘJ���ӂ̓�' 1948�`
              if (dYear >= 1948) and (dDay = 23) then
                begin
                  result := True;
                  AName := '�ΘJ���ӂ̓�';
                  Exit;
                end;
              // �����ʗ琳�a�̋V(1990/11/12)
              if (dYear = 1990) and (dDay = 12) then
                begin
                  result := True;
                  AName := '���ʗ琳�a�̋V';
                  Exit;
                end;
            end;
          12:
            begin
              // '�V�c�a����' 1948�`
              if (dYear >= 1989) and (dDay = 23) then
                begin
                  result := True;
                  AName := '�V�c�a����';
                  Exit;
                end;
            end;
        end;
    end;
  {_IsSpecialHoliday End}
begin
    result := False;
    AName := '';
    if _IsSpecialHoliday(ADate, DName) then
    begin
        result := True;
        AName := DName;
    end
    else if (ADate >= EncodeDate(1973, 4, 12)) and (DayOfWeek(ADate) = 2) and
          _IsSpecialHoliday(ADate - 1, DName) then
    begin
        // �U�֋x���@�@1973/04/12�ȍ~
        // ���j���Əj�Փ����d�Ȃ����ꍇ�ɂ�'�U�֋x��'�ƂȂ�
        result := True;
        AName := '�U�֋x��';
    end
    else if (ADate >= EncodeDate(1988, 5, 4)) and (DayOfWeek(ADate) <> 1) and
          _IsSpecialHoliday(ADate - 1, DName) and _IsSpecialHoliday(ADate + 1, DName) then
    begin
        // �����̋x�� 1988/05/04�ȍ~
        // �j���Əj���ɋ��܂ꂽ������'�����̋x��'�ƂȂ�B
        result := True;
        AName := '�����̋x��';
    end
    else if (ADate >= EncodeDate(2008, 5, 6)) and (DayOfWeek(ADate) <> 1) and
          _IsSpecialHoliday(ADate - DayOfWeek(ADate) + 1, DName) then
    begin
        // �U�֋x���A�@2008/05/06�ȍ~
        // '�j��'�����j���ɓ�����Ƃ��́A���̓���ɂ����Ă��̓��ɍł��߂�'�j��'�łȂ������x���Ƃ���
        result := True;
        AName := '�U�֋x��';
        for i:=1 to DayOfWeek(ADate) - 2 do
        begin
            if not _IsSpecialHoliday(ADate - i, DName) then
            begin
              result := False;
              AName := '';
              Break;
            end;
        end;
    end;
end;
end.
