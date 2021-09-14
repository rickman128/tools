unit SettingExu;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Data.DB, Vcl.Grids,
  IniFiles,
  Vcl.DBGrids, Datasnap.DBClient, Vcl.StdCtrls;

type
	// �z�u���������\����
	TBusyoAttr = record
		iNo			: Integer;		        // No
		strBusyo	: String;		        // ����
        iMaxCount   : Integer;              // �z�u�l��
		iMinCount	: Integer;	            // �Œ�l��
	end;

	// �����o�����\����
	TMemberAttr = record
		iNo			: Integer;		// No
		strName		: String;		// ���O
        strRole     : String;       // �S��
		bSinge		: Boolean;		// �S�O�t���O	    T: �S�O�ɓ���
        bDayICU     : Boolean;      // ����ICU�t���O    T: ����ICU����
		bNight	    : Boolean;		// �h���t���O	    T: �h������
	end;


const
    COUNT_BUSYO     = 5;                    // �����̐�
// <#004> MOD start
	//COUNT_MEMBER    = 10;					// �����o��
	COUNT_MEMBER    = 12;					// �����o��
// <#004> MOD end

type
  TSettingf = class(TForm)
    PageControl: TPageControl;
    SheetBusyo: TTabSheet;
    SheetMember: TTabSheet;
    EGridBusyo: TDBGrid;
    DSBusyo: TDataSource;
    DMemBusyo: TClientDataSet;
    BWriteIni: TButton;
    EGridMember: TDBGrid;
    DSMember: TDataSource;
    DMemMember: TClientDataSet;
    procedure evtFormCreate(Sender: TObject);
    procedure evtBWriteIniClick(Sender: TObject);
  private
    { Private �錾 }
                                            // ����
    m_aryBusyo      : array[1..COUNT_BUSYO] of TBusyoAttr;
                                            // �����o
    m_aryMember     : array[1..COUNT_MEMBER] of TMemberAttr;


    procedure Init();
    procedure ReadIni();
    procedure WriteIni();
    procedure InitBusyoGrid();
    procedure InitMemberGrid();
  public
    { Public �錾 }
  end;


const
    { �f�[�^�Z�b�g�̃t�B�[���h�� }
    FLD_BUSYO_NO        = 'No';             // No
    FLD_BUSYO_BUSYO     = 'Busyo';          // ����
    FLD_BUSYO_MAX_COUNT = 'MaxCount';       // �z�u�l��
    FLD_BUSYO_MIN_COUNT = 'MinCount';       // �Œ�l��
    FLD_BUSYO_NOIDX     = 'Idx_No';         // �����Ŏg�p����index

    FLD_MEMBER_NO       = 'No';             // No
    FLD_MEMBER_NAME     = 'Name';           // ����
    FLD_MEMBER_ROLE     = 'Role';           // �z�u
    FLD_MEMBER_SINGE    = 'Singe';          // �S�O�t���O
    FLD_MEMBER_DAYICU   = 'DayICU';         // ����ICU
    FLD_MEMBER_NIGHTICU = 'NightICU';       // ���ICU
    FLD_MEMBER_NOIDX    = 'Idx_No';         // �����Ŏg�p����index

    { ini�t�@�C�� - �Z�N�V���� }
    SEC_BUSYO   = 'Busyo';                  // ����
    SEC_MEMBER  = 'Member';                 // �����o

    { ini�t�@�C�� - �����Z�N�V���� - �L�[ }
    KEY_BUSYO   = 'Busyo';                  // ����
    KEY_MAX_COUNT   = 'MaxCount';           // �z�u�l��
    KEY_MIN_COUNT   = 'MinCount';           // �Œ�l��

    { ini�t�@�C�� - �����o�Z�N�V���� - �L�[ }
    KEY_Name        = 'Name';               // ����
    KEY_ROLE        = 'Role';               // �S��
    KEY_SINGE       = 'Singe';              // �S�O�t���O
    KEY_DAYICU      = 'DayICU';             // ����ICU�t���O
    KEY_NIGHTICU    = 'NightICU';           // ���ICU�t���O


implementation

{$R *.dfm}

// ***********************************************************************
// *    event  : �t�H�[����OnCreate�C�x���g
// ***********************************************************************
procedure TSettingf.evtFormCreate(Sender: TObject);
begin

    Init();                                 // ��������
end;

// ***********************************************************************
// *    event  : �ۑ��{�^����OnClick�C�x���g
// ***********************************************************************
procedure TSettingf.evtBWriteIniClick(Sender: TObject);
begin
    WriteIni();                             // ini�t�@�C���ɏ�������

end;

// ***********************************************************************
// *    Name    : Init
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/04
// *    Memo    : ��������
// ***********************************************************************
procedure TSettingf.Init();
begin
    PageControl.ActivePageIndex := 0;       // �z�u�����V�[�g��\��

    // --------------------------------
    //  �ݒ�t�@�C���̓ǂݍ���
    // --------------------------------
    ReadIni();

    // --------------------------------
    //  �O���b�h�\�z
    // --------------------------------
    InitBusyoGrid();                        // �z�u����
    InitMemberGrid();                       // �����o

end;

// ***********************************************************************
// *    Name    : InitBusyoGrid
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/07/07
// *    Memo    : �z�u�����O���b�h�̏�������
// ***********************************************************************
procedure TSettingf.InitBusyoGrid();
var
    iCnt    : Integer;
    bActive : Boolean;
begin
    // --------------------------------
    //  �����f�[�^�Z�b�g�̏�������
    // --------------------------------
    DMemBusyo.Close;
    DMemBusyo.FieldDefs.Clear();
    // ��{�t�B�[���h�쐬
                                            // No
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_NO, ftInteger);
                                            // ����
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_BUSYO, ftString, 20);
                                            // �z�u�l��
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_MAX_COUNT, ftInteger);
                                            // �Œ�l��
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_MIN_COUNT, ftInteger);

                                            // index�p�t�B�[���h
    DMemBusyo.IndexDefs.Add(FLD_BUSYO_NOIDX, FLD_BUSYO_NO, [ixPrimary]);
    DMemBusyo.CreateDataSet();
    DMemBusyo.Close();

    for iCnt := 0 to DMemBusyo.FieldDefs.Count - 1 do
    begin
        DMemBusyo.FieldDefs[iCnt].CreateField(DMemBusyo);
    end;

    // �C���f�b�N�X�̐ݒ�
    DMemBusyo.IndexName := FLD_BUSYO_NOIDX;


    // --------------------------------
    //  �����o�[�̍s�쐬
    // --------------------------------
    bActive := DMemBusyo.Active;
    if not bActive then
    begin
        DMemBusyo.Open();
    end;

    try
        for iCnt := Low(m_aryBusyo) to High(m_aryBusyo) do
        begin
            DMemBusyo.AppendRecord([m_aryBusyo[iCnt].iNo,
                            m_aryBusyo[iCnt].strBusyo,
                            m_aryBusyo[iCnt].iMaxCount,
                            m_aryBusyo[iCnt].iMinCount]);
        end;
        DMemBusyo.CheckBrowseMode();

    finally
        if not bActive then
        begin
            DMemBusyo.Close();
        end;
    end;

    DMemBusyo.Open();

    // --------------------------------
    //  �O���b�h�̏�������
    // --------------------------------
    EGridBusyo.Columns[0].Title.Caption := 'No';
    EGridBusyo.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGridBusyo.Columns[0].ReadOnly := True;
    EGridBusyo.Columns[0].Width := 24;

    EGridBusyo.Columns[1].Title.Caption := '����';
    EGridBusyo.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGridBusyo.Columns[1].ReadOnly := True;

    EGridBusyo.Columns[2].Title.Caption := '�z�u�l��';
    EGridBusyo.Columns[2].Title.Alignment := TAlignment.taCenter;

    EGridBusyo.Columns[3].Title.Caption := '�Œ�l��';
    EGridBusyo.Columns[3].Title.Alignment := TAlignment.taCenter;

end;

// ***********************************************************************
// *    Name    : InitMemberGrid
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/07/07
// *    Memo    : �����o�O���b�h�̏�������
// ***********************************************************************
procedure TSettingf.InitMemberGrid();
var
    iCnt    : Integer;
    bActive : Boolean;
begin
    // --------------------------------
    //  �����o�f�[�^�Z�b�g�̏�������
    // --------------------------------
    DMemMember.Close;
    DMemMember.FieldDefs.Clear();
    // ��{�t�B�[���h�쐬
                                            // No
    DMemMember.FieldDefs.Add(FLD_MEMBER_NO, ftInteger);
                                            // ����
    DMemMember.FieldDefs.Add(FLD_MEMBER_NAME, ftString, 20);
                                            // �z�u
    DMemMember.FieldDefs.Add(FLD_MEMBER_ROLE, ftString, 20);
                                            // �S�O�t���O
    DMemMember.FieldDefs.Add(FLD_MEMBER_SINGE, ftBoolean);
                                            // ����ICU
    DMemMember.FieldDefs.Add(FLD_MEMBER_DAYICU, ftBoolean);
                                            // ���ICU
    DMemMember.FieldDefs.Add(FLD_MEMBER_NIGHTICU, ftBoolean);

                                            // index�p�t�B�[���h
    DMemMember.IndexDefs.Add(FLD_MEMBER_NOIDX, FLD_MEMBER_NO, [ixPrimary]);
    DMemMember.CreateDataSet();

    DMemMember.Close();

    for iCnt := 0 to DMemMember.FieldDefs.Count - 1 do
    begin
        DMemMember.FieldDefs[iCnt].CreateField(DMemMember);
    end;

    // �C���f�b�N�X�̐ݒ�
    DMemMember.IndexName := FLD_MEMBER_NOIDX;


    // --------------------------------
    //  �����o�[�̍s�쐬
    // --------------------------------
    bActive := DMemMember.Active;
    if not bActive then
    begin
        DMemMember.Open();
    end;

    try
        for iCnt := Low(m_aryMember) to High(m_aryMember) do
        begin
            DMemMember.AppendRecord([m_aryMember[iCnt].iNo,
                             m_aryMember[iCnt].strName,
                             m_aryMember[iCnt].strRole,
                             m_aryMember[iCnt].bSinge,
                             m_aryMember[iCnt].bDayICU,
                             m_aryMember[iCnt].bNight]);
        end;
        DMemMember.CheckBrowseMode();

    finally
        if not bActive then
        begin
            DMemMember.Close();
        end;
    end;

    DMemMember.Open();

    // --------------------------------
    //  �O���b�h�̏�������
    // --------------------------------
    EGridMember.Columns[0].Title.Caption := 'No';
    EGridMember.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[0].ReadOnly := True;
    EGridMember.Columns[0].Width := 24;

    EGridMember.Columns[1].Title.Caption := '����';
    EGridMember.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[1].ReadOnly := True;

    EGridMember.Columns[2].Title.Caption := '�z�u';
    EGridMember.Columns[2].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[2].ReadOnly := True;
    EGridMember.Columns[2].Width := 60;

    EGridMember.Columns[3].Title.Caption := '�S�O';
    EGridMember.Columns[3].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[3].Width := 72;
    EGridMember.Columns[3].Alignment := TAlignment.taCenter;

    EGridMember.Columns[4].Title.Caption := '����ICU';
    EGridMember.Columns[4].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[4].Width := 72;
    EGridMember.Columns[4].Alignment := TAlignment.taCenter;

    EGridMember.Columns[5].Title.Caption := '���ICU';
    EGridMember.Columns[5].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[5].Width := 72;
    EGridMember.Columns[5].Alignment := TAlignment.taCenter;

    // �t���O�̕\����True/False����Z/�~�ɕύX����
    TBooleanField(DMemMember.FieldByName(FLD_MEMBER_SINGE)).DisplayValues := '�Z;�~';
    TBooleanField(DMemMember.FieldByName(FLD_MEMBER_DAYICU)).DisplayValues := '�Z;�~';
    TBooleanField(DMemMember.FieldByName(FLD_MEMBER_NIGHTICU)).DisplayValues := '�Z;�~';

end;

// ***********************************************************************
// *    Name    : ReadIni
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/04
// *    Memo    : ini�t�@�C���ǂݍ��݁@�����o�̕����z��ɂ߂�
// ***********************************************************************
procedure TSettingf.ReadIni();
var
    iCnt    : Integer;
    ini     : TIniFile;
begin
                                            // exe�̂���p�X��ini�t�@�C��������
    ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));

    try
        { �����Z�N�V���� }
        for iCnt := Low(m_aryBusyo) to High(m_aryBusyo) do
        begin
            m_aryBusyo[iCnt].iNo := iCnt;   // No
                                            // ����
            m_aryBusyo[iCnt].strBusyo := ini.ReadString(SEC_BUSYO + IntToStr(iCnt), KEY_BUSYO, '');
                                            // �z�u�l��
            m_aryBusyo[iCnt].iMaxCount := ini.ReadInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MAX_COUNT, 0);
                                            // �Œ�l��
            m_aryBusyo[iCnt].iMinCount := ini.ReadInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MIN_COUNT, 0);
        end;

        { �����o�Z�N�V���� }
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
// *    Name    : WriteIni
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/06
// *    Memo    : ini�t�@�C����������
// ***********************************************************************
procedure TSettingf.WriteIni();
var
    iCnt    : Integer;
    ini     : TIniFile;
begin
                                            // exe�̂���p�X��ini�t�@�C��������
    ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));

    try
        DMemBusyo.First;
        { �����Z�N�V���� }
        for iCnt := Low(m_aryBusyo) to High(m_aryBusyo) do
        begin
                                            // �z�u�l��
            ini.WriteInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MAX_COUNT, DMemBusyo.FieldByName(FLD_BUSYO_MAX_COUNT).AsInteger);
                                            // �Œ�l��
            ini.WriteInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MIN_COUNT, DMemBusyo.FieldByName(FLD_BUSYO_MIN_COUNT).AsInteger);

            DMemBusyo.Next;
        end;

        { �����o�Z�N�V���� }
{ ���������݂͕ۗ�
        for iCnt := Low(m_aryMember) to High(m_aryMember) do
        begin
            m_aryMember[iCnt].iNo := iCnt;  // No
                                            // ����
            m_aryMember[iCnt].strName := ini.String(SEC_MEMBER + IntToStr(iCnt), KEY_NAME, '');
                                            // �S��
            m_aryMember[iCnt].strRole := ini.ReadString(SEC_MEMBER + IntToStr(iCnt), KEY_ROLE, '');
                                            // �S�O�t���O
            m_aryMember[iCnt].bSinge := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_SINGE, False);
                                            // ����ICU
            m_aryMember[iCnt].bDayICU := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_DAYICU, False);
                                            // ���ICU
            m_aryMember[iCnt].bNight := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_NIGHTICU, False);
        end;
 }
    finally
        ini.Free;
    end;
end;

end.
