unit SettingExu;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Data.DB, Vcl.Grids,
  IniFiles,
  Vcl.DBGrids, Datasnap.DBClient, Vcl.StdCtrls;

type
	// 配置部署属性構造体
	TBusyoAttr = record
		iNo			: Integer;		        // No
		strBusyo	: String;		        // 部署
        iMaxCount   : Integer;              // 配置人数
		iMinCount	: Integer;	            // 最低人数
	end;

	// メンバ属性構造体
	TMemberAttr = record
		iNo			: Integer;		// No
		strName		: String;		// 名前
        strRole     : String;       // 担当
		bSinge		: Boolean;		// 心外フラグ	    T: 心外に入る
        bDayICU     : Boolean;      // 日勤ICUフラグ    T: 日勤ICUする
		bNight	    : Boolean;		// 宿直フラグ	    T: 宿直する
	end;


const
    COUNT_BUSYO     = 5;                    // 部署の数
// <#004> MOD start
	//COUNT_MEMBER    = 10;					// メンバ数
	COUNT_MEMBER    = 12;					// メンバ数
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
    { Private 宣言 }
                                            // 部署
    m_aryBusyo      : array[1..COUNT_BUSYO] of TBusyoAttr;
                                            // メンバ
    m_aryMember     : array[1..COUNT_MEMBER] of TMemberAttr;


    procedure Init();
    procedure ReadIni();
    procedure WriteIni();
    procedure InitBusyoGrid();
    procedure InitMemberGrid();
  public
    { Public 宣言 }
  end;


const
    { データセットのフィールド名 }
    FLD_BUSYO_NO        = 'No';             // No
    FLD_BUSYO_BUSYO     = 'Busyo';          // 部署
    FLD_BUSYO_MAX_COUNT = 'MaxCount';       // 配置人数
    FLD_BUSYO_MIN_COUNT = 'MinCount';       // 最低人数
    FLD_BUSYO_NOIDX     = 'Idx_No';         // 内部で使用するindex

    FLD_MEMBER_NO       = 'No';             // No
    FLD_MEMBER_NAME     = 'Name';           // 氏名
    FLD_MEMBER_ROLE     = 'Role';           // 配置
    FLD_MEMBER_SINGE    = 'Singe';          // 心外フラグ
    FLD_MEMBER_DAYICU   = 'DayICU';         // 日勤ICU
    FLD_MEMBER_NIGHTICU = 'NightICU';       // 夜間ICU
    FLD_MEMBER_NOIDX    = 'Idx_No';         // 内部で使用するindex

    { iniファイル - セクション }
    SEC_BUSYO   = 'Busyo';                  // 部署
    SEC_MEMBER  = 'Member';                 // メンバ

    { iniファイル - 部署セクション - キー }
    KEY_BUSYO   = 'Busyo';                  // 部署
    KEY_MAX_COUNT   = 'MaxCount';           // 配置人数
    KEY_MIN_COUNT   = 'MinCount';           // 最低人数

    { iniファイル - メンバセクション - キー }
    KEY_Name        = 'Name';               // 氏名
    KEY_ROLE        = 'Role';               // 担当
    KEY_SINGE       = 'Singe';              // 心外フラグ
    KEY_DAYICU      = 'DayICU';             // 日勤ICUフラグ
    KEY_NIGHTICU    = 'NightICU';           // 夜間ICUフラグ


implementation

{$R *.dfm}

// ***********************************************************************
// *    event  : フォームのOnCreateイベント
// ***********************************************************************
procedure TSettingf.evtFormCreate(Sender: TObject);
begin

    Init();                                 // 初期処理
end;

// ***********************************************************************
// *    event  : 保存ボタンのOnClickイベント
// ***********************************************************************
procedure TSettingf.evtBWriteIniClick(Sender: TObject);
begin
    WriteIni();                             // iniファイルに書き込み

end;

// ***********************************************************************
// *    Name    : Init
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/04
// *    Memo    : 初期処理
// ***********************************************************************
procedure TSettingf.Init();
begin
    PageControl.ActivePageIndex := 0;       // 配置部署シートを表示

    // --------------------------------
    //  設定ファイルの読み込み
    // --------------------------------
    ReadIni();

    // --------------------------------
    //  グリッド構築
    // --------------------------------
    InitBusyoGrid();                        // 配置部署
    InitMemberGrid();                       // メンバ

end;

// ***********************************************************************
// *    Name    : InitBusyoGrid
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/07/07
// *    Memo    : 配置部署グリッドの初期処理
// ***********************************************************************
procedure TSettingf.InitBusyoGrid();
var
    iCnt    : Integer;
    bActive : Boolean;
begin
    // --------------------------------
    //  部署データセットの初期処理
    // --------------------------------
    DMemBusyo.Close;
    DMemBusyo.FieldDefs.Clear();
    // 基本フィールド作成
                                            // No
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_NO, ftInteger);
                                            // 部署
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_BUSYO, ftString, 20);
                                            // 配置人数
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_MAX_COUNT, ftInteger);
                                            // 最低人数
    DMemBusyo.FieldDefs.Add(FLD_BUSYO_MIN_COUNT, ftInteger);

                                            // index用フィールド
    DMemBusyo.IndexDefs.Add(FLD_BUSYO_NOIDX, FLD_BUSYO_NO, [ixPrimary]);
    DMemBusyo.CreateDataSet();
    DMemBusyo.Close();

    for iCnt := 0 to DMemBusyo.FieldDefs.Count - 1 do
    begin
        DMemBusyo.FieldDefs[iCnt].CreateField(DMemBusyo);
    end;

    // インデックスの設定
    DMemBusyo.IndexName := FLD_BUSYO_NOIDX;


    // --------------------------------
    //  メンバーの行作成
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
    //  グリッドの初期処理
    // --------------------------------
    EGridBusyo.Columns[0].Title.Caption := 'No';
    EGridBusyo.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGridBusyo.Columns[0].ReadOnly := True;
    EGridBusyo.Columns[0].Width := 24;

    EGridBusyo.Columns[1].Title.Caption := '部署';
    EGridBusyo.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGridBusyo.Columns[1].ReadOnly := True;

    EGridBusyo.Columns[2].Title.Caption := '配置人数';
    EGridBusyo.Columns[2].Title.Alignment := TAlignment.taCenter;

    EGridBusyo.Columns[3].Title.Caption := '最低人数';
    EGridBusyo.Columns[3].Title.Alignment := TAlignment.taCenter;

end;

// ***********************************************************************
// *    Name    : InitMemberGrid
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/07/07
// *    Memo    : メンバグリッドの初期処理
// ***********************************************************************
procedure TSettingf.InitMemberGrid();
var
    iCnt    : Integer;
    bActive : Boolean;
begin
    // --------------------------------
    //  メンバデータセットの初期処理
    // --------------------------------
    DMemMember.Close;
    DMemMember.FieldDefs.Clear();
    // 基本フィールド作成
                                            // No
    DMemMember.FieldDefs.Add(FLD_MEMBER_NO, ftInteger);
                                            // 氏名
    DMemMember.FieldDefs.Add(FLD_MEMBER_NAME, ftString, 20);
                                            // 配置
    DMemMember.FieldDefs.Add(FLD_MEMBER_ROLE, ftString, 20);
                                            // 心外フラグ
    DMemMember.FieldDefs.Add(FLD_MEMBER_SINGE, ftBoolean);
                                            // 日勤ICU
    DMemMember.FieldDefs.Add(FLD_MEMBER_DAYICU, ftBoolean);
                                            // 夜間ICU
    DMemMember.FieldDefs.Add(FLD_MEMBER_NIGHTICU, ftBoolean);

                                            // index用フィールド
    DMemMember.IndexDefs.Add(FLD_MEMBER_NOIDX, FLD_MEMBER_NO, [ixPrimary]);
    DMemMember.CreateDataSet();

    DMemMember.Close();

    for iCnt := 0 to DMemMember.FieldDefs.Count - 1 do
    begin
        DMemMember.FieldDefs[iCnt].CreateField(DMemMember);
    end;

    // インデックスの設定
    DMemMember.IndexName := FLD_MEMBER_NOIDX;


    // --------------------------------
    //  メンバーの行作成
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
    //  グリッドの初期処理
    // --------------------------------
    EGridMember.Columns[0].Title.Caption := 'No';
    EGridMember.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[0].ReadOnly := True;
    EGridMember.Columns[0].Width := 24;

    EGridMember.Columns[1].Title.Caption := '氏名';
    EGridMember.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[1].ReadOnly := True;

    EGridMember.Columns[2].Title.Caption := '配置';
    EGridMember.Columns[2].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[2].ReadOnly := True;
    EGridMember.Columns[2].Width := 60;

    EGridMember.Columns[3].Title.Caption := '心外';
    EGridMember.Columns[3].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[3].Width := 72;
    EGridMember.Columns[3].Alignment := TAlignment.taCenter;

    EGridMember.Columns[4].Title.Caption := '日勤ICU';
    EGridMember.Columns[4].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[4].Width := 72;
    EGridMember.Columns[4].Alignment := TAlignment.taCenter;

    EGridMember.Columns[5].Title.Caption := '夜間ICU';
    EGridMember.Columns[5].Title.Alignment := TAlignment.taCenter;
    EGridMember.Columns[5].Width := 72;
    EGridMember.Columns[5].Alignment := TAlignment.taCenter;

    // フラグの表示をTrue/Falseから〇/×に変更する
    TBooleanField(DMemMember.FieldByName(FLD_MEMBER_SINGE)).DisplayValues := '〇;×';
    TBooleanField(DMemMember.FieldByName(FLD_MEMBER_DAYICU)).DisplayValues := '〇;×';
    TBooleanField(DMemMember.FieldByName(FLD_MEMBER_NIGHTICU)).DisplayValues := '〇;×';

end;

// ***********************************************************************
// *    Name    : ReadIni
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/04
// *    Memo    : iniファイル読み込み　メンバの部署配列につめる
// ***********************************************************************
procedure TSettingf.ReadIni();
var
    iCnt    : Integer;
    ini     : TIniFile;
begin
                                            // exeのあるパスにiniファイルがある
    ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));

    try
        { 部署セクション }
        for iCnt := Low(m_aryBusyo) to High(m_aryBusyo) do
        begin
            m_aryBusyo[iCnt].iNo := iCnt;   // No
                                            // 部署
            m_aryBusyo[iCnt].strBusyo := ini.ReadString(SEC_BUSYO + IntToStr(iCnt), KEY_BUSYO, '');
                                            // 配置人数
            m_aryBusyo[iCnt].iMaxCount := ini.ReadInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MAX_COUNT, 0);
                                            // 最低人数
            m_aryBusyo[iCnt].iMinCount := ini.ReadInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MIN_COUNT, 0);
        end;

        { メンバセクション }
        for iCnt := Low(m_aryMember) to High(m_aryMember) do
        begin
            m_aryMember[iCnt].iNo := iCnt;  // No
                                            // 氏名
            m_aryMember[iCnt].strName := ini.ReadString(SEC_MEMBER + IntToStr(iCnt), KEY_NAME, '');
                                            // 担当
            m_aryMember[iCnt].strRole := ini.ReadString(SEC_MEMBER + IntToStr(iCnt), KEY_ROLE, '');
                                            // 心外フラグ
            m_aryMember[iCnt].bSinge := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_SINGE, False);
                                            // 日勤ICU
            m_aryMember[iCnt].bDayICU := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_DAYICU, False);
                                            // 夜間ICU
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
// *    Memo    : iniファイル書き込み
// ***********************************************************************
procedure TSettingf.WriteIni();
var
    iCnt    : Integer;
    ini     : TIniFile;
begin
                                            // exeのあるパスにiniファイルがある
    ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));

    try
        DMemBusyo.First;
        { 部署セクション }
        for iCnt := Low(m_aryBusyo) to High(m_aryBusyo) do
        begin
                                            // 配置人数
            ini.WriteInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MAX_COUNT, DMemBusyo.FieldByName(FLD_BUSYO_MAX_COUNT).AsInteger);
                                            // 最低人数
            ini.WriteInteger(SEC_BUSYO + IntToStr(iCnt), KEY_MIN_COUNT, DMemBusyo.FieldByName(FLD_BUSYO_MIN_COUNT).AsInteger);

            DMemBusyo.Next;
        end;

        { メンバセクション }
{ ★書き込みは保留
        for iCnt := Low(m_aryMember) to High(m_aryMember) do
        begin
            m_aryMember[iCnt].iNo := iCnt;  // No
                                            // 氏名
            m_aryMember[iCnt].strName := ini.String(SEC_MEMBER + IntToStr(iCnt), KEY_NAME, '');
                                            // 担当
            m_aryMember[iCnt].strRole := ini.ReadString(SEC_MEMBER + IntToStr(iCnt), KEY_ROLE, '');
                                            // 心外フラグ
            m_aryMember[iCnt].bSinge := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_SINGE, False);
                                            // 日勤ICU
            m_aryMember[iCnt].bDayICU := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_DAYICU, False);
                                            // 夜間ICU
            m_aryMember[iCnt].bNight := ini.ReadBool(SEC_MEMBER + IntToStr(iCnt), KEY_NIGHTICU, False);
        end;
 }
    finally
        ini.Free;
    end;
end;

end.
