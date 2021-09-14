// ***********************************************************************
// *    Name    : AutoScheduleExu
// *    Create  : K.Endo    2020/12/03
// *    Memo    : 勤務割自動作成ツールEx
// *    History : AutoScheduleを元に2020年現在の仕様で再作成
// *    Memo    : Midas.dllをパスの通ったところに設置する必要がある
// *              AutoScheduleを元に新規作成
// *			  ※個人名削除ver
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
	COUNT_MEMBER    = 12;					// メンバ数
    COUNT_DAY       = 28;                   // 表示する日数
    COUNT_TERM      = 12;                   // コンボのアイテム数

type
    // 年月コンボ情報
    TYMCombo = record
        strTitle    : String;
        // 日付範囲の開始日
        dtStartDay  : TDateTime;
        iYear       : Integer;              // 年
        iMonth      : Integer;              // 月
        iDay        : Integer;              // 日
    end;

	// 一日分構造体
	TDay = record
        dtDay       : TDateTime;            // 日にち
        iDspDay     : Integer;              // 表示する日にち
		iWeekday   	: Integer;		        // 曜日(DateUtils.DayMonday～)
		bHoliday	: Boolean;		        // 休日フラグ	T: 休日
        iDayICU     : Integer;              // 日勤ICUの人のNo
        iNightICU   : Integer;              // 夜間ICUの人のNo
                                            // メンバのシフト値配列
        aryMember   : array[0..COUNT_MEMBER] of Integer;
                                            // fixフラグ配列(T: シフト確定/ F: ランダム設定可)
        aryFix      : array[0..COUNT_MEMBER] of Boolean;
	end;

	// メンバ属性構造体
	TMemberAttr = record
		iNo			: Integer;		// No
		strName		: String;		// 名前
        strRole     : String;       // 担当
		bSinge		: Boolean;		// 心外フラグ	    T: 心外に入る
        bDayICU     : Boolean;      // 日勤ICUフラグ    T: 日勤ICUする
		bNight	    : Boolean;		// 宿直フラグ	    T: 宿直する
        iHolidayWork    : Integer;  // 休日出勤回数（土日祝日の○）
        iWeekendNight   : Integer;  // 週末宿直回数（金土日祝日の宿）
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
    { Private 宣言 }
                                            // コンボに表示するサイクルの配列
    m_aryYM         : array[0..COUNT_TERM - 1] of TYMCombo;

    m_aryDay        : array[0..COUNT_DAY - 1] of TDay;
                                            // 0は未設定用
    m_aryNightCnt   : array[0..COUNT_MEMBER] of Integer;
                                            // メンバ用
    m_aryMember     : array[1..COUNT_MEMBER] of TMemberAttr;

    procedure Init();
    procedure ReadIni();

  public
    { Public 宣言 }
  end;

type
    // シフト設定値構造体
    TStatusArr = record
        iNo         : Integer;      // No
        strName     : String;       // 文言
    end;


// ***********************************************************************
// *    const
// ***********************************************************************
const
    DATE_ORIGIN     = '2021/02/21';         // 初期表示の初日
    MAX_COUNT_NIGHT = 3;                    // 宿直最大回数
    MAX_COUNT_RETRY = 30;                   // ランダム設定のリトライ最大回数

	{ 勤務割設定値 }
    ST_EMPTY    = 0;                        // 未設定
	ST_NORMAL	= 1;                        // 〇
	ST_4HOUR	= 2;                        // ４
	ST_LONG		= 3;                        // 長
	ST_STAY		= 4;                        // 宿
	ST_HIBAN	= 5;                        // 非
	ST_X		= 6;                        // ×
	ST_HOLIDAY	= 7;                        // 休
    ST_HAYAA    = 8;                        // 早A
    ST_HAYAB    = 9;                        // 早B
    ST_PM       = 10;                       // P

    ARY_STATUS  : array[ST_EMPTY..ST_PM] of TStatusArr = (
        (iNo: ST_EMPTY;     strName: ''),
        (iNo: ST_NORMAL;    strName: '〇'),
        (iNo: ST_4HOUR;     strName: '４'),
        (iNo: ST_LONG;      strName: '長'),
        (iNo: ST_STAY;      strName: '宿'),
        (iNo: ST_HIBAN;     strName: '非'),
        (iNo: ST_X;         strName: '×'),
        (iNo: ST_HOLIDAY;   strName: '休'),
        (iNo: ST_HAYAA;   strName: '早A'),
        (iNo: ST_HAYAB;   strName: '早B'),
        (iNo: ST_PM;   strName: 'Ｐ'));


    { メンバ(値は表示順) }
    MEM_ALL = [1..COUNT_MEMBER];            // メンバ全員の列挙型
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


    { メインデータセットのフィールド名 }
    FLD_NO      = 'No';                     // No
    FLD_NAME    = 'Name';                   // 氏名
    FLD_ROLE    = 'Role';                   // 担当
    FLD_DAY     = 'Day';                    // Day1～28
    FLD_NOIDX   = 'Idx_No';                 // 内部で使用するindex

    { サブデータセットのフィールド名 }
    FLD_HOLIDAYWORK = 'HolidayWork';        // 休日出勤数
    FLD_WEEKENDNIGHT = 'WeekendNight';      // 金・土日宿直数

    { iniファイル - セクション }
    SEC_MEMBER   = 'Member';                // メンバー

    { iniファイル - キー }
    KEY_Name        = 'Name';               // 氏名
    KEY_ROLE        = 'Role';               // 担当
    KEY_SINGE       = 'Singe';              // 心外フラグ
    KEY_DAYICU      = 'DayICU';             // 日勤ICUフラグ
    KEY_NIGHTICU    = 'NightICU';           // 夜間ICUフラグ

    { チェック結果 }
    CHECK_RESULT_OK = 0;                    // OK
    CHECK_RESULT_ALERT = 1;                 // 警告あり
    CHECK_RESULT_ERROR = 2;                 // NG
    CHECK_RESULT_NONE = 9;                  // 未チェック


    { ログ }
    LOG_LINE    = '-----------------------------------';

    { メッセージ }
    MSG_CONFIRM_RETRY = 'リトライしましたが失敗しました。' + ''#13#10'' +
                        'もう一度ランダム実行しますか？';
var
  MainFormf: TMainFormf;

implementation

{$R *.dfm}



// ***********************************************************************
// *    event  : フォームのOnCreateイベント
// ***********************************************************************
procedure TMainFormf.evtFormCreate(Sender: TObject);
begin
    LogWrite(LOG_LINE);
    LogWrite('■FormCreate');
    LogWrite(LOG_LINE);

    Init();                                 // 初期処理
end;

// ***********************************************************************
// *    event  : フォームのOnCloseイベント
// ***********************************************************************
procedure TMainFormf.evtFormClose(Sender: TObject; var Action: TCloseAction);
begin
    LogWrite(LOG_LINE);
    LogWrite('■FormClose');
    LogWrite(LOG_LINE);

    // 終了処理
    EComboYM.Items.Clear();
    DDSet.Close();
    DDSetSub.Close();
    EGrid.Columns.Clear();
    EGridSub.Columns.Clear();
    DDSet.Fields.Clear();
    DDSetSub.Fields.Clear();
end;


// ***********************************************************************
// *    event  : ランダム実行ボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBExecClick(Sender: TObject);
var
    iCnt    : Integer;
    iRetry  : Integer;
begin
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // チェック結果バーもクリア

    iCnt := 0;
    iRetry := 0;

    while iCnt < MAX_COUNT_RETRY do         // リトライ回数をカウント
    begin
        LogWrite(LOG_LINE);
        LogWrite('SetRandom: ' + IntToStr(iCnt + 1 + (MAX_COUNT_RETRY * iRetry)) + '回目');
        LogWrite(LOG_LINE);

        if SetRandom() = 0 then             // 配列にランダム値を設定
        begin
//            MessageDlg('正常に終了しました。', mtInformation, [mbOk], 0);
            break;
        end
        else
        begin
            Inc(iCnt);
        end;

        if iCnt = MAX_COUNT_RETRY then      // リトライ最大回数になった
        begin
                                            // まだリトライするか確認する
            if MessageDlg(MSG_CONFIRM_RETRY, mtConfirmation, mbYesNo, 0, mbYes) = mrYes then
            begin
                iCnt := 0;
                Inc(iRetry);
                continue;
            end;
        end;
    end;

    AryToDS();                              // 配列につめた情報をDataSetにコピーする
end;

// ***********************************************************************
// *    event  : ランダムクリアボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBClearRandomClick(Sender: TObject);
begin
    ClearAryRandom();                       // 手入力したセル以外をクリアする
    AryToDS();
end;

// ***********************************************************************
// *    event  : 全部クリアボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBClearAllClick(Sender: TObject);
begin
    ClearAryAll();
    AryToDS();
end;

// ***********************************************************************
// *    event  : インポートボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBImportClick(Sender: TObject);
var
    bDirect : Boolean;
begin
    bDirect := False;

    { Ctrlを押しながらクリックで手入力インポートにする }
    if GetAsynckeyState(VK_CONTROL) < 0 then
    begin
        bDirect := True;                    // 手入力フラグON
    end;

    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // チェック結果バーもクリア

    // ★
    // 日付範囲を先に選択してもらう
    if ImportCSV(bDirect) then
    begin
        AryToDS();
    end;

end;

// ***********************************************************************
// *    event  : エクスポートボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBExportClick(Sender: TObject);
begin
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // チェック結果バーもクリア

    // CSV出力
    OutputCSV();
end;

// ***********************************************************************
// *    event  : 再集計ボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBSummaryClick(Sender: TObject);
begin
    SetSummary();
    AryToDS();
end;

// ***********************************************************************
// *    event  : 年月コンボのOnChangeイベント
// ***********************************************************************
procedure TMainFormf.evtEComboYMChange(Sender: TObject);
begin
    { グリッドの値をいったんすべてクリア }
    ClearAryAll();
    AryToDS();

    { グリッドの日にちを作り直す }
    SetColumnCaption(EComboYM.ItemIndex);   // カラムのタイトル設定
    SetWeekDays();                          // 曜日再設定
    SetColumnCaption2();                    // 曜日カラムのタイトル設定
    SetColumnColor();                       // 日にちカラムの背景色設定

    { ファイル出力設定 }
    SetOutputFileName();                    // 出力ダイアログのデフォルトファイル名変更
end;

// ***********************************************************************
// *    event  : グリッドのOnDoubleClickイベント
// ***********************************************************************
procedure TMainFormf.evtEGridDblClick(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
    iStatus : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // クリックした列index(0オリジン)から日付を取得(固定列分を引く)

    if iDay < 1 then                        // 日付列でなければスルー
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // メンバID取得

                                            // Wクリックしたセルの設定値を変更する（トグル）
    iStatus := GetNextStatus(m_aryDay[iDay - 1].aryMember[iNo]);

    { 配列にセットして、手入力値としてフラグをたてる }
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


    { データセット更新 }
    EGrid.DataSource.DataSet.Edit;
    EGrid.SelectedField.AsString := GetStatusStr(iStatus);
    EGrid.DataSource.DataSet.Post;

end;

// ***********************************************************************
// *    event  : グリッドのOnDrawColumnCellイベント
// ***********************************************************************
procedure TMainFormf.evtEGridDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
    iDay    : Integer;
    iNo     : Integer;
begin
                                            // 日にちフィールドの場合
    if LeftStr(Column.FieldName, 3) = FLD_DAY then
    begin
                                            // フィールド名から日にちを抽出(1オリジン)
        iDay := StrToInt(StringReplace(Column.FieldName, FLD_DAY, '', [rfReplaceAll]));
        iNo := Column.Field.DataSet.RecNo;  // 行からメンバIDを取得

        { ICUの人のマークを青字にする }
        if iNo = m_aryDay[iDay - 1].iDayICU then    // 配列は0オリジン
        begin
            EGrid.Canvas.Font.Color := clBlue;
        end;

        { 手入力した値はセル色を変える }
        if m_aryDay[iDay - 1].aryFix[iNo] then
        begin
            EGrid.Canvas.Brush.Color := clWebLightYellow;
        end;
    end;

{
    // 参考
    // フィールド名がDayから始まり、表示値が「宿」のセルに色をつける
    if (LeftStr(Column.FieldName, 3) = 'Day') and
        (Column.Field.AsString = '宿') then
}


    //描画
    EGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);

end;

// ***********************************************************************
// *    event   : エクセル出力ボタンのOnClickイベント
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
    { Excel起動 }
    MsExcel := CreateOleObject('Excel.Application');
    MsApplication := MsExcel.Application;
//    MsApplication.Visible := True;
    WBook := MsApplication.WorkBooks.Add;
    WSheet :=WBook.ActiveSheet;

    { 表のタイトル }
    WSheet.Cells[1, 1].Value := DlgSave.FileName;
    WSheet.Cells[1, 1].Font.Size := 15;

    { カラムタイトル出力 }
    WSheet.Cells[3, 1].Value := 'No';
    WSheet.Cells[3, 2].Value := '氏名';
    WSheet.Cells[3, 3].Value := '担当';
                                            // 日にち
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        WSheet.Cells[3, iCnt + 4].Value := IntToStr(m_aryDay[iCnt].iDspDay) + '日';
    end;

    { データ出力 }
    try
        DDSet.DisableControls;

        DDSet.First;
        iRow := 4;
        while not DDSet.Eof do
        begin
                                            // NO
            WSheet.Cells[iRow, 1].Value := DDSet.FieldByName(FLD_NO).AsString;
                                            // 氏名
            WSheet.Cells[iRow, 2].Value := DDSet.FieldByName(FLD_NAME).AsString;
                                            // 担当
            WSheet.Cells[iRow, 3].Value := DDSet.FieldByName(FLD_ROLE).AsString;
                                            // 日にち
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

    //保存の確認を行う
    WBook.Saved := False;

end;

// ***********************************************************************
// *    event   : 設定ボタンのOnClickイベント
// ***********************************************************************
procedure TMainFormf.evtBSettingsClick(Sender: TObject);
var
    dlg : TSettingf;
begin
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // チェック結果バーもクリア

    dlg := TSettingf.Create(Self);          // 設定ダイアログ

    dlg.ShowModal();                        // ダイアログ表示

end;

// ***********************************************************************
// *    event   : ポップアップのOnMenuPopupイベント
// *    memo    :
// ***********************************************************************
procedure TMainFormf.evtPopupMenuPopup(Sender: TObject);
var
    iDay    : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // クリックした列index(0オリジン)から日付を取得(固定列分を引く)

    if iDay < 1 then                        // 日付列でなければスルー
    begin
        N1.Enabled := False;
    end
    else
    begin
        N1.Enabled := True;
    end;

end;

// ***********************************************************************
// *    event   : ポップアップのクリアボタンのOnClickイベント
// *    memo    :
// ***********************************************************************
procedure TMainFormf.evtN1Click(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // クリックした列index(0オリジン)から日付を取得(固定列分を引く)

    if iDay < 1 then                        // 日付列でなければスルー
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // メンバID取得

                                            // Wクリックしたセルの設定値をクリア
    { 配列にセットして、手入力値としてフラグをOFF }
    m_aryDay[iDay - 1].aryMember[iNo] := ST_EMPTY;
    m_aryDay[iDay - 1].aryFix[iNo] := False;

    { データセット更新 }
    EGrid.DataSource.DataSet.Edit;
    EGrid.SelectedField.AsString := '';
    EGrid.DataSource.DataSet.Post;
end;

// ***********************************************************************
// *    event   : ポップアップの固定ボタンのOnClickイベント
// *    memo    : 入力されている値を手入力扱いにする
// ***********************************************************************
procedure TMainFormf.evtN2Click(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // クリックした列index(0オリジン)から日付を取得(固定列分を引く)

    if iDay < 1 then                        // 日付列でなければスルー
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // メンバID取得

    { 未入力セルでなければ固定する }
    if m_aryDay[iDay - 1].aryMember[iNo] <> ST_EMPTY then
    begin
        m_aryDay[iDay - 1].aryFix[iNo] := True; // 手入力とする
    end;

end;

// ***********************************************************************
// *    event   : ポップアップのICUボタンのOnClickイベント
// *    memo    : ICU設定/解除
// ***********************************************************************
procedure TMainFormf.evtICU1Click(Sender: TObject);
var
    iDay    : Integer;
    iNo     : Integer;
begin
    iDay := EGrid.SelectedIndex - 2;        // クリックした列index(0オリジン)から日付を取得(固定列分を引く)

    if iDay < 1 then                        // 日付列でなければスルー
    begin
        Exit;
    end;

    iNo := EGrid.DataSource.DataSet.RecNo;  // メンバID取得

    { 未入力セルでなければICU設定/解除する }
    if m_aryDay[iDay - 1].aryMember[iNo] <> ST_EMPTY then
    begin
        if m_aryDay[iDay - 1].iDayICU <> iNo then
        begin
            m_aryDay[iDay - 1].iDayICU := iNo; // ICU設定
        end
        else
        begin
            m_aryDay[iDay - 1].iDayICU := 0;   // ICU解除
        end;
    end;

end;

// ★チェック項目修正
// ***********************************************************************
// *    event   : チェックボタンのOnClickイベント
// *    memo    : チェック項目
// *                ① ×が期間内に8回あるか
// *                ②宿直と宿直の間が１日しかないときに警告
// *                ③笠川・関口・吉田のうち2人が○じゃないと警告
// *                ④日勤ICUが1日に一人いないと警告
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

    { ①各人の×が期間内に８回あるかチェック }
    for iNo := Low(m_aryMember) to High(m_aryMember) do
    begin
        iCntX := 0;
        for iCnt := Low(m_aryDay) to High(m_aryDay) do
        begin
            // 祝日でない×の数をカウント
            if (not m_aryDay[iCnt].bHoliday) and (m_aryDay[iCnt].aryMember[iNo] = ST_X) then
            begin
                Inc(iCntX);
            end
            // または土日で〇が透析以外の人についている場合も実際は休日なのでカウント
            else if (IsDonitiSyukujitu(iCnt)) and (m_aryDay[iCnt].aryMember[iNo] = ST_NORMAL) and
                (m_aryMember[iNo].strRole <> '透析') then
            begin
                Inc(iCntX);
            end;
        end;
        if iCntX < 8 then
        begin
            ShowMessage('×が８回ない： ' + GetMemberStr(iNo));
            bErr := True;
        end;
    end;

    if bErr then
    begin
        Dec(iResult);
    end;

    bErr := False;

    { ②宿直と宿直の間が１日しかないときに警告 }
    for iNo := Low(m_aryMember) to High(m_aryMember) do
    begin
        iDay := 0;
        for iCnt := Low(m_aryDay) to High(m_aryDay) do
        begin
            // 宿直の場合
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
                        ShowMessage('宿直と宿直の間が近すぎる： ' + GetMemberStr(iNo));
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
    { ③笠川・関口・吉田のうち2人が○じゃないと警告 }
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        if IsDonitiSyukujitu(iCnt) then
        begin
            continue;
        end;

        iCntX := 3;
        // 休み・午後休チェック　
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
            ShowMessage('、関口CE、のうち2人以上がいない: ' + IntToStr(m_aryDay[iCnt].iDspDay) + '日');
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

    { ④日勤ICUが1日に一人いないと警告 }
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        if m_aryDay[iCnt].iDayICU = 0 then
        begin
            ShowMessage('日勤ICUがいない: ' + IntToStr(m_aryDay[iCnt].iDspDay) + '日');
            bErr := True;
            break;
        end;
    end;

    if bErr then
    begin
        Dec(iResult);
    end;

    bErr := False;

    { ⑤その他チェック }
    // ★


    { チェック結果　（問題なし）緑→黄色→赤（クレームくるレベル）}

    if iResult = iCheckCnt then
    begin
        PaintBar(CHECK_RESULT_OK);
        LResult.Caption := '問題なし';
    end
    else if iResult <= 0 then
    begin
        PaintBar(CHECK_RESULT_ERROR);
        LResult.Caption := 'クレーム必至';
    end
    else if iResult < iCheckCnt then
    begin
        PaintBar(CHECK_RESULT_ALERT);
        LResult.Caption := '微妙';
    end;

    Self.Repaint;                           // メッセージの後ろが描画されないことがあるので再描画

    ShowMessage('チェックロジック終了');

end;

// ***********************************************************************
// *    event  : 隠しパネルのOnDoubleClickイベント
// ***********************************************************************
procedure TMainFormf.evtPHiddenDblClick(Sender: TObject);
begin
    { 隠し機能表示 }
    BExport.Visible := True;
end;

// ***********************************************************************
// *    Name    : Init
// *    Args    : Nothing
// *    Create  : K.Endo    2017/06/05
// *    Memo    : 初期処理
// ***********************************************************************
procedure TMainFormf.Init();
var
    bActive : Boolean;
    iCnt    : Integer;
begin

    PHidden.Color := clBtnFace;


    ReadIni();                              // 設定ファイル読み込み

    // --------------------------------
    // 年月コンボの初期処理
    // --------------------------------
    MakeYMComboItem();
    EComboYM.ItemIndex := 0;
    SetOutputFileName();                    // ファイル出力Dlgのデフォルトファイル名設定

    // --------------------------------
    //  メインデータセットの初期処理
    // --------------------------------
    DDSet.Close;
    DDSet.FieldDefs.Clear();
    // 基本フィールド作成
                                            // No
    DDSet.FieldDefs.Add(FLD_NO, ftInteger);
                                            // 氏名
    DDSet.FieldDefs.Add(FLD_NAME, ftString, 16);
                                            // 担当
    DDSet.FieldDefs.Add(FLD_ROLE, ftString, 4);

    // 28日分のフィールドを固定で作成
    for iCnt := 1 to COUNT_DAY do
    begin                                   // Day1～28
        DDSet.FieldDefs.Add(FLD_DAY + IntToStr(iCnt), ftString, 6);
    end;
                                            // index用フィールド
    DDSet.IndexDefs.Add(FLD_NOIDX, FLD_NO, [ixPrimary]);

    DDSet.CreateDataSet();
    DDSet.Close();

    for iCnt := 0 to DDSet.FieldDefs.Count - 1 do
    begin
        DDSet.FieldDefs[iCnt].CreateField(DDSet);
    end;

    // インデックスの設定
    DDSet.IndexName := FLD_NOIDX;

    // --------------------------------
    //  サブデータセットの初期処理
    // --------------------------------
    DDSetSub.Close;
    DDSetSub.FieldDefs.Clear();
                                            // No
    DDSetSub.FieldDefs.Add(FLD_NO, ftInteger);
                                            // 休日出勤
    DDSetSub.FieldDefs.Add(FLD_HOLIDAYWORK, ftInteger);
                                            // 金・休日宿直
    DDSetSub.FieldDefs.Add(FLD_WEEKENDNIGHT, ftInteger);
                                            // index用フィールド
    DDSetSub.IndexDefs.Add(FLD_NOIDX, FLD_NO, [ixPrimary]);

    DDSetSub.CreateDataSet();
    DDSetSub.Close();

    for iCnt := 0 to DDSetSub.FieldDefs.Count - 1 do
    begin
        DDSetSub.FieldDefs[iCnt].CreateField(DDSetSub);
    end;

    // インデックスの設定
    DDSetSub.IndexName := FLD_NOIDX;

    // --------------------------------
    //  メンバーの行作成
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
            // メイングリッド
            DDSet.AppendRecord([m_aryMember[iCnt].iNo,
                            m_aryMember[iCnt].strName,
                            m_aryMember[iCnt].strRole]);
            // サブグリッド
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
    //  メイングリッドの初期処理
    // --------------------------------
    EGrid.Columns[0].Title.Caption := 'No';
    EGrid.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGrid.Columns[0].ReadOnly := True;
    EGrid.Columns[0].Width := 24;

    EGrid.Columns[1].Title.Caption := '氏名';
    EGrid.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGrid.Columns[1].ReadOnly := True;

    EGrid.Columns[2].Title.Caption := '担当';
    EGrid.Columns[2].Title.Alignment := TAlignment.taCenter;
    EGrid.Columns[2].Alignment := TAlignment.taCenter;
    EGrid.Columns[2].ReadOnly := True;

    // 日付カラム
    for iCnt := 1 to COUNT_DAY do
    begin
        EGrid.Columns[iCnt + 2].Title.Alignment := TAlignment.taCenter;
        EGrid.Columns[iCnt + 2].Alignment := TAlignment.taCenter;
        EGrid.Columns[iCnt + 2].Width := 26;
    end;

    SetColumnCaption(EComboYM.ItemIndex);   // カラムのタイトル設定

    SetWeekDays();                          // 曜日再設定

    // --------------------------------
    //  サブグリッドの初期処理
    // --------------------------------
                                            // スクロールバーOFF
                                            // （本当はssNoneで縦横なしになるけどssVerticalで両方なしになる）
    TssDBGrid(EGridSub).ScrollBars := ssVertical;

    EGridSub.Columns[0].Title.Caption := 'No';
    EGridSub.Columns[0].Title.Alignment := TAlignment.taCenter;
    EGridSub.Columns[0].Visible := False;

    EGridSub.Columns[1].Title.Caption := '休日出勤';
    EGridSub.Columns[1].Title.Alignment := TAlignment.taCenter;
    EGridSub.Columns[1].Width := 54;

    EGridSub.Columns[2].Title.Caption := '金/祝日宿直';
    EGridSub.Columns[2].Title.Alignment := TAlignment.taCenter;
    EGridSub.Columns[2].Width := 74;

    // --------------------------------
    //  グリッド（曜日）の初期処理
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

    // 日付カラム
    for iCnt := 1 to COUNT_DAY do
    begin
        EGrid2.Columns[iCnt + 2].Title.Alignment := TAlignment.taCenter;
        EGrid2.Columns[iCnt + 2].Width := 26;
    end;

    SetColumnCaption2();                    // カラムのタイトル設定

    SetColumnColor();                       // 休日の列に色設定(EGrid, EGrid2)

end;

// ***********************************************************************
// *    Name    : ReadIni
// *    Args    : Nothing
// *    Create  : K.Endo    2017/07/04
// *    Memo    : iniファイル読み込み　メンバの配列につめる
// ***********************************************************************
procedure TMainFormf.ReadIni();
var
    iCnt    : Integer;
    ini     : TIniFile;
begin
                                            // exeのあるパスにiniファイルがある
    ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));

    try
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
// *    Name    : MakeYMComboItem
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo  2017/06/06
// *    Memo    : 年月コンボのアイテム作成
// ***********************************************************************
procedure TMainFormf.MakeYMComboItem();
var
    iCnt        : Integer;
    iYear       : Integer;
    iMonth      : Integer;
    iDay        : Integer;
    strTitle    : String;                   // コンボのアイテムに入れる名前
    strFormat   : String;
    dt          : TDateTime;
begin
    strFormat := '%d年%d月%d日';

    EComboYM.Items.Clear();

    dt := StrToDate(DATE_ORIGIN);

    // 初日から12サイクル分のアイテムを動的に作成
    for iCnt := Low(m_aryYM) to High(m_aryYM) do
    begin
        // 初日の年月日取得 ex) 2017年5月21日
        iYear := YearOf(dt);
        iMonth := MonthOf(dt);
        iDay := DayOf(dt);

        strTitle := Format(strFormat, [iYear, iMonth, iDay]);

        { 配列に開始日を保存 }
        m_aryYM[iCnt].dtStartDay := dt;
        m_aryYM[iCnt].iYear := iYear;
        m_aryYM[iCnt].iMonth := iMonth;
        m_aryYM[iCnt].iDay := iDay;

        strTitle := strTitle + ' ～ ';
        dt := IncDay(dt, COUNT_DAY - 1);    // 28日目

        // 最終日の年月日取得 ex) 2017年6月17日
        iYear := YearOf(dt);
        iMonth := MonthOf(dt);
        iDay := DayOf(dt);

        strTitle := strTitle + Format(strFormat, [iYear, iMonth, iDay]);

        m_aryYM[iCnt].strTitle := strTitle; // ex) 2017年5月21日 ～ 2017年6月17日
        EComboYM.Items.Add(strTitle);

        dt := IncDay(dt, 1);                // 次のサイクル
    end;

end;

// ***********************************************************************
// *    Name    : SetColumnCaption
// *    Args    : iIndex    : 日付コンボindex
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/12
// *    Memo    : グリッドのカラムキャプションを設定する
// ***********************************************************************
procedure TMainFormf.SetColumnCaption(iIndex : Integer);
var
    iCnt    : Integer;
    iDay    : Integer;
    dt      : TDateTime;
begin
    // 初日
    dt := m_aryYM[iIndex].dtStartDay;

    // 日付カラム
    for iCnt := 1 to COUNT_DAY do
    begin
        iDay := DayOf(dt);
        EGrid.Columns[iCnt + 2].Title.Caption := IntToStr(iDay);
        dt := IncDay(dt);                   // 次の日へ
    end;
end;

// ***********************************************************************
// *    Name    : SetColumnCaption2
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/12
// *    Memo    : グリッド（曜日）のカラムキャプションを設定する
// ***********************************************************************
procedure TMainFormf.SetColumnCaption2();
var
    iCnt    : Integer;
begin
    // 日付カラムに曜日を設定
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
// *    Memo    : 配列に表示用日にち・曜日・休日設定
// ***********************************************************************
procedure TMainFormf.SetWeekDays();
var
    iCnt    : Integer;
    strBuf  : String;
    dt      : TDateTime;
begin
                                            // 開始年月日取得
    dt := m_aryYM[EComboYM.ItemIndex].dtStartDay;

                                            // 何月でも無条件に28日分設定する
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
                                            // 日にち
        m_aryDay[iCnt].iDspDay := DayOf(dt);
                                            // TDateTimeの日にち
        m_aryDay[iCnt].dtDay := dt;
                                            // 曜日
        m_aryDay[iCnt].iWeekDay := DayOfTheWeek(dt);
                                            // 土日の場合
        if m_aryDay[iCnt].iWeekDay in [DaySaturday, DaySunday] then
        begin
            m_aryDay[iCnt].bHoliday := False;
        end
                                            // 祝日の場合
        else if IsSpecialHoliday(dt, strBuf) then
        begin
            m_aryDay[iCnt].bHoliday := True;
        end
        else                                // それ以外
        begin
            m_aryDay[iCnt].bHoliday := False;
        end;

        dt := IncDay(dt);                   // 翌日へ
    end;
end;

// ***********************************************************************
// *    Name    : SetColumnColor
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/08
// *    Memo    : グリッドの日にちカラムの色設定
// ***********************************************************************
procedure TMainFormf.SetColumnColor();
var
    iCnt    : Integer;
begin
    // 土日・祝日の背景色をピンクにする
    for iCnt := Low(m_aryDay) to High(m_aryDay) do
    begin
        if IsDonitiSyukujitu(iCnt) then     // 土日祝日

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
// *    Return  : Integer   : -1: エラー
// *    Create  : K.Endo    2017/06/06
// *    Memo    : ランダム枠を埋める
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
    ClearAryRandom();                       // 配列の手入力以外をクリア

    { 条件を見ながら構造体の配列にいったんつめる }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        // ----------------------------------------
        // ①前日によって翌日のシフトが決まる
        // ----------------------------------------
        if iDay = 0 then                    // 1日の場合は前月末日をチェックする
        begin
            // ★前月末日と比較
        end
        else
        begin

            for iCnt := Low(m_aryDay[iDay - 1].aryMember) to High(m_aryDay[iDay - 1].aryMember) do
            begin
                // 前日が土以外のとき
                if not (m_aryDay[iDay - 1].iWeekday in [DayFriday, DaySaturday, DaySunday]) then
                begin
                    // 前日に宿直の人は、翌日が非になる
                    iNo := m_aryDay[iDay - 1].iNightICU;
                    SetStatus(iDay, iNo, ST_HIBAN);
                end;
            end;
        end;

        // ----------------------------------------
        // ②土曜・祝日は透析に1人出る
        // ----------------------------------------
        if (m_aryDay[iDay].iWeekDay = DaySaturday) or
            ((m_aryDay[iDay].iWeekDay <> DaySunday) and (m_aryDay[iDay].bHoliday)) then
        begin
                                            // 透析で土曜・祝日に出られる人をランダムで1人抽出
            GetSatTouseki(iDay, iTouseki);
            if iTouseki > 0 then            // 手入力で設定済みの場合、0が返される
            begin
                SetStatus(iDay, iTouseki, ST_NORMAL);
            end;
        end;

        // ----------------------------------------
        // ③未設定の箇所に宿直を埋める
        // ----------------------------------------
        // 宿直
        iNo := GetInputNight(iDay);         // 手入力で宿直が入っていれば取得
        if iNo = 0 then                     // 手入力なし
        begin
            iNo := GetNight(iDay);          // ランダムで選択
            if iNo = -1 then
            begin
{$IFDEF DEBUG}
//            ShowMessage('宿直できるメンバがいません。処理を中断します。');
{$ENDIF}
                Exit;
            end;
        end;
        m_aryDay[iDay].iNightICU := iNo;
        if SetStatus(iDay, iNo, ST_STAY) then
        begin
            Inc(m_aryNightCnt[iNo]);        // 宿直回数＋
        end;

        // ----------------------------------------
        // ④日曜は出られる人が決まっている
        // ----------------------------------------
        if m_aryDay[iDay].iWeekDay = DaySunday then
        begin
            iNo := GetInputSunMember(iDay); // 手入力で日曜に〇がついている人を返す
            if iNo = 0 then                 // 手入力なし
            begin
                iNo := GetSunMember(iDay);  // ランダムで選択
                if iNo = -1 then
                begin
{$IFDEF DEBUG}
//                ShowMessage('日曜出勤できるメンバがいません。処理を中断します。');
{$ENDIF}
                    Exit;
                end;
                SetStatus(iDay, iNo, ST_NORMAL);
            end;
        end;

        // ----------------------------------------
        // ⑤未設定の箇所に日勤ICUを埋める（仮）
        // ----------------------------------------
        // 日勤ICU
        iNo := GetDayICU(iDay);
        if iNo = -1 then
        begin
{$IFDEF DEBUG}
//            ShowMessage('日勤ICUできるメンバがいません。処理を中断します。');
{$ENDIF}
            Exit;
        end;

        m_aryDay[iDay].iDayICU := iNo;

        // ----------------------------------------
        // ⑥残った未設定の箇所に〇、×を埋める
        // ----------------------------------------
        for iCnt := Low(m_aryDay[iDay].aryMember) to High(m_aryDay[iDay].aryMember) do
        begin
            // 未設定の欄を無条件に埋める
            if m_aryDay[iDay].aryMember[iCnt] = ST_EMPTY then
            begin
                // 土日祝日は×
                if IsDonitiSyukujitu(iDay) then
                begin
                    SetStatus(iDay, iCnt, ST_X);
                end
                // それ以外は〇
                else
                begin
                    SetStatus(iDay, iCnt, ST_NORMAL);
                end;
            end;
        end;

    end;

    // ----------------------------------------
    // ⑦金曜宿直の前日は４になる
    // ----------------------------------------
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        if m_aryDay[iDay].iWeekday = DayFriday then
        begin
            iNo := m_aryDay[iDay].iNightICU;

            if iDay > 0 then                // 初日でなければ
            begin
                                            // 前日を４にする
                SetStatus(iDay - 1, iNo, ST_4HOUR);

                // ICUを設定した後に４に上書きするため、
                // この日のICU担当をほかの人にずらす必要がある
                                            // この人がICU担当の場合
                if iNo = m_aryDay[iDay - 1].iDayICU then
                begin
                    iNo2 := GetDayICU(iDay - 1);
                    m_aryDay[iDay - 1].iDayICU := iNo2;
                end;
            end;
        end;
    end;

    // ----------------------------------------
    // ⑧土日に勤務した人の代休を埋める
    // ----------------------------------------

    for iNo := 1 to COUNT_MEMBER do
    begin
        iCnt := 0;
        { 休日出勤の回数取得 }
        for iDay := Low(m_aryDay) to High(m_aryDay) do
        begin
                                            // 土日祝日
            if IsDonitiSyukujitu(iDay) then
            begin
                                            // 宿
                if m_aryDay[iDay].aryMember[iNo] =ST_STAY then
                begin
                    Inc(iCnt);
                                            // 金・休日宿直＋
                    Inc(m_aryMember[iNo].iWeekendNight);
                end
                                            // ○
                else if m_aryDay[iDay].aryMember[iNo] =ST_NORMAL then
                begin
                    // 透析勤務の場合(日勤ICUだけの○は代休不要)
                    // ★透析担当以外の人が業務で出勤したときの考慮がない
                        if m_aryMember[iNo].strRole = '透析' then
                    begin
                        Inc(iCnt);
                    end;
                end;

            end
                                            // 金曜日
            else if m_aryDay[iDay].iWeekday = DayFriday then
            begin
                                            // 宿
                if m_aryDay[iDay].aryMember[iNo] =ST_STAY then
                begin
                                            // 金・休日宿直＋
                    Inc(m_aryMember[iNo].iWeekendNight);
                end;
            end;
        end;

        { 休日出勤の回数分代休セット }
        while iCnt > 0 do
        begin
            iDay := GetDaikyu(iNo);
            if iDay < 0 then
            begin
{$IFDEF DEBUG}
//                ShowMessage('代休が取れない( ﾉД`)');
{$ENDIF}
                Exit;
            end;
            SetStatus(iDay, iNo, ST_X);
            Dec(iCnt);
        end;
    end;


    SetSummary();                           // 休日出勤数を再集計
    Result := 0;

    LogWrite(LOG_LINE);
    LogWrite('SetRandom: 正常終了');
    LogWrite(LOG_LINE);

end;

// ***********************************************************************
// *    Name    : SetSummary
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2021/01/28
// *    Memo    : 集計値をメンバ変数内にセットする（AryToDSでグリッドに表示）
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

        // いったんクリアして集計しなおす
        m_aryMember[iMember].iWeekendNight := 0;
        m_aryMember[iMember].iHolidayWork := 0;

        { 休日出勤・金/休日宿直の回数取得 }
        for iDay := Low(m_aryDay) to High(m_aryDay) do
        begin
                                            // 土日祝日
            if IsDonitiSyukujitu(iDay) then
            begin
                                            // 宿
                if m_aryDay[iDay].aryMember[iMember] =ST_STAY then
                begin
                                            // 金・休日宿直＋
                    Inc(iCntNight);
                end
                                            // ○か早A
                else if m_aryDay[iDay].aryMember[iMember] in [ST_NORMAL, ST_HAYAA] then
                begin
                    Inc(iCntDayTime);
                end;
            end
                                            // 金曜日
            else if m_aryDay[iDay].iWeekday = DayFriday then
            begin
                                            // 宿
                if m_aryDay[iDay].aryMember[iMember] =ST_STAY then
                begin
                                            // 金・休日宿直＋
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
// *    Args    : iDay      : 日にち配列index
// *              iNo       : メンバID
// *              iStatus   : 設定値
// *    Return  : Boolean   : T: 設定した/ F: 設定していない(手入力済みだった)
// *    Create  : K.Endo    2017/06/13
// *    Memo    : 配列に設定値をうめる(手入力されていなければうめない)
// ***********************************************************************
function TMainFormf.SetStatus(iDay: Integer; iNo: Integer; iStatus: Integer): Boolean;
begin
    Result := False;

    if not m_aryDay[iDay].aryFix[iNo] then  // 手入力されていなければ
    begin
                                            // 指定された設定値をうめる
        m_aryDay[iDay].aryMember[iNo] := iStatus;
        Result := True;
    end;
end;

// ***********************************************************************
// *    Name    : AryToDS
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/08
// *    Memo    : 配列からデータセットにコピーする
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

        // m_aryDay構造体配列からデータセットにコピーする
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
                                            // スクロールバーOFF
                                            // データセットのカーソルが動くとスクロールバーがまた表示される
                                            // さらに、ScrollBarsプロパティを一度変更しないとスクロールバーが消えない
        TssDBGrid(EGridSub).ScrollBars := ssNone;
                                            // （本当はssNoneで縦横なしになるけどssVerticalで両方なしになる）
        TssDBGrid(EGridSub).ScrollBars := ssVertical;
    end;
end;

// ***********************************************************************
// *    Name    : ClearAll
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/08
// *    Memo    : 配列をすべてクリアする
// ***********************************************************************
procedure TMainFormf.ClearAryAll();
var
    iCnt    : Integer;
    iDay    : Integer;
begin
    { 日付配列クリア }
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

    { 宿直回数配列クリア }
    for iCnt := Low(m_aryNightCnt) to High(m_aryNightCnt) do
    begin
        m_aryNightCnt[iCnt] := 0;
        m_aryMember[iCnt].iHolidayWork := 0;
        m_aryMember[iCnt].iWeekendNight := 0;
    end;

    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // チェック結果バーもクリア
end;

// ***********************************************************************
// *    Name    : ClearAryRandom
// *    Args    : Nothing
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/13
// *    Memo    : 配列から手入力したセル以外をクリアする
// ***********************************************************************
procedure TMainFormf.ClearAryRandom();
var
    iCnt    : Integer;
    iDay    : Integer;
begin
    { 日付配列クリア }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        for iCnt := 1 to COUNT_MEMBER do
        begin
            // 手入力したところ以外
            SetStatus(iDay, iCnt, ST_EMPTY);
        end;
        m_aryDay[iDay].iDayICU := 0;
        m_aryDay[iDay].iNightICU := 0;      // 判断が難しいのでいったんクリア
    end;

    { 宿直回数配列クリア }
    for iCnt := Low(m_aryNightCnt) to High(m_aryNightCnt) do
    begin
        m_aryNightCnt[iCnt] := 0;
    end;

    { 宿が手入力されていた場合、宿直担当を後から補完 }
    for iDay := Low(m_aryDay) to High(m_aryDay) do
    begin
        for iCnt := 1 to COUNT_MEMBER do
        begin
            if m_aryDay[iDay].aryMember[iCnt] = ST_STAY then
            begin
                m_aryDay[iDay].iNightICU := iCnt;
                Inc(m_aryNightCnt[iCnt]);   // 宿直回数＋
            end;
        end;
    end;

    SetSummary();                           // 休日出勤数を再集計
    LResult.Caption := '';
    PaintBar(CHECK_RESULT_NONE);            // チェック結果バーもクリア
end;


// ***********************************************************************
// *    Name    : GetNextStatus
// *    Args    : iStatus   : ステータスID
// *    Return  : Integer   : 次のステータスID(トグル)
// *    Create  : K.Endo    2017/06/13
// *    Memo    : ステータスIDから次のステータスIDを返す
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
// *    Args    : String    : 表示用文字列(〇、非、×など)
// *    Return  : iStatus   : ステータスID
// *    Create  : K.Endo    2017/10/27
// *    Memo    : 表示用文字列からステータスIDを返す
// ***********************************************************************
function TMainFormf.GetStatusID(strStatus: String): Integer;
var
    iCnt    : Integer;
begin
    Result := ST_EMPTY;

    if strStatus = '○' then                // 漢字を記号に置き換える(インポート対応)
    begin
        strStatus := '〇';
    end
    else if strStatus = '休暇' then
    begin
        strStatus := '休';
    end
    else if strStatus = '4' then
    begin
        strStatus := '４';
    end;
    // ★早A、早B、Pも変換が必要？

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
// *    Args    : iStatus   : ステータスID
// *    Return  : String    : 表示用文字列(〇、非、×など)
// *    Create  : K.Endo    2017/06/08
// *    Memo    : ステータスIDから表示用文字列を返す
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
// *    Args    : iMemberNo : メンバーNo
// *    Return  : String    : 氏名
// *    Create  : K.Endo    2017/06/09
// *    Memo    : メンバーNoから氏名を返す
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
// *    Args    : iWeekday  : 曜日定義(DayMonday,～）
// *    Return  : String    : 曜日
// *    Create  : K.Endo    2017/06/09
// *    Memo    : 曜日定義値から曜日の文字列を返す
// ***********************************************************************
function TMainFormf.GetWeekdayStr(iWeekday: Integer): String;
begin
    Result := '';

    case iWeekday of
        DayMonday:      Result := '月';
        DayTuesday:     Result := '火';
        DayWednesday:   Result := '水';
        DayThursday:    Result := '木';
        DayFriday:      Result := '金';
        DaySaturday:    Result := '土';
        DaySunday:      Result := '日';
    end;

end;

// ***********************************************************************
// *    Name    : GetSunMember
// *    Args    : iDay      :
// *    Return  : Integer   : メンバNo
// *    Create  : K.Endo    2017/06/09
// *    Memo    : 日曜に出るメンバNoを返却する
// ***********************************************************************
function TMainFormf.GetSunMember(iDay: Integer): Integer;
var
    iRandom : Integer;
    sTrg    : set of 0..COUNT_MEMBER;       // 対象者リスト
begin
    iRandom := 0;

                                            // 日曜に出る人
    sTrg := MEM_ALL - [MEM_I];

    while not (iRandom in sTrg) do
    begin
        if sTrg = [] then
        begin
            LogWrite('GetSunMember: 日曜出勤できる人がいない');
            LogWrite(IntToStr(m_aryDay[iDay].iDspDay) + '日');
            Result := -1;
            Exit;
        end;
                                            // ランダム取得
        iRandom := Random(COUNT_MEMBER + 1);

        if not (iRandom in sTrg) then       // ランダム値が対象者リストにない
        begin
            continue;
        end;

                                            // すでに×、休、宿の人は除く
        if m_aryDay[iDay].aryMember[iRandom] in [ST_X, ST_HOLIDAY, ST_STAY] then
        begin
            sTrg := sTrg - [iRandom];       // 対象者リストから外す
            continue;
        end;

        // 金土日祝日の宿直回数が多い人は除外する
        if m_aryMember[iRandom].iWeekendNight = MAX_COUNT_NIGHT then
        begin
            sTrg := sTrg - [iRandom];       // 対象者リストから外す
            continue;
        end;

    end;

    Result := iRandom;

end;

// ***********************************************************************
// *    Name    : GetInputSunMember
// *    Args    : iDay      : 日付配列index
// *    Return  : Integer   : メンバNo(0: 手入力なし)
// *    Create  : K.Endo    2017/06/09
// *    Memo    : 手入力で日曜に〇をつけられているメンバNoを返却する
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
// *    Args    : iDay  : 日にち配列index
// *              iNo   : 透析メンバ格納域
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/09
// *    Memo    : 土曜・祝日の透析メンバをランダムで1人抽出する
// ***********************************************************************
procedure TMainFormf.GetSatTouseki(iDay: Integer; var iNo: Integer);
var
    iCnt    : Integer;
    iRandom : Integer;
    sTrg    : set of 0..COUNT_MEMBER;       // 対象者リスト
    procedure IgnoreMember(iMember: Integer; var iFixCnt: Integer);
    begin
        { 手入力で土曜・祝日に〇がついているメンバは除外する }
        if m_aryDay[iDay].aryMember[iMember] = ST_NORMAL then
        begin
            sTrg := sTrg - [iMember];
            Inc(iFixCnt);
        end;

        { 金土日祝日の宿直回数が多いメンバを除外する }
        if m_aryMember[iMember].iWeekendNight = MAX_COUNT_NIGHT then
        begin
            sTrg := sTrg - [iMember];       // 対象者リストから外す
            Inc(iFixCnt);
        end;
    end;
begin
    iCnt := 0;
    iRandom := 0;
    iNo := 0;

                                            // 土曜に出る透析メンバ
    sTrg := [MEM_J, MEM_G, MEM_K, MEM_L];

    { 手入力済みまたは休日出勤が多い人を除外する }
    IgnoreMember(MEM_J, iCnt);
    IgnoreMember(MEM_G, iCnt);
    IgnoreMember(MEM_K, iCnt);
    IgnoreMember(MEM_L, iCnt);

    if iCnt = 1 then                        // 手入力で1人既に確定している
    begin
        Exit;
    end;

    // ランダムでメンバ設定
    while not (iRandom in sTrg) do
    begin
        iRandom := Random(COUNT_MEMBER + 1);
    end;
    iNo := iRandom;

end;

// ***********************************************************************
// *    Name    : GetNight
// *    Args    : iDay  : 日にち
// *    Return  : 宿直するメンバNo
// *    Create  : K.Endo    2017/06/06
// *    Memo    : 宿直するメンバNoを返却する
// ***********************************************************************
function TMainFormf.GetNight(iDay: Integer): Integer;
var
    iRandom : Integer;
    iCnt    : Integer;
    sIgnore : set of 0..COUNT_MEMBER;       // 除外リスト
begin
    iRandom := 0;
    sIgnore := [0];

    while iRandom in sIgnore do
    begin
        if sIgnore = [0..COUNT_MEMBER] then
        begin
            Result := -1;
            LogWrite('GetNight: 宿直できる人がいない。 ' + IntToStr(m_aryDay[iDay].iDspDay) + '日');
            Exit;
        end;
                                            // メンバのNoをランダムで選択
        iRandom := Random(COUNT_MEMBER + 1);

        if iRandom in sIgnore then          // 除外リストのNoの場合はやり直し
        begin
            continue;
        end;

        // ★今は透析で出た人も宿直しているぽい
        // 〇、非番、×の人の場合(透析の人は土曜に〇がつくことがあるが宿直しない）
{        if m_aryDay[iDay].aryMember[iRandom] in  [ST_NORMAL, ST_HIBAN, ST_4HOUR, ST_X, ST_HOLIDAY] then
        begin
            sIgnore := sIgnore + [iRandom];
            continue;
        end;
}
        // 宿直しない人だったらやり直し
        if not m_aryMember[iRandom].bNight then
        begin
            sIgnore := sIgnore + [iRandom];
            continue;
        end;

        // ★可能な範囲か？試用してみてから見直す
        // 前日～4日前までが宿直の場合はやり直し（次の宿直までに4日あけたい）
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

        // 3回宿直している人は対象外になる
        if m_aryNightCnt[iRandom] = MAX_COUNT_NIGHT then
        begin
            sIgnore := sIgnore + [iRandom];
            continue;
        end;

        // ★今は宿直に関する個人の条件がないぽい
        { 月曜日にの宿直はNGのためやり直し }
{        if iRandom = MEM_F then
        begin
            if m_aryDay[iDay].iWeekday = DayMonday then
            begin
                sIgnore := sIgnore + [iRandom];
                continue;
            end;
        end;
}
        {TODO: その他宿直メンバを決定する条件}

    end;


    Result := iRandom;
end;

// ***********************************************************************
// *    Name    : GetInputNight
// *    Args    : iDay  : 日にち
// *    Return  : 宿直するメンバNo (0: 宿直なし)
// *    Create  : K.Endo    2017/06/15
// *    Memo    : 手入力で宿直に設定されたメンバNoを返す
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
// *    Args    : iDay  : 日にち
// *    Return  : 日勤ICUするメンバNo
// *    Create  : K.Endo    2017/06/08
// *    Memo    : 日勤ICUするメンバNoを返却する
// ***********************************************************************
function TMainFormf.GetDayICU(iDay: Integer): Integer;
var
    iRandom : Integer;
    sTrg    : set of 0..COUNT_MEMBER;       // 対象者リスト
begin
    iRandom := 0;

    // 日勤ICUする人
    sTrg := MEM_ALL - [MEM_I];


    { 除外する条件}

    { 曜日の条件 }
    // 土日祝日以外は
    if not IsDonitiSyukujitu(iDay) then
    begin
        // ★2人以下→1.5人以下 みんな透析固定じゃないけどどうやって算出する？
        // ★透析・光学・カテ・PM・心外を半日単位で設定する？
        // 透析が2人以下のときはICUにまわさない
{        if (m_aryDay[iDay].aryMember[MEM_A] in [ST_4HOUR, ST_HIBAN, ST_X, ST_HOLIDAY]) or
            (m_aryDay[iDay].aryMember[MEM_J] in [ST_4HOUR, ST_HIBAN, ST_X, ST_HOLIDAY]) or
            (m_aryDay[iDay].aryMember[MEM_G] in [ST_4HOUR, ST_HIBAN, ST_X, ST_HOLIDAY]) then
        begin
            LogWrite('GetDayICu: ' + '透析メンバを除外 ' + IntToStr(m_aryDay[iDay].iDspDay) + '日');
                                                // 透析メンバを全員除外
            sTrg := sTrg - [MEM_A, MEM_J, MEM_G];
        end;
}
    end;

    case m_aryDay[iDay].iWeekDay of
        DayTuesday:                         // 【火曜日】心外の人を除外する
        begin
            sTrg := sTrg - [MEM_B, MEM_D, MEM_L];
        end;
        DayWednesday:                       // 【水曜日】
        begin
                                            // 2月、8月以外はを除外する
            if not (MonthOf(m_aryDay[iDay].dtDay) in [2, 8]) then
            begin
                sTrg := sTrg - [MEM_G];
            end;
        end;
    end;


    { 残った人から選択 }
    while not (iRandom in sTrg) do
    begin
        if sTrg = [] then
        begin
            LogWrite('GetDayICU: 日勤ICUできる人がいない ' + IntToStr(m_aryDay[iDay].iDspDay) + '日');
            Result := -1;
            Exit;
        end;

        { 日勤ICUメンバを決定する条件 }

                                            // メンバのNoをランダムで選択
        iRandom := Random(COUNT_MEMBER + 1);

        if not (iRandom in sTrg) then       // 対象者リストのNo以外はやり直し
        begin
            continue;
        end;

        { 土日祝日で未設定の人、宿直の人は除外する(この時点で〇がついている人が対象になる） }
        if IsDonitiSyukujitu(iDay) then
        begin
            if m_aryDay[iDay].aryMember[iRandom] in [ST_EMPTY, ST_STAY] then
            begin
                sTrg := sTrg - [iRandom];
                continue;
            end;
        end;

        { 非番、×、休の人の場合(ランダムの場合はまだ設定されていない。手入力されたとき用) }
        if m_aryDay[iDay].aryMember[iRandom] in [ST_HIBAN, ST_4HOUR, ST_X, ST_HOLIDAY] then
        begin
            sTrg := sTrg - [iRandom];
            continue;
        end;
    end;

    // ex) 1日(月) xxxx 宿
    LogWrite(IntToStr(m_aryDay[iDay].iDspDay) + '(' + GetWeekdayStr(m_aryDay[iDay].iWeekDay) + ') ' +
                GetMemberStr(iRandom) + ' ' +  GetStatusStr(m_aryDay[iDay].aryMember[iRandom]));

    Result := iRandom;
end;

// ***********************************************************************
// *    Name    : GetDaikyu
// *    Args    : iNo       : メンバNo
// *    Return  : Integer   : 代休をとる日のindex
// *    Create  : K.Endo    2017/06/08
// *    Memo    : 代休をとれる日のindexを返す
// ***********************************************************************
function TMainFormf.GetDaikyu(iNo: Integer): Integer;
var
    iRandom     : Integer;
//★    iHibanAke   : Integer;
    iCathe      : Integer;
    sTrg        : set of 0..COUNT_DAY - 1;  // 表示する日付分のリスト

    // ***********************************************************************
    // *    Name    : GetHibanAke
    // *    Args    : Nothing
    // *    Return  : Integer   : 非番の翌日の日付index
    // *    Create  : K.Endo    2017/06/15
    // *    Memo    : 対象者リストから非番の翌日の日付indexを返す
    // ***********************************************************************
    function FindHibanAke(): Integer;
    var
        iCnt    : Integer;
    begin
        Result := 0;
        for iCnt := Low(m_aryDay) to High(m_aryDay) do
        begin
                                            // 非番
            if (m_aryDay[iCnt].aryMember[iNo] = ST_HIBAN) and
                                            // 翌日が対象日付範囲内
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
        if sTrg = [] then                   // 代休をとれる日がない
        begin
            Exit;
        end;

        { できるだけ非番の後に×を入れる }
{★今回は入れない仕様
        iHibanAke := FindHibanAke();        // 対象リストの中から非番の翌日を返す
        if iHibanAke > 0 then               // 非番の翌日が対象リストから見つかった場合
        begin
            iRandom := iHibanAke;           // 非番明け
        end
        else
        begin
}
                                            // 日にち配列のindexをランダムで返す(0オリジン)
            iRandom := Random(High(m_aryDay) + 1);
//★        end;

        if not (iRandom in sTrg) then       // 既に除外されている日
        begin
            continue;
        end;

        { 曜日依存の条件 }
        if IsDonitiSyukujitu(iRandom) then  // 【土日祝日】だったらやり直し
        begin
            sTrg := sTrg - [iRandom];
            continue;
        end;

        Case m_aryDay[iRandom].iWeekday of
            DayMonday:                      // 【月曜日】、が休めない
            begin
                if iNo in [MEM_B, MEM_L] then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;

            DayTuesday:                     // 【火曜日】全員休みNG
            begin
                sTrg := sTrg - [iRandom];
                continue;
            end;

            DayWednesday:                   // 【水曜日】2月、8月以外が休めない
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
                                            // 【木曜日】カテ3人
        if m_aryDay[iRandom].iWeekday = DayThursday then
        begin
            iCathe := 3;
        end
        else                                // 【木曜日以外】カテ2人
        begin
            iCathe := 2;
        end;

        { メンバ依存の条件 }
        case iNo of
            MEM_A:                   // 
            begin
                                            // とが休み、午後休のときは休めない
                if (IsTakeOff(iRandom, MEM_B) and
                    IsTakeOff(iRandom, MEM_E)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_B:                   // 
            begin
                                            // が休み、午後休のときは休めない
                if IsTakeOff(iRandom, MEM_C) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // とが休み、午後休のときは休めない
                if (IsTakeOff(iRandom, MEM_A) and
                    IsTakeOff(iRandom, MEM_E)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_C:                       // 
            begin
                                            // が休み、午後休のときは休めない
                if IsTakeOff(iRandom, MEM_B) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_K:                     // 
            begin
                                            // が休み、午後休のときは休めない
                if IsTakeOff(iRandom, MEM_L) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;

                                            // とが休み、午後休のときは休めない
                if (IsTakeOff(iRandom, MEM_J) and
                    IsTakeOff(iRandom, MEM_I)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_E:                     // 
            begin
                                            // が休み、午後休のときは休めない
                if IsTakeOff(iRandom, MEM_G) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // とが休み、午後休のときは休めない
                if (IsTakeOff(iRandom, MEM_A) and
                    IsTakeOff(iRandom, MEM_B)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // カテ室に3or2人いないときは休めない
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_J:                   // 
            begin
                                            // とが休み、午後休のときは休めない
                if (IsTakeOff(iRandom, MEM_K) and
                    IsTakeOff(iRandom, MEM_I)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_G:                   // 
            begin
                                            // が休み、午後休のときは休めない
                if IsTakeOff(iRandom, MEM_E) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
                                            // カテ室に3or2人いないときは休めない
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_F:                   // 
            begin
                                            // カテ室に3or2人いないときは休めない
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_L:                   // 
            begin
                                            // が休み、午後休のときは休めない
                if IsTakeOff(iRandom, MEM_K) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_I:                    // 
            begin
                                            // とが休み、午後休のときは休めない
                if (IsTakeOff(iRandom, MEM_J) and
                    IsTakeOff(iRandom, MEM_K)) then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
            MEM_H:                    // 
            begin
                                            // カテ室に3or2人いないときは休めない
                if CountCathe(iRandom) < iCathe then
                begin
                    sTrg := sTrg - [iRandom];
                    continue;
                end;
            end;
        end;

                                            // 〇の日でICU当番でなければ代休にできる
        if (m_aryDay[iRandom].aryMember[iNo] = ST_NORMAL) and
            (iNo <> m_aryDay[iRandom].iDayICU) then
        begin
            Result := iRandom;
            Exit;
        end
        else
        begin
            sTrg := sTrg - [iRandom];       // NGな日はリストから除外
        end;
    end;
end;

// ***********************************************************************
// *    Name    : IsTakeOff
// *    Args    : iDay  : 日にち配列index
// *              iNo   : メンバID
// *    Return  : Boolean   : T: iNoのメンバが×、休、４、非/ F: それ以外
// *    Create  : K.Endo    2017/06/12
// *    Memo    : 指定された日に指定されたメンバが休みor午後休かどうかを返す
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
// *    Args    : iDay  : 日にち配列index
// *    Return  : カテ室の休み・午後休でない人の人数

// *    Create  : K.Endo    2021/01/25
// *    Memo    : カテ室の休み・午後休でない人数を返す
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
// *    Args    : iDay  : 日にち配列index
// *    Return  : Boolean   : T: 土日祝日 / F: 平日
// *    Create  : K.Endo    2017/06/16
// *    Memo    : 指定された日が土日祝日か
// ***********************************************************************
function TMainFormf.IsDonitiSyukujitu(iDay: Integer): Boolean;
begin
                                            // 土日
    if (m_aryDay[iDay].iWeekday in [DaySaturday, DaySunday]) or
                                            // または祝日
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
// *    Args    : iLevel  : チェックレベル
// *    Return  : Nothing
// *    Create  : K.Endo    2017/10/31
// *    Memo    : チェック結果バーの表示切替
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
// *    Memo    : ファイル出力ダイアログのデフォルトファイル名変更
// ***********************************************************************
procedure TMainFormf.SetOutputFileName();
begin
    DlgSave.FileName := '勤務割_' + EComboYM.Text;
end;

// ***********************************************************************
// *    Name    : ImportCSV
// *    Args    : bDirect   : T: 手入力/ F: 通常データ
// *    Return  : Nothing
// *    Create  : K.Endo    2017/10/27
// *    Memo    : CSVインポート
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

    // ファイルの指定
    DlgOpen.Filter := 'CSVファイル (*.csv)|*.csv';
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
        stl.LoadFromFile(CSVFile);          // StringListにファイル読み込み

        // 配列につめる
        for iMember := 1 to stl.Count - 1 do// 日付の行は飛ばすので１からでOK
        begin
            // 1行とりだす
            str := stl[iMember];
            // カンマで分ける
            strSplitted.DelimitedText := stl[iMember];

            for iDay := Low(m_aryDay) to High(m_aryDay) do
            begin
                // メンバNoのカラムに「ICU」が入っている行に日勤ICUの人のメンバNoが入っている
                if strSplitted[0] = 'ICU' then
                begin
                    // 日勤ICU
                    m_aryDay[iDay].iDayICU := StrToInt(strSplitted[iDay + 3]);
                end
                else
                begin
                    // ステータス文字列からIDにして配列につめる
                    m_aryDay[iDay].aryMember[iMember] := GetStatusID(strSplitted[iDay + 3]);

                    // 手入力インポートの場合は値が入っている箇所の手入力フラグをONにする
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
// *    Memo    : CSV出力
// ***********************************************************************
procedure TMainFormf.OutputCSV();
var
    F       : TextFile;
    CSVFile : String;
    stl     : TStringList;
    i       : Integer;
    iDay    : Integer;
begin
    // 保存場所の指定
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

        DDSet.DisableControls;              // ちらつき防止

        // ファイル出力
        AssignFile(F, CSVFile);             // ファイルと実ファイル結びつける
        ReWrite(F);                         // ファイルを新規作成して開く

        DDSet.First;
        // タイトル(フィールド)行の出力
        for i := 0 to DDSet.FieldCount - 1 do
        begin
            stl.Add(DDSet.Fields[i].FieldName);
        end;

        Writeln(F, stl.CommaText);
        stl.Clear;

        // リスト出力
        while not DDSet.Eof do
        begin
            for i := 0 to DDSet.FieldCount - 1 do
            begin
                stl.Add(DDSet.Fields[i].AsString);
            end;

            Writeln(F, stl.CommaText);      // テキストファイルに1行出力
            stl.Clear;

            DDSet.Next;
        end;

        // 日勤ICUの人のNoも出力する
        stl.Add('ICU');
        stl.Add('');                        // Name
        stl.Add('');                        // Role
        for iDay := Low(m_aryDay) to High(m_aryDay) do
        begin
            stl.Add(IntToStr(m_aryDay[iDay].iDayICU));
        end;

        Writeln(F, stl.CommaText);          // テキストファイルに1行出力
        stl.Clear;

        CloseFile(F);                       //ファイルを閉じる

    finally
        DDSet.EnableControls;               // ちらつき防止
        stl.Free;
    end;
end;

// ***********************************************************************
// *    Name    : LogWrite
// *    Args    : strLog    : 出力ログ
// *    Return  : Nothing
// *    Create  : K.Endo    2017/06/09
// *    Memo    : ログ出力
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
    path := 'C:\end\03 開発\AutoSchedule.log';
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
// *    Return  : Boolean   : T: 祝日/ F: 祝日でない
// *    Create  : K.Endo    2017/06/06
// *    Memo    : 祝日かどうか判断する
// ***********************************************************************
function TMainFormf.IsSpecialHoliday(ADate: TDate; var AName: string): Boolean;
// ------------------------------------------------------------
// http://koyomi.vis.ne.jp/
// http://www.asahi-net.or.jp/~CI5M-NMR/misc/equinox.html#Rule
// ------------------------------------------------------------
// ADateが祝日かどうかを返す。
// 祝日=True,祝日ではない=False
// AName には祝日の名前を返す
var
    DName: string;
    i:Integer;
    {FreqOfWeek Begin}
    function FreqOfWeek(AYear, AMonth: Word; AWeekNo, ADayOfWeeek: Byte): TDateTime;
    // AYear年AMonth月の第AWeekNo「ADayOfWeeek曜日」の日付を返す
    // ADayOfWeeek　日曜日=1..土曜日=7
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
    // SYearからEYear迄に何回閏年があるかを返す
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
    // Ayearの春分の日を求める
    var
        dDay: Word;
    begin
        dDay := Trunc((21.147 + ((AYear - 1940) * 0.2421904) - (LeapYearCount(1940, AYear) - 1)));
        result := EncodeDate(AYear, 3, dDay);
    end;
    {VernalEquinox End}
    {AutumnalEquinox End}
    function AutumnalEquinox(AYear: Word): TDateTime;
    // Ayearの秋分の日を求める
    var
        dDay: Word;
    begin
        dDay := Trunc((23.5412 + ((AYear - 1940) * 0.2421904) - (LeapYearCount(1940, AYear) - 1)));
        result := EncodeDate(AYear, 9, dDay);
    end;
    {AutumnalEquinox End}
    {_IsSpecialHoliday Begin}
    function _IsSpecialHoliday(ADate: TDate; var AName: string): Boolean;
    // ADateが祝日かどうかを返す。
    // 祝日=True,祝日ではない=False
    // AName には祝日の名前を返す
    // '国民の休日'はここでは算出されない
    var
        dYear, dMonth, dDay: Word;
    begin
        AName := '';
        result := False;
        DecodeDate(ADate, dYear, dMonth, dDay);
        case dMonth of
          1:
            begin
              // '元日' 1948～
              if (dYear >= 1948) and (dDay = 1) then
                begin
                  result := True;
                  AName := '元日';
                  Exit;
                end;
              // '成人の日①' 1948～1999
              if (dYear >= 1948) and (dYear <= 1999) and (dDay = 15) then
                begin
                  result := True;
                  AName := '成人の日';
                  Exit;
                end;
              // '成人の日②' 2000～
              // 第２月曜日(ハッピーマンデー)
              if (dYear >= 2000) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 2, 2) then
                    begin
                      result := True;
                      AName := '成人の日';
                      Exit;
                    end;
                end;
            end;
          2:
            begin
              // '建国記念の日' 1966～
              if (dYear >= 1966) and (dDay = 11) then
                begin
                  result := True;
                  AName := '建国記念の日';
                  Exit;
                end;
              // ※昭和天皇の大喪の礼(1989/02/24)
              if (dYear = 1989) and (dDay = 24) then
                begin
                  result := True;
                  AName := '昭和天皇の大喪の礼';
                  Exit;
                end;
            end;
          3:
            begin
              // '春分の日' 1949～
              if (dYear >= 1949) then
                begin
                  if ADate = VernalEquinox(dYear) then
                    begin
                      result := True;
                      AName := '春分の日';
                      Exit;
                    end;
                end;
            end;
          4:
            begin
              // '天皇誕生日' 1948～1988
              if (dYear >= 1948) and (dYear <= 1988) and (dDay = 29) then
                begin
                  result := True;
                  AName := '天皇誕生日';
                  Exit;
                end;
              // 'みどりの日①' 1989～2006
              if (dYear >= 1989) and (dYear <= 2006) and (dDay = 29) then
                begin
                  result := True;
                  AName := 'みどりの日';
                  Exit;
                end;
              // '昭和の日' 2007～
              if (dYear >= 2007) and (dDay = 29) then
                begin
                  result := True;
                  AName := '昭和の日';
                  Exit;
                end;
              // ※皇太子明仁親王の結婚の儀(1959/04/10)
              if (dYear = 1959) and (dDay = 10) then
                begin
                  result := True;
                  AName := '皇太子明仁親王の結婚の儀';
                  Exit;
                end;
            end;
          5:
            begin
              // '憲法記念日' 1948～
              if (dYear >= 1948) and (dDay = 3) then
                begin
                  result := True;
                  AName := '憲法記念日';
                  Exit;
                end;
              // 'みどりの日②' 2007～
              if (dYear >= 2007) and (dDay = 4) then
                begin
                  result := True;
                  AName := 'みどりの日';
                  Exit;
                end;
              // 'こどもの日' 1948～
              if (dYear >= 1948) and (dDay = 5) then
                begin
                  result := True;
                  AName := 'こどもの日';
                  Exit;
                end;
            end;
          6:
            begin
              // ※皇太子徳仁親王の結婚の儀(1993/06/09)
              if (dYear = 1993) and (dDay = 9) then
                begin
                  result := True;
                  AName := '皇太子徳仁親王の結婚の儀';
                  Exit;
                end;
            end;
          7:
            begin
              // '海の日①' 1995～2002
              if (dYear >= 1995) and (dYear <= 2002) and (dDay = 20) then
                begin
                  result := True;
                  AName := '海の日';
                  Exit;
                end;
              // '海の日②' 2003～
              // 第３月曜日
              if (dYear >= 2003) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 3, 2) then
                    begin
                      result := True;
                      AName := '海の日';
                      Exit;
                    end;
                end;
            end;
          8:
            begin
              // '山の日' 2016～
              if (dYear >= 2016) and (dDay = 11) then
                begin
                  result := True;
                  AName := '山の日';
                  Exit;
                end;
            end;
          9:
            begin
              // '敬老の日①' 1966～2002
              if (dYear >= 1966) and (dYear <= 2002) and (dDay = 15) then
                begin
                  result := True;
                  AName := '敬老の日';
                  Exit;
                end;
              // '敬老の日②' 2003～
              // 第３月曜日
              if (dYear >= 2003) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 3, 2) then
                    begin
                      result := True;
                      AName := '敬老の日';
                      Exit;
                    end;
                end;
              // '秋分の日' 1948～
              if (dYear >= 1948) then
                begin
                  if ADate = AutumnalEquinox(dYear) then
                    begin
                      result := True;
                      AName := '秋分の日';
                      Exit;
                    end;
                end;
            end;
          10:
            begin
              // '体育の日①' 1966～1999
              if (dYear >= 1966) and (dYear <= 1999) and (dDay = 10) then
                begin
                  result := True;
                  AName := '体育の日';
                  Exit;
                end;
              // '体育の日②' 2000～
              // 第２月曜日(ハッピーマンデー)
              if (dYear >= 2000) then
                begin
                  if ADate = FreqOfWeek(dYear, dMonth, 2, 2) then
                    begin
                      result := True;
                      AName := '体育の日';
                      Exit;
                    end;
                end;
            end;
          11:
            begin
              // '文化の日' 1948～
              if (dYear >= 1948) and (dDay = 3) then
                begin
                  result := True;
                  AName := '文化の日';
                  Exit;
                end;
              // '勤労感謝の日' 1948～
              if (dYear >= 1948) and (dDay = 23) then
                begin
                  result := True;
                  AName := '勤労感謝の日';
                  Exit;
                end;
              // ※即位礼正殿の儀(1990/11/12)
              if (dYear = 1990) and (dDay = 12) then
                begin
                  result := True;
                  AName := '即位礼正殿の儀';
                  Exit;
                end;
            end;
          12:
            begin
              // '天皇誕生日' 1948～
              if (dYear >= 1989) and (dDay = 23) then
                begin
                  result := True;
                  AName := '天皇誕生日';
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
        // 振替休日①　1973/04/12以降
        // 日曜日と祝祭日が重なった場合には'振替休日'となる
        result := True;
        AName := '振替休日';
    end
    else if (ADate >= EncodeDate(1988, 5, 4)) and (DayOfWeek(ADate) <> 1) and
          _IsSpecialHoliday(ADate - 1, DName) and _IsSpecialHoliday(ADate + 1, DName) then
    begin
        // 国民の休日 1988/05/04以降
        // 祝日と祝日に挟まれた平日は'国民の休日'となる。
        result := True;
        AName := '国民の休日';
    end
    else if (ADate >= EncodeDate(2008, 5, 6)) and (DayOfWeek(ADate) <> 1) and
          _IsSpecialHoliday(ADate - DayOfWeek(ADate) + 1, DName) then
    begin
        // 振替休日②　2008/05/06以降
        // '祝日'が日曜日に当たるときは、その日後においてその日に最も近い'祝日'でない日を休日とする
        result := True;
        AName := '振替休日';
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
