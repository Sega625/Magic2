unit uMagic2;

interface

uses
  Windows, Messages, SysUtils, Variants, System.Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, IniFiles,
  ADODB, ComObj, Statistica_le, ComCtrls, Menus, FileCtrl, Structs;

type
  TMDBForm = class(TForm)
    LoadMDBBtn: TButton;
    WafersLB: TListBox;
    ProcGammaBtn: TButton;
    LoadNormsBtn: TButton;
    LoadMapBtn: TButton;
    LoadMDBLab: TLabel;
    LoadNormsLab: TLabel;
    LoadMapLab: TLabel;
    ResultRE: TRichEdit;
    MainMenu1: TMainMenu;
    PrefMenu: TMenuItem;
    ExitMenu: TMenuItem;
    Label1: TLabel;
    Label2: TLabel;
    ClearBtn: TButton;
    MSystemCB: TComboBox;
    Label4: TLabel;
    OpenDirBtn: TButton;
    OpenDirLab: TLabel;
    ProcSchusterBtn: TButton;

    procedure LoadMDBBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ProcGammaBtnClick(Sender: TObject);
    procedure LoadNormsBtnClick(Sender: TObject);
    procedure LoadMapBtnClick(Sender: TObject);
    procedure PrefMenuClick(Sender: TObject);
    procedure ExitMenuClick(Sender: TObject);
    procedure ClearBtnClick(Sender: TObject);
    procedure OpenDirBtnClick(Sender: TObject);
    procedure MSystemCBChange(Sender: TObject);
    procedure ProcSchusterBtnClick(Sender: TObject);
    procedure WafersLBDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
  private
    Lot: TLot;

    Params: TTestsParams;
    ExePath, MeasPath, MeasfName: TFileName;
    ToFirstFail, CreateSTS, NoNorms, MapByParams: byte;
    BegTime: Int64;

    function  LoadMDB      (const fName: TFileName): Boolean;
    function  LoadDirectory(const dName: TFileName): Boolean;
    function  LoadNorms    (const fName: TFileName): Boolean;
    function  LoadMap      (const fName: TFileName): Boolean;
    function  LoadTXTMap(): Boolean;
    function  SaveSTS(const fName: TFileName; Waf: TWafer): Boolean;
    procedure GetSchusterFileInfo(const fName: TFileName; var Module: string); overload;
    procedure GetSchusterFileInfo(const fName: TFileName; var Module, Config: string); overload;

    procedure CloneBlankWafer(sWaf, dWaf: TWafer);

    procedure OnEvent(const EvenType: TEventType; const Str: string);
    procedure Print_Result(const Str: string; const Col: TColor=clBlack);
    function  MDBLoaded()  : Boolean;
    function  NormsLoaded(): Boolean;
    function  MapLoaded()  : Boolean;
    function  GetEnableProcess(): Boolean;
    procedure StartTime();
    function  StopTime(): Double;
  public
    { Public declarations }
  end;

var
  MDBForm: TMDBForm;

implementation

uses PrefDlg;


{$R *.dfm}

/////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.FormCreate(Sender: TObject);                                //
                                                                               //
/////////////////////////////////////////////////////////////////////////////  //
  function GetVersion(var Major, Minor, Release, Build: byte): Boolean;    //  //
  var                                                                      //  //
    info: Pointer;                                                         //  //
    infosize: DWORD;                                                       //  //
    fileinfo: PVSFixedFileInfo;                                            //  //
    fileinfosize: DWORD;                                                   //  //
    tmp: DWORD;                                                            //  //
  begin                                                                    //  //
    Major   := 0;                                                          //  //
    Minor   := 0;                                                          //  //
    Release := 0;                                                          //  //
    Build   := 0;                                                          //  //
                                                                           //  //
    infosize := GetFileVersionInfoSize(PChar(Application.ExeName), tmp);   //  //
    Result := infosize <> 0;                                               //  //
                                                                           //  //
    if Result then                                                         //  //
    begin                                                                  //  //
      info := AllocMem(infosize);                                          //  //
      try                                                                  //  //
        GetFileVersionInfo(PChar(Application.ExeName), 0, infosize, info); //  //
        VerQueryValueA(info, nil, Pointer(fileinfo), fileinfosize);        //  //
        Major   := fileinfo.dwFileVersionMS shr 16;                        //  //
        Minor   := fileinfo.dwFileVersionMS and $FFFF;                     //  //
        Release := fileinfo.dwFileVersionLS shr 16;                        //  //
        Build   := fileinfo.dwFileVersionLS and $FFFF;                     //  //
      finally                                                              //  //
        FreeMem(info, fileinfosize);                                       //  //
      end;                                                                 //  //
    end;                                                                   //  //
  end;                                                                     //  //
/////////////////////////////////////////////////////////////////////////////  //
                                                                               //
var                                                                            //
  IniFile: TIniFile;                                                           //
  Major, Minor, Release, Build: byte;                                          //
begin                                                                          //
  if GetVersion(Major, Minor, Release, Build) then                             //
    Caption := 'Magic2   v.'+IntToStr(Major)+'.'+                              //
                             IntToStr(Minor)+'.'+                              //
                             IntToStr(Release)+'.'+                            //
                             IntToStr(Build)                                   //
  else                                                                         //
    Caption := 'Magic2';                                                       //
  FormatSettings.DecimalSeparator := '.';                                      //
  ExePath := ExtractFilePath(Application.ExeName);                             //
                                                                               //
  PrefForm := TPrefForm.Create(self);                                          //
                                                                               //
  try                                                                          //
    IniFile := TIniFile.Create(ExePath+'Magic2.ini');                          //
    MeasPath  := IniFile.ReadString ('System', 'Path', 'C:');                  //
    self.Top  := IniFile.ReadInteger('System', 'Top',  250);                   //
    self.Left := IniFile.ReadInteger('System', 'Left', 500);                   //
                                                                               //
    ToFirstFail := IniFile.ReadInteger('Preference', 'ToFirstFail', 1);        //
    CreateSTS   := IniFile.ReadInteger('Preference', 'CreateSTS',   1);        //
    NoNorms     := IniFile.ReadInteger('Preference', 'NoNorms',     0);        //
    MapByParams := IniFile.ReadInteger('Preference', 'MapByParams', 0);        //
    MSystemCB.ItemIndex := IniFile.ReadInteger('Preference', 'MeasSystem', 0); //
  finally                                                                      //
    IniFile.Free();                                                            //
  end;                                                                         //
                                                                               //
  LoadMDBLab.Font.Color   := clRed;                                            //
  OpenDirLab.Font.Color   := clRed;                                            //
  LoadNormsLab.Font.Color := clRed;                                            //
  LoadMapLab.Font.Color   := clRed;                                            //
                                                                               //
  OpenDirBtn.Top := LoadMDBBtn.Top;                                            //
  OpenDirLab.Top := LoadMDBLab.Top;                                            //
                                                                               //
  Lot := TLot.Create(Handle);                                                  //
  Lot.OnEvent := OnEvent;                                                      //
                                                                               //
  MSystemCBChange(Self);                                                       //
end;                                                                           //
/////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.FormDestroy(Sender: TObject);                               //
var                                                                            //
  IniFile: TIniFile;                                                           //
begin                                                                          //
  try                                                                          //
    IniFile := TIniFile.Create(ExePath+'Magic2.ini');                          //
    try                                                                        //
      IniFile.WriteString ('System', 'Path', MeasPath);                        //
      IniFile.WriteInteger('System', 'Top',  self.Top);                        //
      IniFile.WriteInteger('System', 'Left', self.Left);                       //
                                                                               //
      IniFile.WriteInteger('Preference', 'ToFirstFail', ToFirstFail);          //
      IniFile.WriteInteger('Preference', 'CreateSTS',   CreateSTS  );          //
      IniFile.WriteInteger('Preference', 'NoNorms',     NoNorms    );          //
      IniFile.WriteInteger('Preference', 'MapByParams', MapByParams);          //
      IniFile.WriteInteger('Preference', 'MeasSystem',  MSystemCB.ItemIndex);  //
    except                                                                     //
    end;                                                                       //
  finally                                                                      //
    IniFile.Free();                                                            //
  end;                                                                         //
                                                                               //
  PrefForm.Free();                                                             //
  Lot.Free();                                                                  //
end;                                                                           //
/////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////
procedure TMDBForm.PrefMenuClick(Sender: TObject);   //
begin                                                //
  with PrefForm do                                   //
  begin                                              //
    ToFirstFailChB.Checked := ToFirstFail = 1;       //
    CreateSTSChB.Checked   := CreateSTS   = 1;       //
    NoNormsChB.Checked     := NoNorms     = 1;       //
    MapByParamsChB.Checked := MapByParams = 1;       //
                                                     //
    Top  := self.Top +50;                            //
    Left := self.Left+0;                             //
    ShowModal();                                     //
                                                     //
    if ToFirstFailChB.Checked then ToFirstFail := 1  //
                              else ToFirstFail := 0; //
    if CreateSTSChB.Checked   then CreateSTS   := 1  //
                              else CreateSTS   := 0; //
    if NoNormsChB.Checked     then NoNorms     := 1  //
                              else NoNorms     := 0; //
    if MapByParamsChB.Checked then MapByParams := 1  //
                              else MapByParams := 0; //
  end;                                               //
                                                     //
  ProcGammaBtn.Enabled := GetEnableProcess();        //
end;                                                 //
///////////////////////////////////////////////////////
///////////////////////////////////////////////////////
procedure TMDBForm.ExitMenuClick(Sender: TObject);   //
begin                                                //
  Close();                                           //
end;                                                 //
///////////////////////////////////////////////////////

///////////////////////////////////////////////////////
procedure TMDBForm.MSystemCBChange(Sender: TObject); //
begin                                                //
  case MSystemCB.ItemIndex of                        //
    0: begin // �����-156                            //
         LoadMDBBtn.Visible := True;                 //
         LoadMDBLab.Visible := True;                 //
                                                     //
         OpenDirBtn.Visible := False;                //
         OpenDirLab.Visible := False;                //
                                                     //
         LoadNormsBtn.Visible := True;               //
         LoadNormsLab.Visible := True;               //
                                                     //
         LoadMapBtn.Visible := True;                 //
         LoadMapLab.Visible := True;                 //
                                                     //
         ProcSchusterBtn.Left := 400;                //
         ProcSchusterBtn.Visible := False;           //
       end;                                          //
    1: begin // Schuster TSM 664                     //
         LoadMDBBtn.Visible := False;                //
         LoadMDBLab.Visible := False;                //
                                                     //
         OpenDirBtn.Visible := True;                 //
         OpenDirLab.Visible := True;                 //
                                                     //
         LoadNormsBtn.Visible := False;              //
         LoadNormsLab.Visible := False;              //
                                                     //
         LoadMapBtn.Visible := False;                //
         LoadMapLab.Visible := False;                //
                                                     //
         ProcSchusterBtn.Left := WafersLB.Left;      //
         ProcSchusterBtn.Visible := True;            //
       end;                                          //
    2: begin // GAMMA TSSemi 2000-400                //
                                                     //
       end;                                          //
  end;                                               //
                                                     //
  Lot.Init();                                        //
  WafersLB.Clear();                                  //
  ResultRE.Clear();                                  //
end;                                                 //
///////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.LoadMDBBtnClick(Sender: TObject);                                          //
var                                                                                           //
  OpenDlg: TOpenDialog;                                                                       //
begin                                                                                         //
  OpenDlg := TOpenDialog.Create(self);                                                        //
  with OpenDlg do                                                                             //
  begin                                                                                       //
    InitialDir := MeasPath;                                                                   //
    Filter := '����� ��������� �����-156 (*.mdb)|*.mdb*';                                     //
    Title := '��������� ���� ����������';                                                     //
                                                                                              //
    if Execute then                                                                           //
      if LoadMDB(FileName) then                                                               //
      begin                                                                                   //
        MeasfName := FileName;                                                                //
        MeasPath := ExtractFilePath(MeasfName);                                               //
                                                                                              //
        LoadMDBLab.Font.Color := $00CAF90D;                                                   //
                                                                                              //
        Print_Result('* ���� ��������� '+ExtractFileName(MeasfName)+' ��������!', clTeal);    //
                                                                                              //
        ProcGammaBtn.Enabled := GetEnableProcess();                                           //
      end                                                                                     //
      else                                                                                    //
      begin                                                                                   //
        LoadMDBLab.Font.Color := clRed;                                                       //
                                                                                              //
        Print_Result('* ������ �������� ����� ��������� '+ExtractFileName(MeasfName), clRed); //
                                                                                              //
        ProcGammaBtn.Enabled := False;                                                        //
      end;                                                                                    //
                                                                                              //
    Free;                                                                                     //
  end;                                                                                        //
end;                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.LoadNormsBtnClick(Sender: TObject);                                        //
var                                                                                           //
  OpenDlg: TOpenDialog;                                                                       //
begin                                                                                         //
  OpenDlg := TOpenDialog.Create(self);                                                        //
  with OpenDlg do                                                                             //
  begin                                                                                       //
    InitialDir := MeasPath;                                                                   //
    Filter := '����� ���� (*.nrm)|*.nrm*';                                                    //
    Title := '��������� ���� ����';                                                           //
                                                                                              //
    if Execute then                                                                           //
      if LoadNorms(FileName) then                                                             //
      begin                                                                                   //
        LoadNormsLab.Font.Color := $00CAF90D;                                                 //
                                                                                              //
        Print_Result('* ����� �� ����� '+ExtractFileName(FileName)+' ���������!', clTeal);    //
                                                                                              //
        ProcGammaBtn.Enabled := GetEnableProcess();                                           //
      end                                                                                     //
      else                                                                                    //
      begin                                                                                   //
        LoadNormsLab.Font.Color := clRed;                                                     //
                                                                                              //
        Print_Result('* ������ �������� ���� '+ExtractFileName(FileName), clRed);             //
                                                                                              //
        ProcGammaBtn.Enabled := GetEnableProcess();                                           //
      end;                                                                                    //
                                                                                              //
    Free;                                                                                     //
  end;                                                                                        //
end;                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.LoadMapBtnClick(Sender: TObject);                                          //
var                                                                                           //
  OpenDlg: TOpenDialog;                                                                       //
begin                                                                                         //
  OpenDlg := TOpenDialog.Create(self);                                                        //
  with OpenDlg do                                                                             //
  begin                                                                                       //
    InitialDir := MeasPath;                                                                   //
    Filter := '����� ������ (*.sts)|*.sts*';                                                  //
    Title := '��������� ����� ������';                                                        //
                                                                                              //
    if Execute then                                                                           //
      if LoadMap(FileName) then                                                               //
      begin                                                                                   //
        LoadMapLab.Font.Color := $00CAF90D;                                                   //
                                                                                              //
        Print_Result('* ����� ������ '+ExtractFileName(FileName)+' ���������!', clTeal);      //
                                                                                              //
        ProcGammaBtn.Enabled := GetEnableProcess();                                           //
      end                                                                                     //
      else                                                                                    //
      begin                                                                                   //
        LoadMapLab.Font.Color := clRed;                                                       //
                                                                                              //
        Print_Result('* ������ �������� ����� ������ '+ExtractFileName(FileName), clRed);     //
                                                                                              //
        ProcGammaBtn.Enabled := GetEnableProcess();                                           //
      end;                                                                                    //
                                                                                              //
    Free;                                                                                     //
  end;                                                                                        //
end;                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.OpenDirBtnClick(Sender: TObject);                                          //
var                                                                                           //
  sDir: string;                                                                               //
begin                                                                                         //
  sDir := MeasPath;                                                                           //
                                                                                              //
  if SelectDirectory(sDir, [], 0) then                                                        //
    if LoadDirectory(sDir) then                                                               //
    begin                                                                                     //
      MeasPath := sDir;                                                                       //
      OpenDirLab.Font.Color := $00CAF90D;                                                     //
                                                                                              //
      Print_Result('* ����� ��������� ���������!', clTeal);                                   //
                                                                                              //
      ProcSchusterBtn.Enabled := True;                                                        //
    end                                                                                       //
    else                                                                                      //
    begin                                                                                     //
      OpenDirLab.Font.Color := clRed;                                                         //
                                                                                              //
      Print_Result('* ������ �������� ������ ���������!', clRed);                             //
                                                                                              //
      ProcSchusterBtn.Enabled := False;                                                       //
    end;                                                                                      //
end;                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.ProcGammaBtnClick(Sender: TObject);                                //
var                                                                                   //
  n, nWf: DWORD;                                                                      //
  Str: string;                                                                        //
  P: WORD;                                                                            //
  STSFName: TFileName;                                                                //
begin                                                                                 //
  if WafersLB.SelCount = 0 then                                                       //
  begin                                                                               //
    Print_Result('�� ������� ��������!', clRed);                                      //
    Exit;                                                                             //
  end;                                                                                //
                                                                                      //
  StartTime();                                                                        //
                                                                                      //
  Lot.LoadHistGroupsFromIni();                                                        //
                                                                                      //
  nWf := 0;                                                                           //
                                                                                      //
  Lot.Init();                                                                         //
  SetLength(Lot.Wafer, WafersLB.SelCount);                                            //
                                                                                      //
  Print_Result('... ��� �������� �������!');                                         //
                                                                                      //
  with WafersLB do                                                                    //
    for n := 0 to Items.Count-1 do                                                    //
      if Selected[n] then                                                             //
      begin                                                                           //
        Str := Items[n];                                                              //
        P := Pos('�', Str);                                                           //
        if P <> 0 then Str := Copy(Items[n], P+2, Length(Items[n]));                  //
                                                                                      //
        if Lot.Wafer[nWf] = nil then                                                  //
          Lot.Wafer[nWf] := TWafer.Create(Handle);                                    //
                                                                                      //
        if not Lot.Wafer[nWf].LoadGammaMDB(MeasfName, Str) then                       //
        begin                                                                         //
          Print_Result('������ �������� ��������!', clRed);                           //
          Continue;                                                                   //
        end;                                                                          //
        Lot.LfName := MeasfName;                                                      //
                                                                                      //
                                                                                      //
        if NormsLoaded then // ���� ����� ����                                        //
          if not Lot.Wafer[nWf].AddNorms(Params) then                                 //
          begin                                                                       //
            Print_Result('������ ���������� ����!', clRed);                           //
            Exit;                                                                     //
          end;                                                                        //
                                                                                      //
//        if MapLoaded then // ���� ����� ������ ����                                   //
//          if Lot.BlankWafer.NTotal <> Lot.Wafer[nWf].NTotal then                      //
//          begin                                                                       //
//            Print_Result('����� ������ �� �������� � ��������!', clRed);              //
//            Exit;                                                                     //
//          end;                                                                        //
                                                                                      //
        if CreateSTS = 1 then                                                         //
        begin                                                                         //
          STSFName := ChangeFileExt(MeasfName, '')+'_'+Str+'.sts';                    //
                                                                                      //
          if SaveSTS(STSFName, Lot.Wafer[nWf]) then                                   //
            Print_Result('������ ����: '+ExtractFileName(STSFName), clBlue)           //
          else                                                                        //
            Print_Result('������ �������� �����: '+ExtractFileName(STSFName), clRed); //
        end;                                                                          //
                                                                                      //
        Inc(nWf);                                                                     //
      end;                                                                            //
                                                                                      //
  SetLength(Lot.Wafer, nWf);                                                          //
                                                                                      //
  Print_Result('>>> �������� ���������!', clGreen);                                   //
                                                                                      //
  if NormsLoaded then // ���� ����� ����                                              //
    Lot.SaveXLS(ToFirstFail = 1, MapByParams = 1)                                     //
  else                                                                                //
    Lot.SaveXLS(False, MapByParams = 1);                                              //
                                                                                      //
  Print_Result( '����� ��������� '+FormatFloat('0.0', StopTime())+' ���.');           //
end;                                                                                  //
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.ProcSchusterBtnClick(Sender: TObject);                             //
var                                                                                   //
  n, nWf: DWORD;                                                                      //
  Str, sModule, sConfig: string;                                                      //
  P: WORD;                                                                            //
  STSFName: TFileName;                                                                //
begin                                                                                 //
  if WafersLB.SelCount = 0 then                                                       //
  begin                                                                               //
    Print_Result('�� ������� ����� ���������!', clRed);                               //
    Exit;                                                                             //
  end;                                                                                //
                                                                                      //
  StartTime();                                                                        //
                                                                                      //
  Lot.LoadHistGroupsFromIni();                                                        //
                                                                                      //
  nWf := 0;                                                                           //
                                                                                      //
  Lot.Init();                                                                         //
  SetLength(Lot.Wafer, WafersLB.SelCount);                                            //
                                                                                      //
  Print_Result('... ��� �������� ������ ���������!');                                //
                                                                                      //
  with WafersLB do                                                                    //
    for n := 0 to Items.Count-1 do                                                    //
      if Selected[n] then                                                             //
      begin                                                                           //
        MeasfName := MeasPath+'\'+Items[n];                                           //
                                                                                      //
        sModule := '';                                                                //
        sConfig := '';                                                                //
        GetSchusterFileInfo(MeasfName, sModule, sConfig);                             //
        if sModule <> '' then                                                         //
          if nWf = 0 then                                                             //
          begin                                                                       //
            Lot.LDevice := sModule;                                                   //
            Lot.LConfig := sConfig;                                                   //
          end                                                                         //
          else                                                                        //
            if (sModule <> Lot.LDevice) or                                            //
               (sConfig <> Lot.LConfig)                                               //
            then                                                                      //
            begin                                                                     //
              Print_Result('������ �������� ����� '+Items[n]+'!', clRed);             //
              Continue;                                                               //
            end;                                                                      //
                                                                                      //
        if Lot.Wafer[nWf] = nil then                                                  //
          Lot.Wafer[nWf] := TWafer.Create(Handle);                                    //
                                                                                      //
        if not Lot.Wafer[nWf].LoadSchusterTXT(MeasfName) then                         //
        begin                                                                         //
          Print_Result('������ �������� ����� ���������!', clRed);                    //
          Continue;                                                                   //
        end;                                                                          //
        Lot.LfName := MeasfName;                                                      //
                                                                                      //
                                                                                      //
        if Lot.BlankWafer.NTotal = 0 then                                             //
          if Lot.Wafer[nWf].Diameter <> 0 then // �������� ����� ������ �� ��������   //
          begin
            if not Lot.CreateBlankWafer(nWf) then Print_Result('������ �������� ������� ������!', clRed);
          end;
                                                                                      //
        Inc(nWf);                                                                     //
      end;                                                                            //
                                                                                      //
  SetLength(Lot.Wafer, nWf);                                                          //
                                                                                      //
  Print_Result('>>> ����� ���������!', clGreen);                                      //
                                                                                      //
  if Lot.BlankWafer.NTotal = 0 then
    if QuestMess(Handle, '����� ������ �� ���� ���������!'+#13#10+'������ ��������� ����� ������?') = IDYES then
      if not LoadTXTMap() then
        Lot.BlankWafer.Init();                                                        //
                                                                                      //
  if CreateSTS = 1 then                                                               //
    for nWF := 0 to Length(Lot.Wafer)-1 do                                            //
    begin                                                                             //
      STSFName := ChangeFileExt(MeasfName, '')+'.sts';                                //
                                                                                      //
      if SaveSTS(STSFName, Lot.Wafer[nWf]) then                                       //
        Print_Result('������ ����: '+ExtractFileName(STSFName), clBlue)               //
      else                                                                            //
        Print_Result('������ �������� �����: '+ExtractFileName(STSFName), clRed);     //
    end;                                                                              //
                                                                                      //
  Lot.SaveXLS(ToFirstFail = 1, MapByParams = 1);                                      //
                                                                                      //
  Print_Result(' ����� ��������� '+FormatFloat('0.0', StopTime())+' ���.')            //
end;                                                                                  //
////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////
procedure TMDBForm.ClearBtnClick(Sender: TObject); //
begin                                              //
  ResultRE.Clear();                                //
end;                                               //
/////////////////////////////////////////////////////


procedure TMDBForm.WafersLBDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
begin
// ����� �������� ����� Style = lbOwnerDrawFixed;

//  with WafersLB.Canvas do
//  begin
//    Brush.Color := RGB(100, 150, 150);
//    FillRect(Rect);
//    Font.Color := clRed;
//    TextOut(Rect.Left, Rect.Top, WafersLB.Items[Index]);
//  end;
end;


//////////////////////////////////////////////////////////////
function TMDBForm.LoadMDB(const fName: TFileName): Boolean; //
var                                                         //
  ADOConnection: TADOConnection;                            //
  n: DWORD;                                                 //
begin                                                       //
  Result := False;                                          //
                                                            //
  ADOConnection := TADOConnection.Create(nil);              //
  with ADOConnection do                                     //
  try                                                       //
    Connected := False;                                     //
    ConnectionString := ConnStr+fName+';';                  //
    LoginPrompt := False;                                   //
    GetTableNames(WafersLB.Items);                          //
  except                                                    //
    Print_Result(fName+': ������ ADO!', clRed);             //
    Exit;                                                   //
  end;                                                      //
                                                            //
  ADOConnection.Free;                                       //
                                                            //
  if WafersLB.Items.Count > 0 then                          //
  begin                                                     //
    for n := 0 to WafersLB.Items.Count-1 do                 //
    WafersLB.Items[n] := '�������� � '+WafersLB.Items[n];   //
  end                                                       //
  else                                                      //
  begin                                                     //
    Print_Result(fName+' : � ����� ��� �������!', clRed);   //
    Exit;                                                   //
  end;                                                      //
                                                            //
  Result := True;                                           //
end;                                                        //
//////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////
function TMDBForm.LoadDirectory(const dName: TFileName): Boolean;                    //
var                                                                                  //
  sr: TSearchRec;                                                                    //
  Str1: string;                                                                      //
begin                                                                                //
  Result := True;                                                                    //
                                                                                     //
  if FindFirst(dName+'\*.txt', faAnyFile, sr) = 0  then  //���� ����� TXT � �������� //
  begin                                                                              //
    WafersLB.Clear;                                                                  //
                                                                                     //
    repeat                                                                           //
      Str1 := '';                                                                    //
      GetSchusterFileInfo(sr.Name, Str1);                                            //
                                                                                     //
      if Str1 <> '' then             // ���� ���� Schuster                           //
        WafersLB.Items.Add(sr.Name); // ������� ������ � ListBox                     //
                                                                                     //
    until FindNext(sr) <> 0;                                                         //
  end                                                                                //
  else                                                                               //
    Result := False;                                                                 //
                                                                                     //
  FindClose(sr);                                                                     //
                                                                                     //
  if WafersLB.Items.Count = 0 then Result := False;                                  //
end;                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////
function TMDBForm.LoadNorms(const fName: TFileName): Boolean;         //
var                                                                   //
  n, Cnt: DWORD;                                                      //
  P, P1, P2: WORD;                                                    //
  SL: TStringList;                                                    //
  Str: string;                                                        //
begin                                                                 //
  Result := False;                                                    //
                                                                      //
  SetLength(Params, 0);                                               //
  SL := TStringList.Create();                                         //
  SL.LoadFromFile(fName);                                             //
  for n := 0 to SL.Count-1 do                                         //
  begin                                                               //
    Str := Trim(SL.Strings[n]);                                       //
    if Str = '' then Continue;                                        //
                                                                      //
////////// ����� ������ � �����������                                //
                                                                      //
    P := Pos('"', Str); // ������ cp, cn, on                          //
    if P <> 0 then                                                    //
    begin                                                             //
      Cnt := Length(Params);                                          //
      SetLength(Params, Cnt+1);                                       //
                                                                      //
      Delete(Str, 1, Pos(',', Str));                                  //
                                                                      //
////////// ��� ���������                                              //
                                                                      //
      P := Pos('"', Str); // 1-� �������                              //
      Delete(Str, 1, P);                                              //
      P := Pos('"', Str); // 2-� �������                              //
      Params[Cnt].Name := Trim(Copy(Str, 1, P-1));                    //
                                                                      //
////////// ����� ���������                                            //
                                                                      //
      P1 := Pos('(', Str);                                            //
      P2 := Pos(')', Str);                                            //
      Params[Cnt].PMode := Copy(Str, P1+1, P2-P1-1);                  //
                                                                      //
      Delete(Str, 1, P); // ������ �� 2-� �������                     //
      P := Pos(',', Str);                                             //
      Delete(Str, 1, P);                                              //
                                                                      //
////////// ����. ���������                                            //
                                                                      //
      P := Pos(',', Str);                                             //
      try                                                             //
        Params[Cnt].Norma.Max := StrToFloat(Trim(Copy(Str, 1, P-1))); //
      except                                                          //
        Params[Cnt].Norma.Max := NotSpec;                             //
      end;                                                            //
                                                                      //
      Delete(Str, 1, P);                                              //
                                                                      //
////////// ���. ���������                                             //
                                                                      //
      P := Pos(',', Str);                                             //
      try                                                             //
        Params[Cnt].Norma.Min := StrToFloat(Trim(Copy(Str, 1, P-1))); //
      except                                                          //
        Params[Cnt].Norma.Min := -NotSpec;                            //
      end;                                                            //
                                                                      //
      Delete(Str, 1, P);                                              //
                                                                      //
////////// ������� ���������                                          //
                                                                      //
      P := Pos('"', Str); // 1-� �������                              //
      Delete(Str, 1, P);                                              //
      P := Pos('"', Str); // 2-� �������                              //
      Params[Cnt].PUnit := Trim(Copy(Str, 1, P-1));                   //
                                                                      //
      Result := True;                                                 //
    end;                                                              //
  end;                                                                //
                                                                      //
  SL.Free();                                                          //
end;                                                                  //
////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////
function TMDBForm.LoadMap(const fName: TFileName): Boolean;  //
begin                                                        //
  Result := False;                                           //
                                                             //
  if Lot.BlankWafer.LoadBlankSTS(fName) then Result := True; //
end;                                                         //
///////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////////////////////////
function TMDBForm.LoadTXTMap(): Boolean;                                                               //
var                                                                                                    //
  OpenDlg: TOpenDialog;                                                                                //
begin                                                                                                  //
  OpenDlg := TOpenDialog.Create(self);                                                                 //
  with OpenDlg do                                                                                      //
  begin                                                                                                //
    InitialDir := MeasPath;                                                                            //
    Filter := '����� ������ (*.map)|*.map*';                                                           //
    Title := '��������� ����� ������';                                                                 //
                                                                                                       //
    if Execute then                                                                                    //
    begin                                                                                              //
      Result :=  Lot.BlankWafer.LoadBlankTXT(FileName);                                                //
      if Result then Print_Result('* ����� ������ '+ExtractFileName(FileName)+' ���������!', clTeal)   //
                else Print_Result('* ������ �������� ����� ������ '+ExtractFileName(FileName), clRed); //
    end;                                                                                               //
                                                                                                       //
    Free;                                                                                              //
  end;                                                                                                 //
end;                                                                                                   //
/////////////////////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////
function TMDBForm.SaveSTS(const fName: TFileName; Waf: TWafer): Boolean;           //
var                                                                                //
  tmpWafer: TWafer;                                                                //
  X, Y, i, vX, vY: DWORD;                                                          //
begin                                                                              //
  Result := False;                                                                 //
                                                                                   //
  if Waf.NTotal = 0 then Exit;                                                     //
                                                                                   //
  if not MapLoaded() then                                                          //
  begin                                                                            //
    Waf.SaveSTS(fName);                                                            //
                                                                                   //
    Result := True;                                                                //
                                                                                   //
    Exit;                                                                          //
  end;                                                                             //
                                                                                   //
  tmpWafer := TWafer.Create(Handle);                                               //
  CloneBlankWafer(Lot.BlankWafer, tmpWafer);                                       //
                                                                                   //
  tmpWafer.fName := Waf.fName;                                                     //
  tmpWafer.Code  := Waf.Code;                                                      //
  tmpWafer.TimeDate := Waf.TimeDate;                                               //
  tmpWafer.NLot := Waf.NLot;                                                       //
  tmpWafer.Num  := Waf.Num;                                                        //
                                                                                   //
  SetLength(tmpWafer.TestsParams, Length(Waf.TestsParams));                        //
  for i := 0 to Length(tmpWafer.TestsParams)-1 do                                  //
  begin                                                                            //
    tmpWafer.TestsParams[i].Name      := Waf.TestsParams[i].Name;                  //
    tmpWafer.TestsParams[i].Norma.Min := Waf.TestsParams[i].Norma.Min;             //
    tmpWafer.TestsParams[i].Norma.Max := Waf.TestsParams[i].Norma.Max;             //
  end;                                                                             //
                                                                                   //
  for Y := 0 to Length(tmpWafer.Chip)-1 do                                         //
    for X := 0 to Length(tmpWafer.Chip[0])-1 do                                    //
      if tmpWafer.Chip[Y, X].Status = 0 then                                       //
      begin                                                                        //
        if (tmpWafer.Chip[Y, X].ID-1) < Waf.NTotal then                            //
        begin                                                                      //
          vX := Waf.ChipN[tmpWafer.Chip[Y, X].ID-1].X;                             //
          vY := Waf.ChipN[tmpWafer.Chip[Y, X].ID-1].Y;                             //
                                                                                   //
          tmpWafer.Chip[Y, X].Status := Waf.Chip[vY, vX].Status;                   //
                                                                                   //
          SetLength(tmpWafer.Chip[Y, X].ChipParams, Length(tmpWafer.TestsParams)); //
          for i := 0 to Length(tmpWafer.TestsParams)-1 do                          //
            tmpWafer.Chip[Y, X].ChipParams[i] := Waf.Chip[vY, vX].ChipParams[i];   //
        end;                                                                       //
      end;                                                                         //
                                                                                   //
  tmpWafer.SaveSTS(fName);                                                         //
                                                                                   //
  tmpWafer.Free();                                                                 //
                                                                                   //
  Result := True;                                                                  //
end;                                                                               //
/////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.GetSchusterFileInfo(const fName: TFileName; var Module: string);         //
var                                                                                         //
  F: TextFile;                                                                              //
  Str: string;                                                                              //
  P: WORD;                                                                                  //
begin                                                                                       //
  AssignFile(F, fName);                                                                     //
  Reset(F);                                                                                 //
                                                                                            //
  while not EOF(F) do                                                                       //
  begin                                                                                     //
    ReadLn(F, Str);                                                                         //
                                                                                            //
    if Pos('MODULE', UpperCase(Str)) <> 0 then                                              //
    begin                                                                                   //
      P := Pos(#9, Str);                                                                    //
      if P <> 0 then                                                                        //
        Module := Trim(Copy(Str, P+1, Length(Str)));                                        //
                                                                                            //
      Break;                                                                                //
    end;                                                                                    //
  end;                                                                                      //
                                                                                            //
  CloseFile(F);                                                                             //
end;                                                                                        //
//////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.GetSchusterFileInfo(const fName: TFileName; var Module, Config: string); //
var                                                                                         //
  F: TextFile;                                                                              //
  Str: string;                                                                              //
  P: WORD;                                                                                  //
begin                                                                                       //
  AssignFile(F, fName);                                                                     //
  Reset(F);                                                                                 //
                                                                                            //
  while not EOF(F) do                                                                       //
  begin                                                                                     //
    ReadLn(F, Str);                                                                         //
                                                                                            //
    if Pos('MODULE', UpperCase(Str)) <> 0 then                                              //
    begin                                                                                   //
      P := Pos(#9, Str);                                                                    //
      if P <> 0 then                                                                        //
        Module := Trim(Copy(Str, P+1, Length(Str)));                                        //
    end;                                                                                    //
                                                                                            //
    if Pos('CONFIG', UpperCase(Str)) <> 0 then                                              //
    begin                                                                                   //
      P := Pos(#9, Str);                                                                    //
      if P <> 0 then                                                                        //
        Config := Trim(Copy(Str, P+1, Length(Str)));                                        //
                                                                                            //
      Break;                                                                                //
    end;                                                                                    //
  end;                                                                                      //
                                                                                            //
  CloseFile(F);                                                                             //
end;                                                                                        //
//////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////
procedure TMDBForm.CloneBlankWafer(sWaf, dWaf: TWafer);          //
var                                                              //
  X, Y: DWORD;                                                   //
begin                                                            //
  dWaf.OKR        := sWaf.OKR;                                   //
  dWaf.Code       := sWaf.Code;                                  //
  dWaf.MPW        := sWaf.MPW;                                   //
  dWaf.MPWPos     := sWaf.MPWPos;                                //
  dWaf.Device     := sWaf.Device;                                //
  dWaf.DscrDev    := sWaf.DscrDev;                               //
  dWaf.MeasSystem := sWaf.MeasSystem;                            //
  dWaf.Prober     := sWaf.Prober;                                //
  dWaf.Diameter   := sWaf.Diameter;                              //
  dWaf.StepX      := sWaf.StepX;                                 //
  dWaf.StepY      := sWaf.StepY;                                 //
                                                                 //
  dWaf.Cadre.StartX := sWaf.Cadre.StartX;                        //
  dWaf.Cadre.StartY := sWaf.Cadre.StartY;                        //
  dWaf.Cadre.ScaleX := sWaf.Cadre.ScaleX;                        //
  dWaf.Cadre.ScaleY := sWaf.Cadre.ScaleY;                        //
                                                                 //
  dWaf.BaseChip.X := sWaf.BaseChip.X;                            //
  dWaf.BaseChip.Y := sWaf.BaseChip.Y;                            //
                                                                 //
  dWaf.Direct  := sWaf.Direct;                                   //
  dWaf.CutSide := sWaf.CutSide;                                  //
                                                                 //
  dWaf.LDiameter := sWaf.LDiameter;                              //
  dWaf.Radius    := sWaf.Radius;                                 //
  dWaf.LRadius   := sWaf.LRadius;                                //
  dWaf.Chord     := sWaf.Chord;                                  //
                                                                 //
  SetLength(dWaf.ChipN, 0);                                      //
  SetLength(dWaf.ChipN, sWaf.NTotal);                            //
  SetLength(dWaf.Chip, 0, 0);                                    //
  SetLength(dWaf.Chip, Length(sWaf.Chip), Length(sWaf.Chip[0])); //
  for Y := 0 to Length(dWaf.Chip)-1 do                           //
    for X := 0 to Length(dWaf.Chip[0])-1 do                      //
    begin                                                        //
      dWaf.Chip[Y, X].Status := sWaf.Chip[Y, X].Status;          //
      dWaf.Chip[Y, X].ID     := sWaf.Chip[Y, X].ID;              //
      if dWaf.Chip[Y, X].ID > 0 then                             //
        dWaf.ChipN[dWaf.Chip[Y, X].ID-1] := Point(X, Y);         //
    end;                                                         //
end;                                                             //
///////////////////////////////////////////////////////////////////

//////////////////////////////////////////
function TMDBForm.MDBLoaded: Boolean;   //
begin                                   //
  Result := WafersLB.Items.Count > 0;   //
end;                                    //
//////////////////////////////////////////
//////////////////////////////////////////
function TMDBForm.NormsLoaded: Boolean; //
begin                                   //
  Result := Length(Params) > 0;         //
end;                                    //
//////////////////////////////////////////
//////////////////////////////////////////
function TMDBForm.MapLoaded: Boolean;   //
begin                                   //
  Result := Lot.BlankWafer.NTotal > 0;  //
end;                                    //
//////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TMDBForm.GetEnableProcess: Boolean;                                                                     //
begin                                                                                                            //
  Result := True;                                                                                                //
                                                                                                                 //
  if MDBLoaded then                                                                                              //
  begin                                                                                                          //
    if (NoNorms = 0) and (not NormsLoaded()) then Result := False; // ���� ��� ���� ������ � ����� �� ���������  //
  end                                                                                                            //
  else Result := False;                                                                                          //
end;                                                                                                             //
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.OnEvent(const EvenType: TEventType; const Str: string); //
begin                                                                      //
  case EvenType of                                                         //
    evError : Print_Result('  '+Str, clRed);                               //
    evOK    : Print_Result('  '+Str, clTeal);                              //
    evInfo  : Print_Result('  '+Str, clBlack);                             //
    evSave  : Print_Result('  '+Str, clNavy);                              //
    evCreate: Print_Result('  '+Str, clBlue);                              //
  end;                                                                     //
end;                                                                       //
/////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////
procedure TMDBForm.Print_Result(const Str: string; const Col: TColor=clBlack); //
begin                                                                          //
  with ResultRE do                                                             //
  begin                                                                        //
    SelStart := Length(Text);                                                  //
    SelAttributes.Color := Col;                                                //
    SelAttributes.Style := [fsBold];                                           //
    Lines.Add(Str);                                                            //
    Perform(EM_SCROLLCARET, 0, 0);                                             //
    Repaint;                                                                   //
  end;                                                                         //
end;                                                                           //
/////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////
procedure TMDBForm.StartTime();       //
begin                                 //
  QueryPerformanceCounter(BegTime);   //
end;                                  //
////////////////////////////////////////
////////////////////////////////////////
function TMDBForm.StopTime(): Double; //
var                                   //
  Freq, EndTime: Int64;               //
begin                                 //
  QueryPerformanceCounter(EndTime);   //
  QueryPerformanceFrequency(Freq);    //
                                      //
  Result := (EndTime-BegTime)/Freq;   //
end;                                  //
////////////////////////////////////////


end.
