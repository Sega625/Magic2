unit Statistica_le;

interface

uses
  Windows, Forms, Messages, Buttons, System.Classes, SysUtils, ExtCtrls, StdCtrls, Controls, Graphics,
  Inifiles, Spin, Math, XMLDoc, XMLIntf, ComCtrls, ComObj, ADODB, ActiveX, Variants, StrUtils,
  Structs;

type
  TOnEvent = procedure(const EventType: TEventType; const ErrStr: string) of object;

  TWafer = class
  public
    Condition : string; // ������� ��������� (��, �+, �-)
    Direct    : byte;
    CutSide   : byte;   // 1: '������'  2: '�����' 3: '�����' 4: '������'
    OKR       : string; // �������� ����
    Code      : string; // ����� ���������
    MPW       : string; // MPW
    MPWPos    : string; // ������� � MPW
    Device    : string;
    DscrDev   : string; // �������� �������
    MeasSystem: string; // ������������� �������
    Prober    : string; // �������� ���������
    Info      : string;
    NWPlace   : WORD;
    NOperator : string;
    NLot      : string;
    Num       : string;
    NTotal    : WORD;
    NMeased   : WORD;
    NOK       : WORD;
    NFailNC   : WORD;
    NFailSC   : WORD;
    NFailFC   : WORD;
    TimeDate  : string;
    Diameter  : WORD;
    LDiameter : Single;
    Radius    : Single;
    LRadius   : Single;
    Chord     : Single;
    StepX     : Single;
    StepY     : Single;

    fName   : TFileName;
    Cadre   : TCadre;
    BaseChip: TPoint;
    HLChip  : TPoint;

    Chip : TChips;
    ChipN: array of TPoint; // ���������� ����� � ������� ���������

    TestsParams  : TTestsParams;
    CalcsParams  : TCalcsParams;
    StatusNamesSL: TStringList;

    constructor Create(Hndl: THandle);
    destructor  Destroy(); override;

    procedure Init();
    function  LoadSTSHeader(): Boolean;
    function  LoadBlankSTSHeader(): Boolean;
    function  LoadNIHeader() : WORD;
    function  LoadAGLHeader(): Boolean;
    function  AddAGLHeader() : Boolean;

    function  LoadGammaMDB   (const MDBfName, WafName: TFileName): Boolean;
    function  LoadSchusterTXT(const TXTfName: TFileName): Boolean;
    function  LoadBlankTXT   (const TXTfName: TFileName): Boolean;
    function  LoadBlank690   (const TXTfName: TFileName): Boolean;
    function  LoadSTS     (const STSfName: TFileName): Boolean;
    function  AddSTS      (const STSfName: TFileName): Boolean;
    function  LoadBlankSTS(const STSfName: TFileName): Boolean;
    function  SaveSTS     (const STSfName: TFileName): Boolean;
    function  LoadNI      (const TXTfName: TFileName): Boolean;
    function  AddNI       (const TXTfName: TFileName): Boolean;
    function  LoadXML     (const XMLfName: TFileName): Boolean;
    function  AddXML      (const XMLfName: TFileName): Boolean;
    function  LoadAGL     (const AGLfName: TFileName): Boolean;
    function  AddAGL      (const AGLfName: TFileName): Boolean;
    function  DetectXLS   (const XLSfName: TFileName): byte;
    function  LoadXLS     (const XLSfName: TFileName): Boolean;
    function  AddXLS      (const XLSfName: TFileName): Boolean;
    function  LoadXLSPxn  (const XLSfName: TFileName): Boolean;

    function  AddNorms(tParams: TTestsParams): Boolean;

    procedure Normalize(); // ������� ������(�� ��������) �����(�����)
    procedure Rotate();
    procedure CalcChips();
    procedure SetChipsID();
    function  IsWafer(): Boolean; // �������� ��� ������?

    function GetChipParamsStat(const Val, Min, Max: Single): byte;
    function GetStatusName(const Status: WORD): String;
  private
    Handle: THandle;
  end;

  TLot = class
  private
    Handle: THandle;

    fOnEvent: TOnEvent;
  public
    LName: string;
    LfName: TFileName;
//    LMSystem: string; // ������������� �������
    LDevice: string; // ���
    LConfig: string; // Schustera

    Wafer: array of TWafer;
    BlankWafer: TWafer;

    HistGroup: array of THistGroup;

    constructor Create(Hndl: THandle);
    destructor  Destroy(); override;

    procedure Init();
    function  LoadHistGroupsFromIni(): Boolean;
    function  CreateBlankWafer(const NWafer: WORD): Boolean;
    function  SaveXLS(const ToFirstFail, MapByParams: Boolean): Boolean;
    function  GetColorByStatus(const Stat: WORD): TColor;
  published
    property OnEvent: TOnEvent read fOnEvent write fOnEvent;
  end;

implementation


{ TWafer }

////////////////////////////////////////////
constructor TWafer.Create(Hndl: THandle); //
begin                                     //
  Handle := Hndl;                         //
                                          //
  StatusNamesSL := TStringList.Create;    //
                                          //
  HLChip.X := -1;                         //
  HLChip.Y := -1;                         //
end;                                      //
////////////////////////////////////////////
////////////////////////////////////////////
destructor TWafer.Destroy();              //
begin                                     //
  StatusNamesSL.Free();                   //
                                          //
  SetLength(Chip, 0, 0);                  //
                                          //
  inherited;                              //
end;                                      //
////////////////////////////////////////////

///////////////////////////////
procedure TWafer.Init();     //
var                          //
  INIfName: TIniFile;        //
  X, Y: WORD;                //
begin                        //
  fName := '';               //
                             //
  Condition := '-';          //
  Direct  := 0;              //
  CutSide := 0;              //
                             //
  OKR        := '-';         //
  Code       := '0';         //
  MPW        := '-';         //
  MPWPos     := '-';         //
  Device     := '-';         //
  DscrDev    := '-';         //
  MeasSystem := '-';         //
  Prober     := '-';         //
  Info       := '';          //
                             //
  NWPlace   := 0;            //
  NOperator := '';           //
  NLot      := '';           //
  Num       := '';           //
  NTotal    := 0;            //
  NMeased   := 0;            //
  NOK       := 0;            //
  NFailNC   := 0;            //
  NFailSC   := 0;            //
  NFailFC   := 0;            //
  TimeDate  := '';           //
                             //
  Diameter  := 0;            //
  LDiameter := 0;            //
  LRadius   := 0;            //
  Chord     := 0;            //
                             //
  StepX := 0;                //
  StepY := 0;                //
  Cadre.StartX := 0;         //
  Cadre.StartY := 0;         //
  Cadre.ScaleX := 0;         //
  Cadre.ScaleY := 0;         //
                             //
  BaseChip.X := 0;           //
  BaseChip.Y := 0;           //
  HLChip.X   := 0;           //
  HLChip.Y   := 0;           //
                             //
  SetLength(TestsParams, 0); //
  SetLength(CalcsParams, 0); //
  SetLength(Chip, 0, 0);     //
  SetLength(ChipN, 0);       //
                             //
  StatusNamesSL.Clear;       //
end;                         //
///////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadSTSHeader(): Boolean;                                           //
var                                                                                 //
  INIfName: TIniFile;                                                               //
  X, Y, n: WORD;                                                                    //
  P: byte;                                                                          //
  tmpSL: TStringList;                                                               //
  Str: String;                                                                      //
begin                                                                               //
  Result := True;                                                                   //
                                                                                    //
  INIfName := TIniFile.Create(fName);                                               //
  with INIfName do                                                                  //
  begin                                                                             //
    OKR        := ReadString ('Main', 'OKR', '-');                                  //
    Code       := ReadString ('Main', 'Code', '0');                                 //
    MPW        := ReadString ('Main', 'MPW', '-');                                  //
    MPWPos     := ReadString ('Main', 'MPWPos', '-');                               //
    Device     := ReadString ('Main', 'Device', '-');                               //
    DscrDev    := ReadString ('Main', 'DscrDev', '-');                              //
    MeasSystem := ReadString ('Main', 'MSystem', '-');                              //
    Prober     := ReadString ('Main', 'Prober', '-');                               //
                                                                                    //
    Diameter   := ReadInteger('Main', 'Diametr', 0);                                //
    StepX      := ReadInteger('Main', 'ChipSizeX', 0)/1000;                         //
    StepY      := ReadInteger('Main', 'ChipSizeY', 0)/1000;                         //
    NWPlace    := ReadInteger('Main', 'WorkPlace', 0);                              //
    NOperator  := ReadString ('Main', 'Operator', '');                              //
    NLot       := ReadString ('Main', 'Lot', '0');                                  //
    P := Pos('-', NLot);                                                            //
    if P > 0 then Delete(NLot, 1, P);                                               //
    Num        := ReadString ('Main', 'Wafer', '0');                                //
    TimeDate   := ReadString ('Main', 'Date', '00.00.00');                          //
    Condition  := ReadString ('Main', 'Condition', '-');                            //
    Info       := ReadString ('Main', 'Info', '');                                  //
                                                                                    //
    Cadre.StartX := ReadInteger('Add', 'OffsetX', 0);                               //
    Cadre.StartY := ReadInteger('Add', 'OffsetY', 0);                               //
    Cadre.ScaleX := ReadInteger('Add', 'CadreX', 0);                                //
    Cadre.ScaleY := ReadInteger('Add', 'CadreY', 0);                                //
    X := ReadInteger('Add', 'MaxX', 0);                                             //
    Y := ReadInteger('Add', 'MaxY', 0);                                             //
    BaseChip.X := ReadInteger('Add', 'BaseChipX', 0);                               //
    BaseChip.Y := ReadInteger('Add', 'BaseChipY', 0);                               //
    Direct  := ReadInteger('Add', 'Path', 0);                                       //
    CutSide := ReadInteger('Add', 'Cut', 0);                                        //
                                                                                    //
    if (X = 0) or (Y = 0) or (Code = '0') then                                      //
    begin                                                                           //
      Result := False;                                                              //
      Exit;                                                                         //
    end;                                                                            //
                                                                                    //
    SetLength(Chip, 0, 0);                                                          //
    SetLength(Chip, Y, X);                                                          //
    for Y := 0 to Length(Chip)-1 do      // �������                                 //
      for X := 0 to Length(Chip[0])-1 do // ������                                  //
      begin                              //                                         //
        Chip[Y, X].Status := 2;          //                                         //
        Chip[Y, X].ID     := 0;          //                                         //
      end;                               //                                         //
                                                                                    //
                                                                                    //
    ReadSectionValues('StatusNames', StatusNamesSL);                                //
                                                                                    //
    tmpSL := TStringList.Create;                                                    //
    ReadSectionValues('TestsParams', tmpSL);                                        //
                                                                                    //
    if tmpSL.Count > 0 then                                                         //
    begin                                                                           //
      SetLength(TestsParams, tmpSL.Count);                                          //
                                                                                    //
      for n := 0 to tmpSL.Count-1 do                                                //
      begin                                                                         //
        P := Pos('=', tmpSL.Strings[n]);                                            //
        if P <> 0 then                                                              //
        begin                                                                       //
          Str := Trim(Copy(tmpSL.Strings[n], P+1, Length(tmpSL.Strings[n])-P));     //
          P := Pos(';', Str);                                                       //
          if P <> 0 then                                                            //
          begin                                                                     //
            try                                                                     //
              FormatSettings.DecimalSeparator := ',';                               //
              TestsParams[n].Norma.Min := StrToFloat(Trim(Copy(Str, 1, P-1)));      //
              FormatSettings.DecimalSeparator := '.';                               //
            except                                                                  //
              try                                                                   //
                FormatSettings.DecimalSeparator := '.';                             //
                TestsParams[n].Norma.Min := StrToFloat(Trim(Copy(Str, 1, P-1)));    //
              except                                                                //
                TestsParams[n].Norma.Min := -NotSpec;                               //
              end                                                                   //
            end;                                                                    //
            System.Delete(Str, 1, P);                                               //
            P := Pos(';', Str);                                                     //
            if P <> 0 then                                                          //
            begin                                                                   //
              try                                                                   //
                FormatSettings.DecimalSeparator := ',';                             //
                TestsParams[n].Norma.Max := StrToFloat(Trim(Copy(Str, 1, P-1)));    //
                FormatSettings.DecimalSeparator := '.';                             //
              except                                                                //
                try                                                                 //
                  FormatSettings.DecimalSeparator := '.';                           //
                  TestsParams[n].Norma.Max := StrToFloat(Trim(Copy(Str, 1, P-1)));  //
                except                                                              //
                   TestsParams[n].Norma.Max := NotSpec;                             //
                end;                                                                //
              end;                                                                  //
              System.Delete(Str, 1, P);                                             //
              try                                                                   //
                TestsParams[n].Status := StrToInt(Trim(Copy(Str, 1, Length(Str)))); //
              except                                                                //
                TestsParams[n].Status := 0;                                         //
              end;                                                                  //
            end;                                                                    //
          end;                                                                      //
        end;                                                                        //
      end;                                                                          //
    end;                                                                            //
    tmpSL.Free;                                                                     //
                                                                                    //
    Free;                                                                           //
  end;                                                                              //
                                                                                    //
  LDiameter := Diameter;                                                            //
  if Diameter = 150 then LDiameter := 144.25;                                       //
  Radius  := Diameter/2;                                                            //
  LRadius := Radius-(Diameter-LDiameter);                                           //
  Chord   := Sqrt(Radius*Radius-LRadius*LRadius);                                   //
end;                                                                                //
//////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////
function TWafer.LoadBlankSTSHeader(): Boolean;              //
var                                                         //
  INIfName: TIniFile;                                       //
  X, Y: WORD;                                               //
begin                                                       //
  Result := True;                                           //
                                                            //
  INIfName := TIniFile.Create(fName);                       //
  with INIfName do                                          //
  begin                                                     //
    OKR        := ReadString ('Main', 'OKR',     '' );      //
    Code       := ReadString ('Main', 'Code',    '0');      //
    MPW        := ReadString ('Main', 'MPW',     '' );      //
    MPWPos     := ReadString ('Main', 'MPWPos',  '' );      //
    Device     := ReadString ('Main', 'Device',  '' );      //
    DscrDev    := ReadString ('Main', 'DscrDev', '' );      //
    MeasSystem := ReadString ('Main', 'MSystem', '' );      //
    Prober     := ReadString ('Main', 'Prober',  '' );      //
                                                            //
    Diameter   := ReadInteger('Main', 'Diametr',   0);      //
    StepX      := ReadInteger('Main', 'ChipSizeX', 0)/1000; //
    StepY      := ReadInteger('Main', 'ChipSizeY', 0)/1000; //
                                                            //
    Cadre.StartX := ReadInteger('Add', 'OffsetX', 0);       //
    Cadre.StartY := ReadInteger('Add', 'OffsetY', 0);       //
    Cadre.ScaleX := ReadInteger('Add', 'CadreX',  0);       //
    Cadre.ScaleY := ReadInteger('Add', 'CadreY',  0);       //
    X := ReadInteger('Add', 'MaxX', 0);                     //
    Y := ReadInteger('Add', 'MaxY', 0);                     //
    BaseChip.X := ReadInteger('Add', 'BaseChipX', 0);       //
    BaseChip.Y := ReadInteger('Add', 'BaseChipY', 0);       //
    Direct  := ReadInteger('Add', 'Path', 0);               //
    CutSide := ReadInteger('Add', 'Cut',  0);               //
                                                            //
    if (X = 0) or (Y = 0) or (Code = '0') then              //
    begin                                                   //
      Result := False;                                      //
      Exit;                                                 //
    end;                                                    //
                                                            //
    SetLength(Chip, 0, 0);               //                 //
    SetLength(Chip, Y, X);               //                 //
    for Y := 0 to Length(Chip)-1 do      // �������         //
      for X := 0 to Length(Chip[0])-1 do // ������          //
      begin                              //                 //
        Chip[Y, X].Status := 2;          //                 //
        Chip[Y, X].ID     := 0;          //                 //
      end;                               //                 //
                                                            //
    Free;                                                   //
  end;                                                      //
                                                            //
  LDiameter := Diameter;                                    //
  if Diameter = 150 then LDiameter := 144.25;               //
  Radius  := Diameter/2;                                    //
  LRadius := Radius-(Diameter-LDiameter);                   //
  Chord   := Sqrt(Radius*Radius-LRadius*LRadius);           //
end;                                                        //
//////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////
function TWafer.LoadNIHeader(): WORD;                       //
var                                                         //
  SL: TStringList;                                          //
  n, TotalChips: DWORD;                                     //
  P: byte;                                                  //
  X, Y: WORD;                                               //
  NumChip, PrevChip, Str, tmpStr: AnsiString;               //
  FirstTime: Boolean;                                       //
begin                                                       //
  Result := 0;                                              //
                                                            //
  SL := TStringList.Create;                                 //
  SL.LoadFromFile(fName);                                   //
                                                            //
  FirstTime := True;                                        //
  X := 0;
  TotalChips := 1;
  for n := 0 to SL.Count-1 do
  begin
    if Trim(SL.Strings[n]) = '' then Continue;
//    if Trim(SL.Strings[n])[1] = '/' then Continue;
    Inc(Result);

//    P := Pos('=', SL.Strings[n]);
//    if P <> 0 then
    tmpStr := Trim(SL.Strings[n]);
    try StrToInt(tmpStr[1])
    except
      begin
        P := Pos('=', SL.Strings[n]);
        if Pos('�������',       SL.Strings[n]) <> 0 then Device   := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))));
        if Pos('����',          SL.Strings[n]) <> 0 then TimeDate := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))));
  //      if Pos('�����', SL.Strings[n])         <> 0 then TimeDate := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))));
        if Pos('��� ���������', SL.Strings[n]) <> 0 then info := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))))+' ';
        if Pos('�������',       SL.Strings[n]) <> 0 then Condition := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))))+' ';
        if Pos('��������', SL.Strings[n])      <> 0 then NOperator := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))));

        Continue;
      end;
    end;

    Dec(Result);

    Str := SL.Strings[n];
//    if Pos(',', Str) = 0 then DecimalSeparator := '.'
//                         else DecimalSeparator := ',';
    NumChip := Trim(Copy(Str, 1, Pos(#9, Str)));
    if FirstTime then
    begin
      PrevChip := NumChip;
      FirstTime := False;
    end;
    if NumChip <> PrevChip then
    begin
      X := 0;
      Inc(TotalChips);
    end;
    if X >= Length(TestsParams) then SetLength(TestsParams, X+1);
    tmpStr := '';
    Delete(Str, 1, Pos(#9, Str)); // ������ ����� ���������
    Delete(Str, 1, Pos(#9, Str)); // ������ ����� �����
    tmpStr := Trim(Copy(Str, 1, Pos(#9, Str))); // �������� ��� ���������
    Delete(Str, 1, Pos(#9, Str)); // ������ �������� ���������
    Delete(Str, 1, Pos(#9, Str)); // ������ ��������
    tmpStr := tmpStr+' ('+Trim(Copy(Str, 1, Pos(#9, Str)))+')'; // �������� ������ ��� ���������
    TestsParams[X].Name := tmpStr;
    Delete(Str, 1, Pos(#9, Str)); // ������ ��������
    Delete(Str, 1, Pos(#9, Str)); // ������ ����� �����

    tmpStr := Trim(Copy(Str, 1, Pos(#9, Str))); // ������� ������ ������
    if Pos('.', tmpStr) <> 0 then FormatSettings.DecimalSeparator := '.'
                             else FormatSettings.DecimalSeparator := ',';
    try
      TestsParams[X].Norma.Min := StrToFloat(tmpStr); // ������� ������ ������
//      DecimalSeparator := ',';
      tmpStr := FormatFloat('0.000', TestsParams[X].Norma.Min);
//      DecimalSeparator := '.';
    except
      TestsParams[X].Norma.Min := -NotSpec;
      tmpStr := tmpStr+'N';
    end;
    Delete(Str, 1, Pos(#9, Str)); // ������ ���. �����
    Delete(Str, 1, Pos(#9, Str)); // ������ ���. �����

    try
      TestsParams[X].Norma.Max := StrToFloat(Trim(Copy(Str, 1, Pos(#9, Str)))); // ������� ������� ������
//      DecimalSeparator := ',';
      tmpStr := tmpStr+';'+FormatFloat('0.000', TestsParams[X].Norma.Max);
//      DecimalSeparator := '.';
    except
      TestsParams[X].Norma.Max := NotSpec;
      tmpStr := tmpStr+';N';
    end;
    TestsParams[X].Status := 2000+X;
//    if TestsParamsSL.Count < Length(TestsParams) then TestsParamsSL.Add(IntToStr(TestsParamsSL.Count)+'='+tmpStr+';'+IntToStr(TestsParams[X].Status));

    PrevChip := NumChip;

    Inc(X);
  end;

  MeasSystem := 'NI';

  X := Ceil(sqrt(TotalChips));
  Y := X;

  SetLength(Chip, 0, 0);
  SetLength(Chip, Y, X);
    for Y := 0 to Length(Chip)-1 do      // �������
      for X := 0 to Length(Chip[0])-1 do // ������
      begin                              //
        Chip[Y, X].Status := 2;          //
        Chip[Y, X].ID     := 0;          //
        SetLength(Chip[Y, X].ChipParams, Length(TestsParams));
      end;                               //
  Direct := 2;

  SL.Free;
end;                                                        //
//////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////
function TWafer.LoadAGLHeader(): Boolean;                   //
var                                                         //
  SL, NKKSL, NFCSL: TStringList;                            //
  n, TotalChips: DWORD;                                     //
  P, eP: byte;                                              //
  X, Y, NFC, NKK: WORD;                                     //
  Str, tmpStr1, tmpStr2: String;                            //
begin                                                       //
  Result := True;                                           //

  NKKSL := TStringList.Create;
  NFCSL := TStringList.Create;
  SL := TStringList.Create;
  SL.LoadFromFile(fName);

//  X := 0;
  TotalChips := 0;
  for n := 0 to SL.Count-1 do
  begin
    Str := Trim(SL.Strings[n]);

    if Str = '' then Continue;

    if Pos('TESTFLOW STARTED', UpperCase(Str)) <> 0 then
    begin
      NFC := 0;
      NKK := 0;
      X := 0;
      Inc(TotalChips);

      if TotalChips = 1 then
      begin
        Str := UpperCase(Str);
        P := Pos('ON', Str)+2;
        Str := Trim(Copy(Str, P, Pos('AT', Str)-P));
        TimeDate := Copy(Str, 4, 2)+'.'+Copy(Str, 1, 2)+'.'+Copy(Str, 7, Length(Str)-6);
      end;

      Continue;
    end;

    if Str[1] = '1' then // ���� ��������
    begin
      Delete(Str, 1, Pos('`', Str)); // ������ ����� �����
      P := Pos('`', Str);
      eP := PosEx('`', Str, P+1);
      tmpStr1 := Trim(Copy(Str, 1, P-1));      // �������� ��� ���������
      tmpStr2 := Trim(Copy(Str, P+1, eP-P-1)); // �������� ������ ��� ���������

      if (Pos('CONTINUITY', UpperCase(Str)) <> 0) or
         (Pos('CONTAKT',    UpperCase(Str)) <> 0) or
         (Pos('CONTACT',    UpperCase(Str)) <> 0) then
      begin
        if NKK >= NKKSL.Count then
        begin
          NKKSL.Add(IntToStr(10+NKK)+'='+tmpStr2);
          Inc(NKK);
        end;
        Continue;
      end;
      if (Pos('FUNCTIONAL', UpperCase(Str)) <> 0) or
         (Pos('FUNCTION',   UpperCase(Str)) <> 0) or
         (Pos('FK',         UpperCase(Str)) <> 0) then
      begin
        if NFC >= NFCSL.Count then
        begin
          NFCSL.Add(IntToStr(3500+NFC)+'='+tmpStr1);
          Inc(NFC);
        end;
        Continue;
      end;

      Inc(X);
      if X > Length(TestsParams) then // ���� �� ������ ������ ���������� - ��������
      begin
        SetLength(TestsParams, X);
        Delete(Str, 1, Pos('`', Str)); // ������ �������� ���������
        //tmpStr := tmpStr+' - '+Trim(Copy(Str, 1, Pos('`', Str)-1)); // �������� ������ ��� ���������
        StatusNamesSL.Add(IntToStr(2000+X-1)+'='+Trim(Copy(Str, 1, Pos('`', Str)-1)));
        Delete(Str, 1, Pos('`', Str)); // ������ ������ ��� ���������
        TestsParams[X-1].Name := tmpStr1+' - '+tmpStr2; // ������� ������ ��� ���������

        Delete(Str, 1, Pos('`', Str)); // ������ passed/FAILED
        Str := Trim(Str);
        try
          TestsParams[X-1].Norma.Min := StrToFloat(Trim(Copy(Str, 1, Pos(' ', Str)-1))); // ������� ������ ������
          FormatSettings.DecimalSeparator := ',';
          tmpStr1 := FormatFloat('0.000', TestsParams[X-1].Norma.Min);
          FormatSettings.DecimalSeparator := '.';
        except
          TestsParams[X-1].Norma.Min := -NotSpec;
          tmpStr1 := tmpStr1+'N';
        end;
        Delete(Str, 1, Pos(' ', Str)); // ������ ������ ������
        TestsParams[X-1].Name := TestsParams[X-1].Name+' ('+Trim(Copy(Str, 1, Pos(' ', Str)-1)+')'); // ������� ������� ���������
        Delete(Str, 1, Pos('`', Str)); // ������ ������� ��������� � ���������
        Delete(Str, 1, Pos('`', Str)); // ������ ��������
        Str := Trim(Str);
        Delete(Str, 1, Pos(' ', Str)); // ������ ���������
        Str := Trim(Str);
        try
          TestsParams[X-1].Norma.Max := StrToFloat(Trim(Copy(Str, 1, Pos(' ', Str)-1))); // ������� ������� ������
          FormatSettings.DecimalSeparator := ',';
          tmpStr1 := tmpStr1+';'+FormatFloat('0.000', TestsParams[X-1].Norma.Max);
          FormatSettings.DecimalSeparator := '.';
        except
          TestsParams[X-1].Norma.Max := NotSpec;
          tmpStr1 := tmpStr1+';N';
        end;
        TestsParams[X-1].Status := 2000+X-1;
      end;
    end;
  end;
  if NKKSL.Count > 0 then for n := 0 to NKKSL.Count-1 do StatusNamesSL.Add(NKKSL.Strings[n]);
  if NFCSL.Count > 0 then for n := 0 to NFCSL.Count-1 do StatusNamesSL.Add(NFCSL.Strings[n]);

  MeasSystem := 'Verigy93K';
  Direct := 2;

  X := Ceil(sqrt(TotalChips));
  Y := X;

  SetLength(Chip, 0, 0);
  SetLength(Chip, Y, X);
    for Y := 0 to Length(Chip)-1 do      // �������
      for X := 0 to Length(Chip[0])-1 do // ������
      begin                              //
        Chip[Y, X].Status := 2;          //
        Chip[Y, X].ID     := 0;          //
        SetLength(Chip[Y, X].ChipParams, Length(TestsParams));
      end;                               //

  SL.Free;
  NKKSL.Free;
  NFCSL.Free;
end;

//////////////////////////////////////////////////////////////
function TWafer.AddAGLHeader(): Boolean;                    //
var                                                         //
  SL, NKKSL, NFCSL: TStringList;                            //
  n, TotalChips: DWORD;                                     //
  P, eP: byte;                                              //
  X, Y, NFC, NKK: WORD;                                     //
  Str, tmpStr1, tmpStr2: String;                            //
begin                                                       //
  Result := True;                                           //

  NKKSL := TStringList.Create;
  NFCSL := TStringList.Create;
  SL := TStringList.Create;
  SL.LoadFromFile(fName);


//  X := 0;
  TotalChips := 0;
  for n := 0 to SL.Count-1 do
  begin
    Str := Trim(SL.Strings[n]);

    if Str = '' then Continue;

    if Pos('TESTFLOW STARTED', UpperCase(Str)) <> 0 then
    begin
      NFC := 0;
      NKK := 0;
      X := 0;
      Inc(TotalChips);

      if TotalChips = 1 then
      begin
        Str := UpperCase(Str);
        P := Pos('ON', Str)+2;
        Str := Trim(Copy(Str, P, Pos('AT', Str)-P));
        TimeDate := Copy(Str, 4, 2)+'.'+Copy(Str, 1, 2)+'.'+Copy(Str, 7, Length(Str)-6);
      end;

      Continue;
    end;

    if Str[1] = '1' then // ���� ��������
    begin
      Delete(Str, 1, Pos('`', Str)); // ������ ����� �����
      P := Pos('`', Str);
      eP := PosEx('`', Str, P+1);
      tmpStr1 := Trim(Copy(Str, 1, P-1));      // �������� ��� ���������
      tmpStr2 := Trim(Copy(Str, P+1, eP-P-1)); // �������� ������ ��� ���������

      if (Pos('CONTINUITY', UpperCase(Str)) <> 0) or
         (Pos('CONTAKT',    UpperCase(Str)) <> 0) or
         (Pos('CONTACT',    UpperCase(Str)) <> 0) then
      begin
        if NKK >= NKKSL.Count then
        begin
          NKKSL.Add(IntToStr(10+NKK)+'='+tmpStr2);
          Inc(NKK);
        end;
        Continue;
      end;
      if (Pos('FUNCTIONAL', UpperCase(Str)) <> 0) or
         (Pos('FUNCTION',   UpperCase(Str)) <> 0) or
         (Pos('FK',         UpperCase(Str)) <> 0) then
      begin
        if NFC >= NFCSL.Count then
        begin
          NFCSL.Add(IntToStr(3500+NFC)+'='+tmpStr1);
          Inc(NFC);
        end;
        Continue;
      end;

      Inc(X);
      if X > Length(TestsParams) then // ���� �� ������ ������ ���������� - ��������
      begin
        SetLength(TestsParams, X);
        Delete(Str, 1, Pos('`', Str)); // ������ �������� ���������
        //tmpStr := tmpStr+' - '+Trim(Copy(Str, 1, Pos('`', Str)-1)); // �������� ������ ��� ���������
        StatusNamesSL.Add(IntToStr(2000+X-1)+'='+Trim(Copy(Str, 1, Pos('`', Str)-1)));
        Delete(Str, 1, Pos('`', Str)); // ������ ������ ��� ���������
        TestsParams[X-1].Name := tmpStr2; // ������� ������ ��� ���������

        Delete(Str, 1, Pos('`', Str)); // ������ passed/FAILED
        Str := Trim(Str);
        try
          TestsParams[X-1].Norma.Min := StrToFloat(Trim(Copy(Str, 1, Pos(' ', Str)-1))); // ������� ������ ������
          FormatSettings.DecimalSeparator := ',';
          tmpStr1 := FormatFloat('0.000', TestsParams[X-1].Norma.Min);
          FormatSettings.DecimalSeparator := '.';
        except
          TestsParams[X-1].Norma.Min := -NotSpec;
          tmpStr1 := tmpStr1+'N';
        end;
        Delete(Str, 1, Pos(' ', Str)); // ������ ������ ������
        TestsParams[X-1].Name := TestsParams[X-1].Name+' ('+Trim(Copy(Str, 1, Pos(' ', Str)-1)+')'); // ������� ������� ���������
        Delete(Str, 1, Pos('`', Str)); // ������ ������� ��������� � ���������
        Delete(Str, 1, Pos('`', Str)); // ������ ��������
        Str := Trim(Str);
        Delete(Str, 1, Pos(' ', Str)); // ������ ���������
        Str := Trim(Str);
        try
          TestsParams[X-1].Norma.Max := StrToFloat(Trim(Copy(Str, 1, Pos(' ', Str)-1))); // ������� ������� ������
          FormatSettings.DecimalSeparator := ',';
          tmpStr1 := tmpStr1+';'+FormatFloat('0.000', TestsParams[X-1].Norma.Max);
          FormatSettings.DecimalSeparator := '.';
        except
          TestsParams[X-1].Norma.Max := NotSpec;
          tmpStr1 := tmpStr1+';N';
        end;
        TestsParams[X-1].Status := 2000+X-1;
      end;
    end;
  end;
  if NKKSL.Count > 0 then for n := 0 to NKKSL.Count-1 do StatusNamesSL.Add(NKKSL.Strings[n]);
  if NFCSL.Count > 0 then for n := 0 to NFCSL.Count-1 do StatusNamesSL.Add(NFCSL.Strings[n]);

  MeasSystem := 'Verigy93K';

//  MessageBox(0, PAnsiChar(IntToStr(Length(TestsParams))), PAnsiChar(IntToStr(Length(TestsParams))), 0);

  for Y := 0 to Length(Chip)-1 do      // �������
    for X := 0 to Length(Chip[0])-1 do // ������
    begin                              // �� �����
      if Chip[Y, X].Status > 9 then Chip[Y, X].Status := 0;
      SetLength(Chip[Y, X].ChipParams, Length(TestsParams));
    end;

  SL.Free;
  NKKSL.Free;
  NFCSL.Free;
end;


////////////////////////////////////////////////////////////////////////
procedure TWafer.Normalize();                                         //
label                                                                 //
  X0_End, X1_End, Y0_End, Y1_End; // ������������                     //
var                                                                   //
  tmpChip: TChips;                                                    //
  X, Y, X0, X1, Y0, Y1: WORD;                                         //
begin                                                                 //
  X0 := 0;                                                            //
  for X := 0 to Length(Chip[0])-1 do                                  //
    for Y := 0 to Length(Chip)-1 do                                   //
      if Chip[Y, X].Status <> 2 then                                  //
      begin                                                           //
        X0 := X;                                                      //
        Goto X0_End;                                                  //
      end;                                                            //
X0_End:                                                               //
                                                                      //
  X1 := 0;                                                            //
  for X := Length(Chip[0])-1 downto 0 do                              //
    for Y := 0 to Length(Chip)-1 do                                   //
      if Chip[Y, X].Status <> 2 then                                  //
      begin                                                           //
        X1 := X;                                                      //
        Goto X1_End;                                                  //
      end;                                                            //
X1_End:                                                               //
                                                                      //
  Y0 := 0;                                                            //
  for Y := 0 to Length(Chip)-1 do                                     //
    for X := 0 to Length(Chip[0])-1 do                                //
      if Chip[Y, X].Status <> 2 then                                  //
      begin                                                           //
        Y0 := Y;                                                      //
        Goto Y0_End;                                                  //
      end;                                                            //
Y0_End:                                                               //
                                                                      //
  Y1 := 0;                                                            //
  for Y := Length(Chip)-1 downto 0 do                                 //
    for X := 0 to Length(Chip[0])-1 do                                //
      if Chip[Y, X].Status <> 2 then                                  //
      begin                                                           //
        Y1 := Y;                                                      //
        Goto Y1_End;                                                  //
      end;                                                            //
Y1_End:                                                               //
                                                                      //
  SetLength(tmpChip, Y1-Y0+1, X1-X0+1);                               //
  for Y := Y0 to Y1 do                                                //
    for X := X0 to X1 do                                              //
    begin                                                             //
      tmpChip[Y-Y0, X-X0] := Chip[Y, X];                              //
//      SetLength(tmpChip[Y-Y0, X-X0].Value, Length(Chip[Y, X].Value)); //
//      if Length(tmpChip[Y-Y0, X-X0].Value) > 0 then                   //
//        for n := 0 to Length(tmpChip[Y-Y0, X-X0].Value)-1 do          //
//          tmpChip[Y-Y0, X-X0].Value[n] := Chip[Y, X].Value[n];        //
    end;                                                              //
                                                                      //
  Chip := tmpChip;                                                    //
//  SetLength(Chip, Length(TmpChip), Length(TmpChip[0]));               //
//  tmpChip := nil;                                                     //
                                                                      //
  BaseChip.X := BaseChip.X-X0;                                        //
  BaseChip.Y := BaseChip.Y-Y0;                                        //
end;                                                                  //
////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////
procedure TWafer.Rotate();                                                                  //
var                                                                                         //
  X, Y: WORD;                                                                               //
  TmpChip: TChips;                                                                          //
  TmpInt: Integer;                                                                          //
  TmpSingle: Single;                                                                        //
begin                                                                                       //
  TmpInt := BaseChip.X;                                                                     //
  BaseChip.X := BaseChip.Y;                                                                 //
  BaseChip.Y := Length(Chip[0])-TmpInt-1;                                                   //
                                                                                            //
  TmpInt := Cadre.StartX;                                                                   //
  Cadre.StartX := Cadre.StartY;                                                             //
  if Cadre.ScaleX <> 0 then Cadre.StartY := (Length(Chip[0])-TmpInt) mod Cadre.ScaleX       //
                       else Cadre.StartY := 0;                                              //
                                                                                            //
  TmpInt := Cadre.ScaleX;                                                                   //
  Cadre.ScaleX := Cadre.ScaleY;                                                             //
  Cadre.ScaleY := TmpInt;                                                                   //
                                                                                            //
  SetLength(TmpChip, Length(Chip[0]), Length(Chip));                                        //
                                                                                            //
  for Y := 0 to Length(Chip)-1 do                                                           //
    for X := 0 to Length(Chip[0])-1 do                                                      //
    begin                                                                                   //
      TmpChip[Length(Chip[0])-X-1, Y] := Chip[Y, X];                                        //
      SetLength(TmpChip[Length(Chip[0])-X-1, Y].ChipParams, Length(Chip[Y, X].ChipParams)); //
    end;                                                                                    //
                                                                                            //
  Chip := TmpChip;                                                                          //
  SetLength(Chip, Length(TmpChip), Length(TmpChip[0]));                                     //
  TmpChip := nil;                                                                           //
                                                                                            //
  if CutSide <> 0 then                                                                      //
    if CutSide < 4 then Inc(CutSide) else CutSide := 1;                                     //
                                                                                            //
  if Direct > 11 then Direct := Direct-12                                                   //
                 else Inc(Direct, 4);                                                       //
                                                                                            //
  if (StepX <> 0) and (StepY <> 0) then                                                     //
  begin                                                                                     //
    TmpSingle := StepX;                                                                     //
    StepX := StepY;                                                                         //
    StepY := TmpSingle;                                                                     //
  end;                                                                                      //
                                                                                            //
  SetChipsID;                                                                               //
end;                                                                                        //
//////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////
procedure TWafer.CalcChips();                                 //
                                                              //
//////////////////////////////////////////                    //
procedure SortFails(var Fails: TFails); //                    //
var                                     //                    //
  n, m, b_val, b_m: WORD;               //                    //
  TmpFail: TFail;                       //                    //
begin                                   //                    //
  if Length(Fails) < 2 then Exit;       //                    //
                                        //                    //
  for n := 0 to Length(Fails)-2 do      //                    //
  begin                                 //                    //
    b_val := Fails[n].Status;           //                    //
    b_m := n;                           //                    //
    for m := n+1 to Length(Fails)-1 do  //                    //
      if Fails[m].Status < b_val then   //                    //
      begin                             //                    //
        b_val := Fails[m].Status;       //                    //
        b_m := m;                       //                    //
      end;                              //                    //
    TmpFail := Fails[b_m];              //                    //
    Fails[b_m] := Fails[n];             //                    //
    Fails[n] := TmpFail;                //                    //
  end;                                  //                    //
end;                                    //                    //
//////////////////////////////////////////                    //
                                                              //
var                                                           //
  X, Y: WORD;                                                 //
begin                                                         //
  if Length(Chip) = 0 then Exit;                              //
                                                              //
  NOK     := 0;                                               //
  NFailNC := 0;                                               //
  NFailSC := 0;                                               //
  NFailFC := 0;                                               //
  NTotal  := 0;                                               //
  NMeased := 0;                                               //
  for Y := 0 to Length(Chip)-1 do                             //
    for X := 0 to Length(Chip[0])-1 do                        //
    begin                                                     //
      if not (Chip[Y, X].Status in [2,3,5]) then Inc(NTotal); //
                                                              //
      case Chip[Y, X].Status of                               //
        1         : begin                                     //
                      Inc(NOK);                               //
                      Inc(NMeased);                           //
                    end;                                      //
                                                              //
        10..1500  : begin                                     //
                      Inc(NFailNC);                           //
                      Inc(NMeased);                           //
                    end;                                      //
                                                              //
        2000..3000: begin                                     //
                      Inc(NFailSC);                           //
                      Inc(NMeased);                           //
                    end;                                      //
                                                              //
        3500..4500: begin                                     //
                      Inc(NFailFC);                           //
                      Inc(NMeased);                           //
                    end;                                      //
      end;                                                    //
    end;                                                      //
end;                                                          //
////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////
procedure TWafer.SetChipsID();                                                    //
var                                                                               //
  N: DWORD;                                                                       //
  X, Y, XY: WORD;                                                                 //
  tmp: byte;                                                                      //
  MassXY: array of WORD; // ������ �� ������ �����                                //
begin                                                                             //
  if Length(Chip) = 0 then Exit;                                                  //
                                                                                  //
  N := 0;                                                                         //
                                                                                  //
  if Direct in [0,1,2,3,8,9,10,11] then // �������������� �����                   //
  begin                                                                           //
    SetLength(MassXY, Length(Chip));                                              //
    for Y := 0 to Length(Chip)-1 do                                               //
      for X := 0 to Length(Chip[0])-1 do                                          //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          if N = 0 then                                                           //
          begin                                                                   //
            MassXY[N] := Y;                                                       //
            Inc(N);                                                               //
          end                                                                     //
          else                                                                    //
            if MassXY[N-1] <> Y then                                              //
            begin                                                                 //
              MassXY[N] := Y;                                                     //
              Inc(N);                                                             //
            end;                                                                  //
        end;                                                                      //
  end                                                                             //
  else                                  // ������������ �����                     //
  begin                                                                           //
    SetLength(MassXY, Length(Chip[0]));                                           //
    for X := 0 to Length(Chip[0])-1 do                                            //
      for Y := 0 to Length(Chip)-1 do                                             //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          if N = 0 then                                                           //
          begin                                                                   //
            MassXY[N] := X;                                                       //
            Inc(N);                                                               //
          end                                                                     //
          else                                                                    //
            if MassXY[N-1] <> X then                                              //
            begin                                                                 //
              MassXY[N] := X;                                                     //
              Inc(N);                                                             //
            end;                                                                  //
        end;                                                                      //
  end;                                                                            //
                                                                                  //
  SetLength(MassXY, N);                                                           //
                                                                                  //
  N := 0;                                                                         //
  SetLength(ChipN, 0);                                                            //
  SetLength(ChipN, Length(Chip[0])*Length(Chip));                                 //
                                                                                  //
////////////////////////// * ������ ������ (������) * //////////////////////////////
                                                                                  //
  if Direct = dURightToLeft then                                                  //
    for XY := 0 to Length(MassXY)-1 do                                            //
    begin                                                                         //
      Y := MassXY[XY];                                                            //
      for X := Length(Chip[0])-1 downto 0 do                                      //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
/////////////////////////// * ����� ������� (������) * /////////////////////////////
                                                                                  //
  if Direct = dULeftToRight then                                                  //
    for XY := 0 to Length(MassXY)-1 do                                            //
    begin                                                                         //
      Y := MassXY[XY];                                                            //
      for X := 0 to Length(Chip[0])-1 do                                          //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
////////////////////// * ������ � ����� ������ (������) * //////////////////////////
                                                                                  //
  if Direct in [dURightSnake, dULeftSnake] then                                   //
  begin                                                                           //
    if Direct = dURightSnake then tmp := 1                                        //
                             else tmp := 0;                                       //
    for XY := 0 to Length(MassXY)-1 do                                            //
    begin                                                                         //
      Y := MassXY[XY];                                                            //
      if (XY mod 2) = tmp then                                                    //
      begin                                                                       //
        for X := 0 to Length(Chip[0])-1 do                                        //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
      end                                                                         //
      else                                                                        //
        for X := Length(Chip[0])-1 downto 0 do                                    //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
    end;                                                                          //
  end;                                                                            //
                                                                                  //
////////////////////////////////////////////////////////////////////////////////////
                                                                                  //
//////////////////////////// * ������ ���� (�����) * ///////////////////////////////
                                                                                  //
  if Direct = dLUpToDown then                                                     //
    for XY := 0 to Length(MassXY)-1 do                                            //
    begin                                                                         //
      X := MassXY[XY];                                                            //
      for Y := 0 to Length(Chip)-1 do                                             //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
//////////////////////////// * ����� ����� (�����) * ///////////////////////////////
                                                                                  //
  if Direct = dLDownToUp then                                                     //
    for XY := 0 to Length(MassXY)-1 do                                            //
    begin                                                                         //
      X := MassXY[XY];                                                            //
      for Y := Length(Chip)-1 downto 0 do                                         //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
////////////////////// * ������� � ������ ������ (�����) * /////////////////////////
                                                                                  //
  if Direct in [dLUpSnake, dLDownSnake] then                                      //
  begin                                                                           //
    if Direct = dLUpSnake then tmp := 0                                           //
                          else tmp := 1;                                          //
    for XY := 0 to Length(MassXY)-1 do                                            //
    begin                                                                         //
      X := MassXY[XY];                                                            //
      if (XY mod 2) = tmp then                                                    //
      begin                                                                       //
        for Y := 0 to Length(Chip)-1 do                                           //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
      end                                                                         //
      else                                                                        //
        for Y := Length(Chip)-1 downto 0 do                                       //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
    end;                                                                          //
  end;                                                                            //
                                                                                  //
////////////////////////////////////////////////////////////////////////////////////
                                                                                  //
/////////////////////////// * ������ ������ (�����) * //////////////////////////////
                                                                                  //
  if Direct = dDRightToLeft then                                                  //
    for XY := Length(MassXY)-1 downto 0 do                                        //
    begin                                                                         //
      Y := MassXY[XY];                                                            //
      for X := Length(Chip[0])-1 downto 0 do                                      //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
/////////////////////////// * ����� ������� (�����) * //////////////////////////////
                                                                                  //
  if Direct = dDLeftToRight then                                                  //
    for XY := Length(MassXY)-1 downto 0 do                                        //
    begin                                                                         //
      Y := MassXY[XY];                                                            //
      for X := 0 to Length(Chip[0])-1 do                                          //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
/////////////////////// * ������ � ����� ������ (�����) * //////////////////////////
                                                                                  //
  if Direct in [dDRightSnake, dDLeftSnake] then                                   //
  begin                                                                           //
    if Direct = dDRightSnake then tmp := 0                                        //
                             else tmp := 1;                                       //
    for XY := Length(MassXY)-1 downto 0 do                                        //
    begin                                                                         //
      Y := MassXY[XY];                                                            //
      if (XY mod 2) = tmp then                                                    //
      begin                                                                       //
        for X := 0 to Length(Chip[0])-1 do                                        //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
      end                                                                         //
      else                                                                        //
        for X := Length(Chip[0])-1 downto 0 do                                    //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
    end;                                                                          //
  end;                                                                            //
                                                                                  //
////////////////////////////////////////////////////////////////////////////////////
                                                                                  //
/////////////////////////// * ������ ���� (������) * ///////////////////////////////
                                                                                  //
  if Direct = dRUpToDown then                                                     //
    for XY := Length(MassXY)-1 downto 0 do                                        //
    begin                                                                         //
      X := MassXY[XY];                                                            //
      for Y := 0 to Length(Chip)-1 do                                             //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
//////////////////////////// * ����� ����� (������) * //////////////////////////////
                                                                                  //
  if Direct = dRDownToUp then                                                     //
    for XY := Length(MassXY)-1 downto 0 do                                        //
    begin                                                                         //
      X := MassXY[XY];                                                            //
      for Y := Length(Chip)-1 downto 0 do                                         //
        if IsChip(Chip[Y, X].Status) then                                         //
        begin                                                                     //
          Inc(N);                                                                 //
          Chip[Y, X].ID := N; // ����� ���������                                  //
          ChipN[N-1] := Point(X, Y);                                              //
        end;                                                                      //
    end;                                                                          //
                                                                                  //
////////////////////// * ������� � ������ ������ (������) * ////////////////////////
                                                                                  //
  if Direct in [dRUpSnake, dRDownSnake] then                                      //
  begin                                                                           //
    if Direct = dRUpSnake then tmp := 1                                           //
                          else tmp := 0;                                          //
    for XY := Length(MassXY)-1 downto 0 do                                        //
    begin                                                                         //
      X := MassXY[XY];                                                            //
      if (XY mod 2) = tmp then                                                    //
      begin                                                                       //
        for Y := 0 to Length(Chip)-1 do                                           //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
      end                                                                         //
      else                                                                        //
        for Y := Length(Chip)-1 downto 0 do                                       //
          if IsChip(Chip[Y, X].Status) then                                       //
          begin                                                                   //
            Inc(N);                                                               //
            Chip[Y, X].ID := N; // ����� ���������                                //
            ChipN[N-1] := Point(X, Y);                                            //
          end;                                                                    //
    end;                                                                          //
  end;                                                                            //
                                                                                  //
  SetLength(ChipN, N);                                                            //
end;                                                                              //
////////////////////////////////////////////////////////////////////////////////////


//////////////////////////////////////
function TWafer.IsWafer(): Boolean; //
begin                               //
  Result := Diameter <> 0;          //
end;                                //
//////////////////////////////////////


/////////////////////////////////////////////////////////////////////
function TWafer.GetStatusName(const Status: WORD): String;         //
var                                                                //
  n, Tmp: WORD;                                                    //
  P: byte;                                                         //
begin                                                              //
  Result := '';                                                    //
                                                                   //
  if StatusNamesSL.Count = 0 then Exit;                            //
                                                                   //
  with StatusNamesSL do                                            //
    for n := 0 to Count-1 do                                       //
      if Trim(Strings[n]) <> '' then                               //
      begin                                                        //
        P := Pos('=', Strings[n]);                                 //
        if P <> 0 then                                             //
        begin                                                      //
          try                                                      //
            Tmp := StrToInt(Trim(Copy(Strings[n], 1, P-1)));       //
          except                                                   //
            Continue;                                              //
          end;                                                     //
          if Status = Tmp then                                     //
          begin                                                    //
            Result := Copy(Strings[n], P+1, Length(Strings[n])-P); //
            Break;                                                 //
          end;                                                     //
        end;                                                       //
      end;                                                         //
end;                                                               //
/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////////////
function TWafer.LoadGammaMDB(const MDBfName, WafName: TFileName): Boolean; //
var                                                                        //
  ADOQuery: TADOQuery;                                                     //
  SL: TStringList;                                                         //
  X, Y, n, i: DWORD;                                                       //
  Str: string;                                                             //
  P: byte;                                                                 //
begin                                                                      //
  Result := False;                                                         //
                                                                           //
  FormatSettings.DecimalSeparator := ',';                                  //
                                                                           //
  fName := MDBfName;                                                       //
  TimeDate := DateToStr(FileDateToDateTime(FileAge(MDBfName)));            //
  Num := WafName;                                                          //
  MeasSystem := 'Gamma 156';                                               //
  Condition := 'NU';                                                       //
                                                                           //
  Str := ExtractFileName(MDBfName);                                        //
  Delete(Str, Pos('.', Str), Length(Str));                                 //
  P := LastDelimiter('-', Str);                                            //
  if P > 0 then                                                            //
  begin                                                                    //
    Code := Copy(Str, 1, P-1);                                             //
    NLot := Copy(Str, P+1, Length(Str));                                   //
  end;                                                                     //
                                                                           //
                                                                           //
  ADOQuery := TADOQuery.Create(nil);                                       //
  with ADOQuery do                                                         //
  begin                                                                    //
    LockType := ltReadOnly;                                                //
    ConnectionString := ConnStr+MDBfName+';';                              //
                                                                           //
    SQL.Clear;                                                             //
    SQL.Add('SELECT * FROM '+WafName+';');                                 //
    try                                                                    //
      Open;                                                                //
    except                                                                 //
      ErrMess(Handle, '������ ������� �������!');                          //
      ADOQuery.Free();                                                     //
    end;                                                                   //
                                                                           //
    SL := TStringList.Create();                                            //
    GetFieldNames(SL);                                                     //
    if SL.Count > 0 then                                                   //
    begin                                                                  //
      if Length(TestsParams) = 0 then SetLength(TestsParams, SL.Count-3);  //
      for n := 3 to SL.Count-1 do         // ��������� 3 �������           //
      begin                                                                //
        TestsParams[n-3].Name := SL[n];                                    //
        TestsParams[n-3].Norma.Min := -NotSpec;                            //
        TestsParams[n-3].Norma.Max :=  NotSpec;                            //
      end;                                                                 //
    end;                                                                   //
    SL.Free();                                                             //
                                                                           //
    with RecordSet do                                                      //
      if RecordCount > 0 then                                              //
      begin                                                                //
        NTotal := RecordCount;                                             //
                                                                           //
        X := Ceil(sqrt(NTotal)); // �������                                //
        Y := X;                  // �������� ����������                    //
                                                                           //
        SetLength(ChipN, NTotal);                                          //
        SetLength(Chip, 0, 0);                                             //
        SetLength(Chip, Y, X);                                             //
        for Y := 0 to Length(Chip)-1 do      // �������                    //
          for X := 0 to Length(Chip[0])-1 do // ������                     //
          begin                              // �����                      //
            Chip[Y, X].Status := 2;          //                            //
            Chip[Y, X].ID     := 0;          //                            //
            SetLength(Chip[Y, X].ChipParams, Length(TestsParams));         //
          end;                               //                            //
                                                                           //
        MoveFirst;                                                         //
                                                                           //
        Y := 0;                                                            //
        X := 0;                                                            //
        for n := 0 to RecordCount-1 do                                     //
        begin                                                              //
          if Fields[1].Value = '�' then Chip[Y, X].Status := 2000          //
                                   else Chip[Y, X].Status := 1;            //
                                                                           //
          Chip[Y, X].ID := n+1;      // ������� ���������                  //
          ChipN[n].X := X;                                                 //
          ChipN[n].Y := Y;                                                 //
//          ChipN[n] := Point(X, Y);                                         //
                                                                           //
          for i := 3 to Fields.Count-1 do // ��������� 3 �������           //
          begin                                                            //
            try                                                            //
              Chip[Y, X].ChipParams[i-3].Value := Single(Fields[i].Value); //
            except                                                         //
              Chip[Y, X].ChipParams[i-3].Value := NotSpec;                 //
            end;                                                           //

            if Chip[Y, X].ChipParams[i-3].Value = 0 then Chip[Y, X].ChipParams[i-3].Value := NotSpec; // ���� 0, �� �������� �� ��������� (��������)
          end;                                                             //
                                                                           //
          Inc(X);                                                          //
          if X = Length(Chip[0]) then                                      //
          begin                                                            //
            X := 0;                                                        //
            Inc(Y);                                                        //
          end;                                                             //
                                                                           //
          MoveNext;                                                        //
        end;                                                               //
      end                                                                  //
      else                                                                 //
      begin                                                                //
        ErrMess(Handle, '������ ��������!');                               //
//        if Assigned(OnEvent) then OnEvent(evError, '������ ��������!');    //
        Close;                                                             //
        ADOQuery.Free();                                                   //
        FormatSettings.DecimalSeparator := '.';                            //
      end;                                                                 //
                                                                           //
    Close;                                                                 //
  end;                                                                     //
                                                                           //
  CalcChips;                                                               //
                                                                           //
  Direct := 2; // ������� ����� �� ����� �� ���������                      //
                                                                           //
  ADOQuery.Free();                                                         //
                                                                           //
  FormatSettings.DecimalSeparator := '.';                                  //
                                                                           //
  Result := True;                                                          //
end;                                                                       //
/////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////
function TWafer.LoadSchusterTXT(const TXTfName: TFileName): Boolean;     //
var                                                                      //
  SL: TStringList;                                                       //
  P: byte;                                                               //
  n, NPStr, NCh: DWORD;                                                  //
  i: WORD;                                                               //
  X, Y, MinX, MaxX, MaxY: Integer;                                       //
  Str: string;                                                           //
  TmpChip: TChips;                                                       //
  TmpValue: Single;                                                      //
begin                                                                    //
  Result := False;                                                       //
                                                                         //
  SL := TStringList.Create;                                              //
  SL.LoadFromFile(TXTfName);                                             //
                                                                         //
  if SL.Count = 0 then Exit;                                             //
                                                                         //
  fName := TXTfName;                                                     //
  TimeDate := DateToStr(FileDateToDateTime(FileAge(TXTfName)));          //
  Num := ChangeFileExt(ExtractFileName(TXTfName), '');                   //
  MeasSystem := 'Schuster';                                              //
  Condition := '��';                                                     //
                                                                         //
  Code := '-';                                                           //
  NLot := '-';                                                           //
                                                                         //
  SetLength(TestsParams, 0);                                             //
                                                                         //
  FormatSettings.DecimalSeparator := ',';                                //
                                                                         //
  NPStr := SL.Count;                                                     //
  NCh := 0;                                                              //
  n := 0;                                                                //
  while n < SL.Count do                                                  //
  begin                                                                  //
    if Trim(SL.Strings[n]) = '' then                                     //
    begin                                                                //
      Inc(n);                                                            //
      Continue;                                                          //
    end;                                                                 //
                                                                         //
    if n < NPStr then // ���� ��������� ���������                        //
    begin                                                                //
      if Pos('MODULE', UpperCase(SL.Strings[n])) <> 0 then               //
      begin                                                              //
        P := Pos(#9, SL.Strings[n]);                                     //
        Device := Trim(Copy(SL.Strings[n], P+1, Length(SL.Strings[n]))); //
                                                                         //
        Inc(n);                                                          //
        Continue;                                                        //
      end;                                                               //
                                                                         //
//      if Pos('CONFIG', SL.Strings[n]) <> 0 then
//      begin
//        P := Pos(#9, SL.Strings[n]);
//        Code := Trim(Copy(SL.Strings[n], P+1, Length((SL.Strings[n]))));
//        Continue;
//      end;

      if Pos('STEP', UpperCase(SL.Strings[n])) <> 0 then
      begin
        i := 0;
        Str := Trim(SL.Strings[n+i+1]);
        while Str <> '' do
        begin
          SetLength(TestsParams, i+1);

          Delete(Str, 1, Pos(#9, Str)); // ������ ����� �����
          P := Pos(#9, Str);
          TestsParams[i].Name := Trim(Copy(Str, 1, P));
          Delete(Str, 1, P); // ������ ��� �����
          P := Pos(#9, Str);
          TestsParams[i].PMode := Trim(Copy(Str, 1, P));
          Delete(Str, 1, P); // ������ ����� �����
          P := Pos(' ', Str);
          TestsParams[i].Norma.Min := StrToFloat(Trim(Copy(Str, 1, P)));
          Delete(Str, 1, P); // ������ ���. �����
          Delete(Str, 1, Pos(' ', Str));
          P := Pos(' ', Str);
          TestsParams[i].Norma.Max := StrToFloat(Trim(Copy(Str, 1, P)));
          Delete(Str, 1, P); // ������ ����. �����
          TestsParams[i].PUnit := Trim(Str);
          TestsParams[i].Name := TestsParams[i].Name+' ('+TestsParams[i].PUnit+')';

          Inc(i);
          Str := Trim(SL.Strings[n+i+1]);
        end;

        Inc(n, i);
        Continue;
      end;

      if Pos('FILE', UpperCase(SL.Strings[n])) <> 0 then
      begin
        Str := SL.Strings[n+1]; // 1-� ���
        Delete(Str, 1, Pos(#9, Str)); // FILE
        Delete(Str, 1, Pos(#9, Str)); // NAME
        Delete(Str, 1, Pos(#9, Str)); // ADAPTER
        if Pos('-',  Str) = 1 then Diameter := 0    // ����������
                              else Diameter := 150; // ��������
        Delete(Str, 1, Pos(#9, Str)); // SERIAL N0#
        Delete(Str, 1, Pos(#9, Str)); // BIN#
        Delete(Str, 1, Pos(#9, Str)); // SERIAL#
        P := Pos(#9, Str);
        TimeDate := Copy(Str, 1, P-1);

        NPStr := n; // ������ ���� ���-�� ��������� �����
        NTotal := SL.Count-n-1; // ���-�� �����

        SetLength(ChipN, NTotal);
        SetLength(Chip, 0, 0);

        if Diameter = 0 then // ���� ����������
        begin
          X := Ceil(sqrt(NTotal)); // �������
          Y := X;                  // �������� ���������� ��� ���������
          SetLength(Chip, Y, X);

          X := 0; // ������� ���
          Y := 0; // ���������
        end;

        Inc(n);
        Continue;
      end;
    end

    else             // ���� ��������� ���-�� ��������� �����

////////////////////////////////////////////////////////////////////////////////

    begin
      Str := SL.Strings[n];         // ������
      Delete(Str, 1, Pos(#9, Str)); // FILE
      Delete(Str, 1, Pos(#9, Str)); // NAME
      Delete(Str, 1, Pos(#9, Str)); // ADAPTER

      if Diameter <> 0 then // ���� �������� - ������� ����������
      begin
        if NCh = 0 then // 1-� ��������
        begin
          P := Pos('ID', UpperCase(Str)); // ������ '*ID='
          Delete(Str, 1, Pos('=', Str));  // ��� ����������
        end;

        P := Pos(',', Str);
        X := StrToInt(Copy(Str, 1, P-1));
        Delete(Str, 1, P);
        P := Pos(#9, Str);
        Y := StrToInt(Copy(Str, 1, P-1));
        Delete(Str, 1, P);

        ChipN[NCh].X := X;
        ChipN[NCh].Y := Y;

        if NCh = 0 then // �����
        begin
          MinX := X; // ����������� X
          MaxX := X; // ������������ X
          MaxY := Y; // ������������ Y
        end
        else
        begin
          if X < MinX then MinX := X;
          if X > MaxX then MaxX := X;
          if Y > MaxY then MaxY := Y;
        end;

      end
      else                 // ���� ����������
      begin
        Delete(Str, 1, 1); // ������ '-' ��� ����������

        ChipN[NCh].X := X;
        ChipN[NCh].Y := Y;

        Inc(X);         ;
        if X = Length(Chip[0]) then   ////////////// �� ������  SetLength(Chip, Y, X)
        begin
          X := 0;
          Inc(Y);
        end;
      end;

      Inc(NCh);
    end;

    Inc(n);
  end;

  if Diameter <> 0 then // ���� ��������
  begin
    X := (MaxX-MinX)+1;
    Y := MaxY+1;
    SetLength(Chip, Y, X);
  end;

  for Y := 0 to Length(Chip)-1 do      // �������
    for X := 0 to Length(Chip[0])-1 do // ������
    begin                              // �����
      Chip[Y, X].Status := 2;          //
      Chip[Y, X].ID     := 0;          //
      SetLength(Chip[Y, X].ChipParams, Length(TestsParams));
    end;                               //

////////////////////////////////////////////////////////////////////////////////

  for n := NPStr+1 to SL.Count-1 do // ��� ��� �������� �� ���-��� ���������
  begin
    NCh := n-(NPStr+1);

    if Diameter <> 0 then ChipN[NCh].X := ChipN[NCh].X-MinX; // ����������� X
    Y := ChipN[NCh].Y;
    X := ChipN[NCh].X;

    Chip[Y, X].ID := NCh+1;

    Str := SL.Strings[n];         // ������
    Delete(Str, 1, Pos(#9, Str)); // FILE
    Delete(Str, 1, Pos(#9, Str)); // NAME
    Delete(Str, 1, Pos(#9, Str)); // ADAPTER
    P := Pos(#9, Str);
//    if Trim(Copy(Str, 1, P-1)) = 'FAIL' then

    Delete(Str, 1, Pos(#9, Str)); // SERIAL N0#
    Delete(Str, 1, Pos(#9, Str)); // BIN#
    Delete(Str, 1, Pos(#9, Str)); // SERIAL#
    Delete(Str, 1, Pos(#9, Str)); // DATE
    Delete(Str, 1, Pos(#9, Str)); // TIME
    P := Pos(#9, Str);
    if Trim(Copy(Str, 1, P-1)) = 'FAIL' then Chip[Y, X].Status := 10 // ����� ����� �� ��
                                        else Chip[Y, X].Status := 1;
    Delete(Str, 1, Pos(#9, Str)); // PASS/FAIL


    if Length(TestsParams) > 0 then
      for i := 0 to Length(TestsParams)-1 do
      begin
        P := Pos(#9, Str);

        TmpValue := StrToFloat(Trim(Copy(Str, 1, P-1)));
        if TmpValue = -999 then TmpValue := NotSpec;

        Chip[Y, X].ChipParams[i].Value := TmpValue;
        Chip[Y, X].ChipParams[i].Stat  := GetChipParamsStat(Chip[Y, X].ChipParams[i].Value,
                                                            TestsParams[i].Norma.Min,
                                                            TestsParams[i].Norma.Max);
        if Chip[Y, X].Status < 2000 then
          if Chip[Y, X].ChipParams[i].Stat <> 1 then Chip[Y, X].Status := 2000+i
                                                else Chip[Y, X].Status := 1;

        Delete(Str, 1, Pos(#9, Str));
      end;
  end;

/////////////// ������� ����� �� Y (������� ��������� ������)///////////////////

  if Diameter <> 0 then // ���� ��������
  begin
    SetLength(TmpChip, Length(Chip[0]), Length(Chip));
    for Y := 0 to Length(Chip)-1 do
      for X := 0 to Length(Chip[0])-1 do
      begin
        TmpChip[Length(Chip)-Y-1, X] := Chip[Y, X];
        SetLength(TmpChip[Length(Chip)-Y-1, X].ChipParams, Length(Chip[Y, X].ChipParams));
      end;
    Chip := TmpChip;
    SetLength(Chip, Length(TmpChip[0]), Length(TmpChip));
    TmpChip := nil;

    Direct  := 9; // ����� �� Schuster� - ������ ����� ������
    CutSide := 3; // ���� �����
  end
  else                  // ���� ����������
  begin
    Direct  := 2;
    CutSide := 1;
  end;

  SetChipsID();

////////////////////////////////////////////////////////////////////////////////

  CalcChips();

  BaseChip := ChipN[0]; // 1-� ��� - �������

  SL.Free();

  FormatSettings.DecimalSeparator := '.';

  Result := True;
end;
//////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadBlankTXT(const TXTfName: TFileName): Boolean;                 //
                                                                                  //
  /////////////////////////////////////////////////////////                       //
  function GetCoordFromStr(S: string; nT: WORD): TPoint; //                       //
  var                                                    //                       //
    P: WORD;                                             //                       //
  begin                                                  //                       //
    Result := Point(-999, -999);                         //                       //
                                                         //                       //
    while nT > 0 do                                      //                       //
    begin                                                //                       //
      P := Pos(#9, S);                                   //                       //
      Delete(S, 1, P);                                   //                       //
                                                         //                       //
      Dec(nT);                                           //                       //
    end;                                                 //                       //
                                                         //                       //
    P := Pos('*ID=', S);                                 //                       //
    if P <> 0 then Delete(S, 1, P+3);                    //                       //
    P := Pos(#9, S);                                     //                       //
    if P <> 0 then Delete(S, P, Length(S));              //                       //
                                                         //                       //
    S := Trim(S);                                        //                       //
    P := Pos(',', S);                                    //                       //
    try                                                  //                       //
      Result.X := StrToInt(Copy(S, 1, P-1));             //                       //
    except                                               //                       //
                                                         //                       //
    end;                                                 //                       //
    Delete(S, 1, P);                                     //                       //
    S := Trim(S);                                        //                       //
    try                                                  //                       //
      Result.Y := StrToInt(Copy(S, 1, Length(S)));       //                       //
    except                                               //                       //
                                                         //                       //
    end;                                                 //                       //
  end;                                                   //                       //
  /////////////////////////////////////////////////////////                       //
                                                                                  //
var                                                                               //
  n, NCh: DWORD;                                                                  //
  X, Y, MinX, MaxX, MaxY: Integer;                                                //
  SL: TStringList;                                                                //
  Str: string;                                                                    //
  Read_Coord: Boolean;                                                            //
  i, P, NTab: WORD;                                                               //
begin                                                                             //
  Result := False;                                                                //
                                                                                  //
  SL := TStringList.Create;                                                       //
  SL.LoadFromFile(TXTfName);                                                      //
                                                                                  //
  if SL.Count = 0 then Exit;                                                      //
                                                                                  //
  fName := TXTfName;                                                              //
  MeasSystem := 'Schuster';                                                       //
  Condition := '��';                                                              //
  Diameter := 150;                                                                //
  LDiameter := 144.25;                                                            //
  Radius  := Diameter/2;                                                          //
  LRadius := Radius-(Diameter-LDiameter);                                         //
  Chord   := Sqrt(Radius*Radius-LRadius*LRadius);                                 //
  Direct  := 9; // ����� �� Schuster� - ������ ����� ������                       //
  CutSide := 3; // ���� �����                                                     //
                                                                                  //
  SetLength(ChipN, SL.Count);                                                     //
  SetLength(Chip, 0, 0);                                                          //
                                                                                  //
  FormatSettings.DecimalSeparator := ',';                                         //
                                                                                  //
  Read_Coord := False;                                                            //
  NTab := 0;                                                                      //
  NCh := 0;                                                                       //
  for n := 0 to SL.Count-1 do                                                     //
  begin                                                                           //
    Str := Trim(SL.Strings[n]);                                                   //
                                                                                  //
    if Str = '' then Continue;                                                    //
                                                                                  //
    if Read_Coord then                                                            //
    begin                                                                         //
      ChipN[NCh] := GetCoordFromStr(Str, NTab);                                   //
                                                                                  //
      if ChipN[NCh].X < MinX then MinX := ChipN[NCh].X;                           //
      if ChipN[NCh].X > MaxX then MaxX := ChipN[NCh].X;                           //
      if ChipN[NCh].Y > MaxY then MaxY := ChipN[NCh].Y;                           //
                                                                                  //
      Inc(NCh);                                                                   //
    end                                                                           //
    else                                                                          //
    begin                                                                         //
      P := Pos('*ID=', Str);                                                      //
      if P <> 0 then                                                              //
      begin                                                                       //
        for i := 1 to P do                                                        //
          if Str[i] = #9 then Inc(NTab); // ����� ���-�� Tab�� �� ���������      //
                                                                                  //
        ChipN[NCh] := GetCoordFromStr(Str, NTab);                                 //
                                                                                  //
        MinX := ChipN[NCh].X; // ����������� X                                    //
        MaxX := ChipN[NCh].X; // ������������ X                                   //
        MaxY := ChipN[NCh].Y; // ������������ Y                                   //
                                                                                  //
        Inc(NCh);                                                                 //
                                                                                  //
        Read_Coord := True;                                                       //
      end;                                                                        //
    end;                                                                          //
  end;                                                                            //
                                                                                  //
  SetLength(ChipN, NCh);                                                          //
  if NCh = 0 then ;                                                               //
                                                                                  //
                                                                                  //
  X := (MaxX-MinX)+1;                                                             //
  Y := MaxY+1;                                                                    //
  SetLength(Chip, Y, X);                                                          //
                                                                                  //
  for Y := 0 to Length(Chip)-1 do      // �������                                 //
    for X := 0 to Length(Chip[0])-1 do // ������                                  //
    begin                              // �����                                   //
      Chip[Y, X].Status := 2;          //                                         //
      Chip[Y, X].ID     := 0;          //                                         //
    end;                               //                                         //
                                                                                  //
  for NCh := 0 to Length(ChipN)-1 do                                              //
  begin                                                                           //
    ChipN[NCh].X := ChipN[NCh].X-MinX; // ����������� X (����� �������������)    //
                                                                                  //
    Y := Length(Chip)-ChipN[NCh].Y-1; // ������� �� Y (������� ��������� ������)  //
    X := ChipN[NCh].X;                                                            //
                                                                                  //
    Chip[Y, X].ID := NCh+1;                                                       //
    Chip[Y, X].Status := 0;                                                       //
  end;                                                                            //
                                                                                  //
  SL.Free();                                                                      //
                                                                                  //
  FormatSettings.DecimalSeparator := '.';                                         //
                                                                                  //
  SetChipsID();                                                                   //
  CalcChips();                                                                    //
                                                                                  //
  BaseChip := ChipN[0]; // 1-� ��� - �������                                      //
                                                                                  //
  Result := True;                                                                 //
end;                                                                              //
////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////
function TWafer.LoadBlank690(const TXTfName: TFileName): Boolean;  //
var                                                                //
  IniFile: TIniFile;                                               //
  Zone: TZones2;                                                   //
  P: byte;                                                         //
  X, Y, n: WORD;                                                   //
  MinX, MinY, MaxX, MaxY: Integer;                                 //
  SL: TStringList;                                                 //
  Str: String;                                                     //
begin                                                              //
  Result := False;                                                 //
                                                                   //
  fName := TXTfName;                                               //
                                                                   //
  FormatSettings.DecimalSeparator := ',';                          //
                                                                   //
  SL := TStringList.Create();                                      //
  IniFile := TIniFile.Create(fName);                               //
  try                                                              //
    Code       := IniFile.ReadString('690', 'Code',      '');      //
//    Device     := IniFile.ReadString('690', 'Device',     0);      //
    Diameter   := IniFile.ReadInteger('690', 'Diametr',   0);      //
    StepX      := IniFile.ReadInteger('690', 'ChipSizeX', 0)/1000; //
    StepY      := IniFile.ReadInteger('690', 'ChipSizeY', 0)/1000; //
    BaseChip.X := IniFile.ReadInteger('690', 'BaseChipX', 0);      //
    BaseChip.Y := IniFile.ReadInteger('690', 'BaseChipY', 0);      //
    Direct     := IniFile.ReadInteger('690', 'Path', 0);           //
    CutSide    := IniFile.ReadInteger('690', 'Notch', 0);          //
                                                                   //
    MinX := BaseChip.X;                                            //
    MinY := BaseChip.Y;                                            //
    MaxX := BaseChip.X;                                            //
    MaxY := BaseChip.Y;                                            //
                                                                   //
    IniFile.ReadSectionValues('Zones', SL);                        //
    SetLength(Zone, SL.Count);                                     //
    if Length(Zone) > 0 then                                       //
      for n := 0 to Length(Zone)-1 do                              //
      begin                                                        //
        Str := Trim(SL.Strings[n]);                                //
                                                                   //
        P := Pos('=', Str);                                        //
        Delete(Str, 1, P);                                         //
        P := Pos(';', Str);                                        //
        Zone[n].Y1 := StrToInt(Copy(Str, 1, P-1));                 //
        if MinY > Zone[n].Y1 then MinY := Zone[n].Y1;              //
                                                                   //
        Delete(Str, 1, P);                                         //
        P := Pos(';', Str);                                        //
        Zone[n].Y2 := StrToInt(Copy(Str, 1, P-1));                 //
        if MaxY < Zone[n].Y2 then MaxY := Zone[n].Y2;              //
                                                                   //
        Delete(Str, 1, P);                                         //
        P := Pos(';', Str);                                        //
        Zone[n].X3 := StrToInt(Copy(Str, 1, P-1));                 //
        if MinX > Zone[n].X3 then MinX := Zone[n].X3;              //
                                                                   //
        Delete(Str, 1, P);                                         //
        Zone[n].X4 := StrToInt(Str);                               //
        if MaxX < Zone[n].X4 then MaxX := Zone[n].X4;              //
      end;                                                         //
  except                                                           //
    IniFile.Free();                                                //
  end;                                                             //
  SL.Free();                                                       //
                                                                   //
/////////////////////////////////////////////////                  //
                                                                   //
  if Code = '' then                                                //
  begin                                                            //
    ErrMess(Handle, '������ ������ �����!');                       //
    Exit;                                                          //
  end;                                                             //
  if Length(Zone) = 0 then                                         //
  begin                                                            //
    ErrMess(Handle, '������ ������ ��� ������!');                  //
    Exit;                                                          //
  end;                                                             //
                                                                   //
//  SetLength(ChipN, SL.Count);                                      //
  SetLength(Chip, 0, 0);                                           //
                                                                   //
  MaxY := MaxY-MinY+1;                                             //
  MaxX := MaxX-MinX+1;                                             //
  SetLength(Chip, MaxY, MaxX);                                     //
  for Y := 0 to MaxY-1 do                                          //
    for X := 0 to MaxX-1 do                                        //
    begin                                                          //
      Chip[Y, X].Status := 2;                                      //
      Chip[Y, X].ID     := 0;                                      //
    end;                                                           //
                                                                   //
  NTotal := 0;                                                     //
  for n := 0 to Length(Zone)-1 do                                  //
    for Y := Zone[n].Y1 to Zone[n].Y2 do                           //
      for X := Zone[n].X3 to Zone[n].X4 do                         //
      begin                                                        //
        Chip[Y-MinY, X-MinX].Status := 0;                          //
                                                                   //
        Inc(NTotal);                                               //
      end;                                                         //
                                                                   //
  BaseChip.X := BaseChip.X-MinX;                                   //
  BaseChip.Y := BaseChip.Y-MinY;                                   //
                                                                   //
  if Diameter = 150 then LDiameter := 144.25                       //
                    else LDiameter := Diameter*0.96;               //
  Radius  := Diameter/2;                                           //
  LRadius := Radius-(Diameter-LDiameter);                          //
  Chord   := Sqrt(Radius*Radius-LRadius*LRadius);                  //
                                                                   //
  FormatSettings.DecimalSeparator := '.';                          //
                                                                   //
  SetChipsID();                                                    //
//  CalcChips();                                                     //
                                                                   //
  Result := True;                                                  //
end;                                                               //
/////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadSTS(const STSfName: TFileName): Boolean;                                         //
var                                                                                                  //
  i, X, Y: WORD;                                                                                     //
  n, Count: DWORD;                                                                                   //
                                                                                                     //
  SL: TStringList;                                                                                   //
  Str, S: String;                                                                                    //
  P: byte;                                                                                           //
  Mass: array of Single;                                                                             //
  Stat: WORD;                                                                                        //
begin                                                                                                //
  Result := False;                                                                                   //
                                                                                                     //
  fName := STSfName;                                                                                 //
                                                                                                     //
  if not LoadSTSHeader() then                                                                        //
  begin                                                                                              //
    ErrMess(Handle, '������ �������� ���������!');                                                   //
//    Init;                                                                                            //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  SL := TStringList.Create;                                                                          //
  SL.LoadFromFile(STSfName);                                                                         //
                                                                                                     //
  Count := 0;                                                                                        //
  n := 0;                                                                                            //
  while (Trim(SL.Strings[0]) <> '[ChipsParams]') do                                                  //
  begin                                                                                              //
    if SL.Count = 1 then                                                                             //
    begin                                                                                            //
      ErrMess(Handle, '�� ������� ���� [ChipsParams]!');                                             //
//      Init;                                                                                          //
      SL.Free;                                                                                       //
      Exit;                                                                                          //
    end;                                                                                             //
    SL.Delete(0);                                                                                    //
    Inc(Count);                                                                                      //
  end;                                                                                               //
  SL.Delete(0);         //                                                                           //
  Inc(Count);           //                                                                           //
  Str := SL.Strings[0]; // ������ ���� [ChipsParams]                                                 //
  SL.Delete(0);         //                                                                           //
  Inc(Count);           //                                                                           //
                                                                                                     //
  if SL.Count = 0 then                                                                               //
  begin                                                                                              //
    ErrMess(Handle, '����� ������!');                                                                //
//    Init;                                                                                            //
    SL.Free;                                                                                         //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  try                                                                                                //
    P := Pos(#9, Str);                                                                               //
    S := Copy(Str, 1, P-1);                                                                          //
    n := 0;                                                                                          //
    while Trim(S) <> 'Status' do                        //                                           //
    begin                                               //                                           //
      TestsParams[n].Name := S;                         //                                           //
      Inc(n);                                           // ���������                                 //
      Delete(Str, 1, P);                                // ��������                                  //
      P := Pos(#9, Str);                                // ��������,                                 //
      S := Copy(Str, 1, P-1);                           //                                           //
      if S = '' then                                    //                                           //
      begin                                             //                                           //
        ErrMess(Handle, '�� ������ ������� ������� !'); //                                           //
        SL.Free;                                        //                                           //
        Exit;                                           //                                           //
      end;                                              //                                           //
    end;                                                //                                           //
    SetLength(Mass, Length(TestsParams));                                                            //
                                                                                                     //
    FormatSettings.DecimalSeparator := ',';                                                          //
                                                                                                     //
    if SL.Count > 0 then                                                                             //
      for n := 0 to SL.Count-1 do                                                                    //
      begin                                                                                          //
        Str := SL.Strings[n];                                                                        //
                                                                                                     //
        if Trim(Str) = '' then Continue;                                                             //
                                                                                                     //
        if Length(TestsParams) > 0 then                                                              //
          for i := 0 to Length(TestsParams)-1 do                                                     //
          begin                                                                                      //
            P := Pos(#9, Str);                                                                       //
            S := Copy(Str, 1, P-1);                                                                  //
            Mass[i] := StrToFloat(S);                                                                //
            Delete(Str, 1, P);                                                                       //
          end;                                                                                       //
                                                                                                     //
        P := Pos(#9, Str);                                                                           //
        S := Copy(Str, 1, P-1);                                                                      //
        Stat := StrToInt(S);                                                                         //
        Delete(Str, 1, P);                                                                           //
                                                                                                     //
        P := Pos(#9, Str);                                                                           //
        S := Copy(Str, 1, P-1);                                                                      //
        X := StrToInt(S);                                                                            //
        Delete(Str, 1, P);                                                                           //
        if X > (Length(Chip[0])-1) then SetLength(Chip[0], X+1);                                     //
                                                                                                     //
        Y := StrToInt(Str);                                                                          //
        Delete(Str, 1, P);                                                                           //
        if Y > (Length(Chip)-1) then SetLength(Chip, Y+1);                                           //
                                                                                                     //
        Chip[Y, X].Status := Stat;                                                                   //
                                                                                                     //
        SetLength(Chip[Y, X].ChipParams, Length(TestsParams));                                       //
        if Length(TestsParams) > 0 then                                                              //
          for i := 0 to Length(TestsParams)-1 do                                                     //
          begin                                                                                      //
            Chip[Y, X].ChipParams[i].Value := Mass[i];                                               //
            Chip[Y, X].ChipParams[i].Stat  := GetChipParamsStat(Mass[i],                             //
                                                     TestsParams[i].Norma.Min,                       //
                                                     TestsParams[i].Norma.Max);                      //
          end;                                                                                       //
                                                                                                     //
        if Trim(SL.Strings[n]) = '' then Break;                                                      //
      end;                                                                                           //
                                                                                                     //
      FormatSettings.DecimalSeparator := '.';                                                        //
                                                                                                     //
  except                                                                                             //
    ErrMess(Handle, '������ � ������ '+IntToStr(n+Count+1));                                         //
//    Init;                                                                                            //
    SL.Free;                                                                                         //
    FormatSettings.DecimalSeparator := '.';                                                          //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  Result := True;                                                                                    //
                                                                                                     //
  SetChipsID;                                                                                        //
  CalcChips;                                                                                         //
                                                                                                     //
  SL.Free;                                                                                           //
end;                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.AddSTS(const STSfName: TFileName): Boolean;                                          //
var                                                                                                  //
  i, X, Y: WORD;                                                                                     //
  n, Count: DWORD;                                                                                   //
  SL: TStringList;                                                                                   //
  Str, S: String;                                                                                    //
  P: byte;                                                                                           //
  Mass: array of Single;                                                                             //
  Stat: WORD;                                                                                        //
  tmpWafer: TWafer;                                                                                  //
begin                                                                                                //
  Result := False;                                                                                   //
                                                                                                     //
  tmpWafer := TWafer.Create(Handle);                                                                 //
  tmpWafer.fName := STSfName;                                                                        //
                                                                                                     //
  FormatSettings.DecimalSeparator := ',';                                                            //
                                                                                                     //
  if not tmpWafer.LoadSTSHeader() then                                                               //
  begin                                                                                              //
    ErrMess(Handle, '������ �������� ���������!');                                                   //
    tmpWafer.Free;                                                                                   //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  if tmpWafer.Code <> Code then                                                                      //
  begin                                                                                              //
    ErrMess(Handle, '����������� ����!');                                                            //
    tmpWafer.Free;                                                                                   //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  if tmpWafer.NLot <> NLot then                                                                      //
  begin                                                                                              //
    ErrMess(Handle, '����������� ������!');                                                          //
    tmpWafer.Free;                                                                                   //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
//  if tmpWafer.Num <> Wafer.Num then                                                                  //
//  begin                                                                                              //
//    ErrMess(Handle, '����������� ������ �������!');                                                  //
//    tmpWafer.Free;                                                                                   //
//    Exit;                                                                                            //
//  end;                                                                                               //
                                                                                                     //
  SL := TStringList.Create;                                                                          //
  SL.LoadFromFile(STSfName);                                                                         //
                                                                                                     //
  Count := 0;                                                                                        //
  n := 0;                                                                                            //
  while (Trim(SL.Strings[0]) <> '[ChipsParams]') do                                                  //
  begin                                                                                              //
    if SL.Count = 1 then                                                                             //
    begin                                                                                            //
      ErrMess(Handle, '�� ������� ���� [ChipsParams]!');                                             //
      tmpWafer.Free;                                                                                 //
      SL.Free;                                                                                       //
      Exit;                                                                                          //
    end;                                                                                             //
    SL.Delete(0);                                                                                    //
    Inc(Count);                                                                                      //
  end;                                                                                               //
  SL.Delete(0);         //                                                                           //
  Inc(Count);           //                                                                           //
  Str := SL.Strings[0]; // ������ ���� [ChipsParams]                                                 //
  SL.Delete(0);         //                                                                           //
  Inc(Count);           //                                                                           //
                                                                                                     //
  try                                                                                                //
    P := Pos(#9, Str);                                                                               //
    S := Copy(Str, 1, P-1);                                                                          //
    n := 0;                                                                                          //
    while Trim(S) <> 'Status' do                                                                     //
    begin                                                                                            //
      tmpWafer.TestsParams[n].Name := S;                //                                           //
      Inc(n);                                           // ���������                                 //
      Delete(Str, 1, P);                                // ��������                                  //
      P := Pos(#9, Str);                                // ��������                                  //
      S := Copy(Str, 1, P-1);                           //                                           //
      if S = '' then                                    //                                           //
      begin                                             //                                           //
        ErrMess(Handle, '�� ������ ������� ������� !'); //                                           //
        tmpWafer.Free;                                  //                                           //
        SL.Free;                                        //                                           //
        Exit;                                           //                                           //
      end;                                              //                                           //
    end;                                                //                                           //
    SetLength(Mass, Length(tmpWafer.TestsParams));                                                   //
                                                                                                     //
    if SL.Count > 0 then                                                                             //
      for n := 0 to SL.Count-1 do                                                                    //
      begin                                                                                          //
        Str := SL.Strings[n];                                                                        //
        if Trim(Str) = '' then Continue;                                                             //
        if Length(tmpWafer.TestsParams) > 0 then // ��������� ������ ������                          //
          for i := 0 to Length(tmpWafer.TestsParams)-1 do                                            //
          begin                                                                                      //
            P := Pos(#9, Str);                                                                       //
            S := Copy(Str, 1, P-1);                                                                  //
            Mass[i] := StrToFloat(S);                                                                //
            Delete(Str, 1, P);                                                                       //
          end;                                                                                       //
                                                                                                     //
        P := Pos(#9, Str);                                                                           //
        S := Copy(Str, 1, P-1);                                                                      //
        Stat := StrToInt(S);                                                                         //
        Delete(Str, 1, P);                                                                           //
                                                                                                     //
        P := Pos(#9, Str);                                                                           //
        S := Copy(Str, 1, P-1);                                                                      //
        X := StrToInt(S);                                                                            //
        Delete(Str, 1, P);                                                                           //
        if X > (Length(tmpWafer.Chip[0])-1) then SetLength(tmpWafer.Chip[0], X+1);                   //
                                                                                                     //
        Y := StrToInt(Str);                                                                          //
        Delete(Str, 1, P);                                                                           //
        if Y > (Length(tmpWafer.Chip)-1) then SetLength(tmpWafer.Chip, Y+1);                         //
                                                                                                     //
        tmpWafer.Chip[Y, X].Status := Stat;                                                          //
        if (tmpWafer.Chip[Y, X].Status < 10) or (tmpWafer.Chip[Y, X].Status > 1500) then             //
        begin                                                                                        //
          SetLength(tmpWafer.Chip[Y, X].ChipParams, Length(tmpWafer.TestsParams));                   //
          if Length(tmpWafer.TestsParams) > 0 then                                                   //
            for i := 0 to Length(tmpWafer.TestsParams)-1 do                                          //
            begin                                                                                    //
              tmpWafer.Chip[Y, X].ChipParams[i].Value := Mass[i];
              tmpWafer.Chip[Y, X].ChipParams[i].Stat  := GetChipParamsStat(Mass[i], tmpWafer.TestsParams[i].Norma.Min, tmpWafer.TestsParams[i].Norma.Max);
            end;
        end;                                                                                         //
                                                                                                     //
        if Trim(SL.Strings[n]) = '' then Break;                                                      //
      end;                                                                                           //
                                                                                                     //
      FormatSettings.DecimalSeparator := '.';                                                        //
  except                                                                                             //
    ErrMess(Handle, '������ � ������ '+IntToStr(n+Count+1));                                         //
    tmpWafer.Free;                                                                                   //
    SL.Free;                                                                                         //
    FormatSettings.DecimalSeparator := '.';                                                          //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  if CutSide <> 0 then                                     // ��������                               //
     while tmpWafer.CutSide <> CutSide do tmpWafer.Rotate; // ���� ��������                          //
  if (Length(tmpWafer.Chip[0]) <> Length(Chip[0])) and                                               //
     (Length(tmpWafer.Chip)    <> Length(Chip))    then                                              //
  begin                                                                                              //
    ErrMess(Handle, '����������� ����������� �������!     '+IntToStr(Length(tmpWafer.Chip))+' ..... '+IntToStr(Length(Chip)));
//    tmpWafer.Free;                                                                                   //
//    Exit;                                                                                            //
  end;                                                                                               //
{
  if Length(tmpWafer.TestsParams) <> Length(TestsParams) then                                        //
  begin                                                                                              //
    ErrMess(Handle, '����������� ���������� ����������!');                                           //
//    tmpWafer.Free;                                                                                   //
//    Exit;                                                                                            //
  end;                                                                                               //
}                                                                                                     //
  for Y := 0 to Length(tmpWafer.Chip)-1 do                                                           //
    for X := 0 to Length(tmpWafer.Chip[0])-1 do                                                      //
      case tmpWafer.Chip[Y,X].Status of                                                              //
        1,                                                                                           //
        10..1500,                                                                                    //
        2000..3000,                                                                                  //
        3500..4500: begin                                                                            //
                      Chip[Y,X].Status := tmpWafer.Chip[Y,X].Status;                                 //
                        if Length(tmpWafer.Chip[Y,X].ChipParams) > 0 then                            //
                        begin                                                                        //
                          if Length(Chip[Y,X].ChipParams) <> Length(tmpWafer.Chip[Y,X].ChipParams) then //
                              SetLength(Chip[Y,X].ChipParams, Length(tmpWafer.Chip[Y,X].ChipParams));   //
                                                                                                     //
                          for n := 0 to Length(tmpWafer.Chip[Y,X].ChipParams)-1 do                   //
                            Chip[Y,X].ChipParams[n].Value := tmpWafer.Chip[Y,X].ChipParams[n].Value; //
                        end;                                                                         //
                    end;                                                                             //
      end;                                                                                           //
  tmpWafer.Free;                                                                                     //
                                                                                                     //
  Result := True;                                                                                    //
                                                                                                     //
  SetChipsID;                                                                                        //
  CalcChips;                                                                                         //
                                                                                                     //
  SL.Free;                                                                                           //
end;                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadBlankSTS(const STSfName: TFileName): Boolean;                                    //
var                                                                                                  //
  X, Y: WORD;                                                                                        //
  n, Count, nParams, i: DWORD;                                                                       //
  SL: TStringList;                                                                                   //
  Str, S: String;                                                                                    //
  P: byte;                                                                                           //
  Stat: WORD;                                                                                        //
begin                                                                                                //
  Result := False;                                                                                   //
                                                                                                     //
  fName := STSfName;                                                                                 //
                                                                                                     //
  if not LoadBlankSTSHeader() then                                                                   //
  begin                                                                                              //
    ErrMess(Handle, '������ �������� ���������!');                                                   //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  SL := TStringList.Create;                                                                          //
  SL.LoadFromFile(STSfName);                                                                         //
                                                                                                     //
  Count := 0;                                                                                        //
  while (Trim(SL.Strings[0]) <> '[ChipsParams]') do                                                  //
  begin                                                                                              //
    if SL.Count = 1 then                                                                             //
    begin                                                                                            //
      ErrMess(Handle, '�� ������� ���� [ChipsParams]!');                                             //
      SL.Free;                                                                                       //
      Exit;                                                                                          //
    end;                                                                                             //
    SL.Delete(0);                                                                                    //
    Inc(Count);                                                                                      //
  end;                                                                                               //
  SL.Delete(0);         //                                                                           //
  Inc(Count);           //                                                                           //
  Str := SL.Strings[0]; // ������ ���� [ChipsParams]                                                 //
  SL.Delete(0);         //                                                                           //
  Inc(Count);           //                                                                           //
                                                                                                     //
  try                                                                                                //
    P := Pos(#9, Str);                                                                               //
    S := Copy(Str, 1, P-1);                                                                          //
    nParams := 0;                                                                                    //
    while Trim(S) <> 'Status' do                                                                     //
    begin                                                                                            //
      Inc(nParams); // ���-�� ���������� (��� ��� ���������� �������� > 0)                           //
                                                                                                     //
      Delete(Str, 1, P);                                //                                           //
      P := Pos(#9, Str);                                //                                           //
      S := Copy(Str, 1, P-1);                           //                                           //
      if S = '' then                                    // ���������                                 //
      begin                                             // ��������                                  //
        ErrMess(Handle, '�� ������ ������� ������� !'); // ��������                                  //
        SL.Free;                                        //                                           //
        Exit;                                           //                                           //
      end;                                              //                                           //
    end;                                                //                                           //
                                                                                                     //
    if SL.Count > 0 then                                                                             //
      for n := 0 to SL.Count-1 do                                                                    //
      begin                                                                                          //
        Str := SL.Strings[n];                                                                        //
        if Trim(Str) = '' then Continue;                                                             //
                                                                                                     //
        if nParams > 0 then                                                                          //
          for i := 0 to nParams-1 do // ���������                                                    //
          begin                      // ���������                                                    //
            P := Pos(#9, Str);       // ���                                                          //
            S := Copy(Str, 1, P-1);  // ���                                                          //
            Delete(Str, 1, P);       // ����������                                                   //
          end;                       // ��������                                                     //
                                                                                                     //
        P := Pos(#9, Str);                                                                           //
        S := Copy(Str, 1, P-1);                                                                      //
        Stat := StrToInt(S); // Status                                                               //
        Delete(Str, 1, P);                                                                           //
                                                                                                     //
        P := Pos(#9, Str);                                                                           //
        S := Copy(Str, 1, P-1);                                                                      //
        X := StrToInt(S);    // X                                                                    //
        Delete(Str, 1, P);                                                                           //
        if X > (Length(Chip[0])-1) then SetLength(Chip[0], X+1);                                     //
                                                                                                     //
        Y := StrToInt(Str);  // Y                                                                    //
        Delete(Str, 1, P);                                                                           //
        if Y > (Length(Chip)-1) then SetLength(Chip, Y+1);                                           //
                                                                                                     //
        if Stat in [0, 2, 3, 4, 5, 7] then Chip[Y, X].Status := Stat                                 //
                                      else Chip[Y, X].Status := 0; // ��� ��� ���������� �������     //
                                                                                                     //
        if Trim(SL.Strings[n]) = '' then Break;                                                      //
      end;                                                                                           //
                                                                                                     //
      FormatSettings.DecimalSeparator := '.';                                                        //
  except                                                                                             //
    ErrMess(Handle, '������ � ������ '+IntToStr(n+Count+1));                                         //
    SL.Free;                                                                                         //
    FormatSettings.DecimalSeparator := '.';                                                          //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  SL.Free;                                                                                           //
                                                                                                     //
  SetChipsID();                                                                                      //
  CalcChips();                                                                                       //
                                                                                                     //
  Result := True;                                                                                    //
end;                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.SaveSTS(const STSfName: TFileName): Boolean;                                         //
const                                                                                                //
  CR = AnsiString(#13#10);                                                                           //
var                                                                                                  //
  INIfName: TIniFile;                                                                                //
  X, Y: WORD;                                                                                        //
  n, P: WORD;                                                                                        //
  FS: TFileStream;                                                                                   //
  DateTime: TDateTime;                                                                               //
  Str: AnsiString;                                                                                   //
  tfName: TFileName;                                                                                 //
begin                                                                                                //
  Result := True;                                                                                    //
                                                                                                     //
  FormatSettings.DecimalSeparator := ',';                                                            //
                                                                                                     //
  tfName := GetFreeFileName(STSfName);                                                               //
                                                                                                     //
  INIfName := TIniFile.Create(tfName);                                                               //
  with INIfName do                                                                                   //
  begin                                                                                              //
    WriteString ('Main', 'OKR', OKR);                                                                //
    WriteString ('Main', 'Code', Code);                                                              //
    WriteString ('Main', 'MPW', MPW);                                                                //
    WriteString ('Main', 'MPWPos', MPWPos);                                                          //
    WriteString ('Main', 'Device', Device);                                                          //
    WriteString ('Main', 'DscrDev', DscrDev);                                                        //
    WriteString ('Main', 'MSystem', MeasSystem);                                                     //
    WriteString ('Main', 'Prober', Prober);                                                          //
                                                                                                     //
    WriteInteger('Main', 'Diametr', Diameter);                                                       //
    WriteInteger('Main', 'ChipSizeX', Round(StepX*1000));                                            //
    WriteInteger('Main', 'ChipSizeY', Round(StepY*1000));                                            //
                                                                                                     //
    WriteString ('Main', 'Lot', NLot);                                                               //
    WriteString ('Main', 'Wafer', Num);                                                              //
    WriteInteger('Main', 'WorkPlace', NWPlace);                                                      //
    WriteString ('Main', 'Operator', NOperator);                                                     //
    WriteString ('Main', 'Date', TimeDate);                                                          //
    WriteString ('Main', 'Condition', Condition);                                                    //
    WriteString ('Main', 'Info', Info);                                                              //
                                                                                                     //
    WriteInteger('Add', 'CadreX',  Cadre.ScaleX);                                                    //
    WriteInteger('Add', 'CadreY',  Cadre.ScaleY);                                                    //
    WriteInteger('Add', 'OffsetX', Cadre.StartX);                                                    //
    WriteInteger('Add', 'OffsetY', Cadre.StartY);                                                    //
    WriteInteger('Add', 'MaxX', Length(Chip[0]));                                                    //
    WriteInteger('Add', 'MaxY', Length(Chip));                                                       //
    WriteInteger('Add', 'BaseChipX', BaseChip.X);                                                    //
    WriteInteger('Add', 'BaseChipY', BaseChip.Y);                                                    //
    WriteInteger('Add', 'Path', Direct);                                                             //
    WriteInteger('Add', 'Cut', CutSide);                                                             //
                                                                                                     //
    Free;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  FS := TFileStream.Create(tfName, fmOpenWrite+fmShareDenyNone);                                     //
  FS.Position := FS.Size;                                                                            //
  FS.Write(CR, 2);                                                                                   //
  Str := '[StatusNames]';                                                                            //
  FS.Write(Pointer(Str)^, Length(Str));                                                              //
  FS.Write(CR, 2);                                                                                   //
  if StatusNamesSL.Count > 0 then                                                                    //
    for n := 0 to StatusNamesSL.Count-1 do                                                           //
    begin                                                                                            //
      FS.Write(Pointer(StatusNamesSL.Strings[n])^, Length(StatusNamesSL.Strings[n]));                //
      FS.Write(CR, 2);                                                                               //
    end;                                                                                             //
                                                                                                     //
  if Length(TestsParams) > 0 then                                                                    //
  begin                                                                                              //
    FS.Write(CR, 2);                                                                                 //
    Str := '[TestsParams]';                                                                          //
    FS.Write(Pointer(Str)^, Length(Str));                                                            //
    FS.Write(CR, 2);                                                                                 //
    for n := 0 to Length(TestsParams)-1 do                                                           //
    begin                                                                                            //
      if TestsParams[n].Norma.Min <> -NotSpec then                                                   //
        Str := IntToStr(n)+'='+FormatFloat('0.000', TestsParams[n].Norma.Min)+';'                    //
      else Str := IntToStr(n)+'=N;';                                                                 //
      if TestsParams[n].Norma.Max <> NotSpec then                                                    //
        Str := Str+FormatFloat('0.000', TestsParams[n].Norma.Max)+';'                                //
      else Str := Str+'N;';                                                                          //
      Str := Str+IntToStr(TestsParams[n].Status);                                                    //
      FS.Write(Pointer(Str)^, Length(Str));                                                          //
      FS.Write(CR, 2);                                                                               //
    end;                                                                                             //
  end;                                                                                               //
                                                                                                     //
  FS.Write(CR, 2);                                                                                   //
  Str := '[ChipsParams]';                                                                            //
  FS.Write(Pointer(Str)^, Length(Str));                                                              //
  FS.Write(CR, 2);                                                                                   //
  if Length(TestsParams) > 0 then                                                                    //
    for n := 0 to Length(TestsParams)-1 do                                                           //
    begin                                                                                            //
      Str := TestsParams[n].Name+#9;                                                                 //
      FS.Write(Pointer(Str)^, Length(Str));                                                          //
    end;                                                                                             //
  Str := 'Status'+#9;                                                                                //
  FS.Write(Pointer(Str)^, Length(Str));                                                              //
  Str := 'X'+#9;                                                                                     //
  FS.Write(Pointer(Str)^, Length(Str));                                                              //
  Str := 'Y'+CR;                                                                                     //
  FS.Write(Pointer(Str)^, Length(Str));                                                              //
                                                                                                     //
  for Y := 0 to Length(Chip)-1 do                                                                    //
    for X := 0 to Length(Chip[0])-1 do                                                               //
      with Chip[Y, X] do                                                                             //
        if Status <> 2 then                                                                          //
        begin                                                                                        //
          Str := '';                                                                                 //
          if Length(ChipParams) > 0 then                                                             //
            for n := 0 to Length(ChipParams)-1 do Str := Str+FormatFloat('0.000', ChipParams[n].Value)+#9 //
          else                                                                                       //
            if Length(TestsParams) > 0 then                                                          //
              for n := 0 to Length(TestsParams)-1 do Str := Str+FormatFloat('0.000', 0.0)+#9;        //
          Str := Str+IntToStr(Status)+#9+IntToStr(X)+#9+IntToStr(Y);                                 //
          FS.Write(Pointer(Str)^, Length(Str));                                                      //
          FS.Write(CR, 2);                                                                           //
        end;                                                                                         //
                                                                                                     //
  FS.Free;                                                                                           //
                                                                                                     //
//  DateTime := StrToDateTime(TimeDate+' 12:00:00');                                                   //
//  FileSetDate(tmpfName, DateTimeToFileDate(DateTime));                                               //
                                                                                                     //
  FormatSettings.DecimalSeparator := '.';                                                            //
end;                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadNI(const TXTfName: TFileName): Boolean;                                          //
var                                                                                                  //
  n, m: DWORD;                                                                                       //
  Str, NumChip, PrevChip: String;                                                                    //
  SL: TStringList;                                                                                   //
  HeaderCount, X, Y, k: WORD;                                                                        //
  FirstTime: Boolean;                                                                                //
begin                                                                                                //
  Result := False;                                                                                   //
                                                                                                     //
  fName := TXTfName;                                                                                 //
                                                                                                     //
  HeaderCount := LoadNIHeader();                                                                     //
  if HeaderCount = 0 then                                                                            //
  begin                                                                                              //
    ErrMess(Handle, '������ �������� ���������!');                                                   //
//    Init;                                                                                            //
    Exit;                                                                                            //
  end;                                                                                               //
                                                                                                     //
  SL := TStringList.Create;                                                                          //
  SL.LoadFromFile(fName);                                                                            //
                                                                                                     //
  for n := 0 to HeaderCount-1 do SL.Delete(0); // ������ ���������                                   //
                                                                                                     //
  FirstTime := True;                                                                                 //
  m := 0;                                                                                            //
  for Y := 0 to Length(Chip)-1 do                                                                    //
    for X := 0 to Length(Chip[0])-1 do                                                               //
      for n := 0 to Length(TestsParams) do                                                           //
      begin                                                                                          //
        if m = SL.Count then Break;                                                                  //
                                                                                                     //
        Str := Trim(SL.Strings[m]);                                                                  //
        if Str = '' then Continue;                                                                   //
                                                                                                     //
        NumChip := Copy(Str, 1, Pos(#9, Str)-1);                                                     //
        Delete(Str, 1, Pos(#9, Str)); // ������ ����� ���������                                      //
        Delete(Str, 1, Pos(#9, Str)); // ������ ����� �����                                          //
        Delete(Str, 1, Pos(#9, Str)); // ������ �������� ���������                                   //
                                                                                                     //
        if FirstTime then                                                                            //
        begin                                                                                        //
          PrevChip := NumChip;                                                                       //
          FirstTime := False;                                                                        //
        end;                                                                                         //
                                                                                                     //
        if NumChip <> PrevChip then                                                                  //
        begin                                                                                        //
          PrevChip := NumChip;                                                                       //
                                                                                                     //
          for k := n to Length(TestsParams)-1 do                                                     //
            Chip[Y, X].ChipParams[k].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[k].Value, TestsParams[k].Norma.Min, TestsParams[k].Norma.Max);
                                                                                                     //
          Break;                                                                                     //
        end;                                                                                         //
                                                                                                     //
        try                                                                                          //
          Chip[Y, X].ChipParams[n].Value := StrToFloat(Trim(Copy(Str, 1, Pos(#9, Str))));            //
        except                                                                                       //
          Chip[Y, X].ChipParams[n].Value := NotSpec;                                                 //
        end;                                                                                         //
                                                                                                     //
        Inc(m);                                                                                      //
                                                                                                     //
        Chip[Y, X].ChipParams[n].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[n].Value, TestsParams[n].Norma.Min, TestsParams[n].Norma.Max);
                                                                                                     //
        if Chip[Y, X].Status < 2000 then                                                             //
          if Chip[Y, X].ChipParams[n].Stat <> 1 then Chip[Y, X].Status := 2000+n                     //
          else Chip[Y, X].Status := 1;                                                               //
                                                                                                     //
        PrevChip := NumChip;                                                                         //
      end;                                                                                           //
                                                                                                     //
  Result := True;                                                                                    //
                                                                                                     //
  SetChipsID;                                                                                        //
  CalcChips;                                                                                         //
                                                                                                     //
  SL.Free;                                                                                           //
                                                                                                     //
  FormatSettings.DecimalSeparator := '.';                                                            //
end;                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.AddNI(const TXTfName: TFileName): Boolean;                                           //
begin                                                                                                //
  Result := False;
end;                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadXML(const XMLfName: TFileName): Boolean;                                                  //
var                                                                                                           //
  n: DWORD;                                                                                                   //
  Str: String;                                                                                                //
  X, Y: WORD;                                                                                                 //
  P1, P2, P3: byte;                                                                                           //
  XMLDoc1: IXMLDocument;                                                                                      //
  SL: TStringList;                                                                                            //
begin                                                                                                         //
  Result := False;                                                                                            //
                                                                                                              //
  fName := XMLfName;                                                                                          //
                                                                                                              //
  FormatSettings.DecimalSeparator := '.';                                                                     //
                                                                                                              //
  try                                                                                                         //
    SL := TStringList.Create;                                                                                 //
    SL.LoadFromFile(fName);                                                                                   //
    SL.Strings[0] := '<?xml version="1.0" encoding="windows-1251"?>';                                         //
                                                                                                              //
    XMLDoc1 := TXMLDocument.Create(nil);                                                                      //
    XMLDoc1.XML := SL;                                                                                        //
    XMLDoc1.Active := True;                                                                                   //
                                                                                                              //
    Device := XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['WAFER_BATCH_ID'].Text;                 //
    Str := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['WAFER_OCR_ID'].Text);                //
    P1 := Pos('-', Str);                                                                                      //
    if P1 <> 0 then                                                                                           //
    begin                                                                                                     //
      NLot := Copy(Str, 1, P1-1);                                                                             //
      Num  := Copy(Str, P1+1, Length(Str)-P1);                                                                //
    end                                                                                                       //
    else NLot := Str;                                                                                         //
    Diameter := StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['WAFER_SIZE'].Text));   //
    case StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['FLAT_LOCATION'].Text)) of     //
        0: CutSide := 3;                                                                                      //
       90: CutSide := 2;                                                                                      //
      180: CutSide := 1;                                                                                      //
      270: CutSide := 4;                                                                                      //
    end;                                                                                                      //
    LDiameter := Diameter;                                                                                    //
    Radius  := Diameter/2;                          //                                                        //
    LRadius := Radius-(Diameter-LDiameter);         //                                                        //
    Chord   := Sqrt(Radius*Radius-LRadius*LRadius); //                                                        //
                                                                                                              //
    StepX := StrToFloat(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['XSTEP'].Text));         //
    StepY := StrToFloat(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['YSTEP'].Text));         //
                                                                                                              //
    X := StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['COLUMN_COUNT'].Text));        //
    Y := StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['ROW_COUNT'].Text));           //
    Str := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['FIRST_DIE'].Text);                   //
    P1 := Pos(',', Str);                                                                                      //
    if P1 <> 0 then                                                                                           //
    begin                                                                                                     //
      BaseChip.X := StrToInt(Copy(Str, 1, P1-1));                                                             //
      BaseChip.Y := StrToInt(Copy(Str, P1+1, Length(Str)-P1));                                                //
    end;                                                                                                      //
    Str := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['PROBE_START_DATETIME'].Text);        //
    P1 := PosEx('-', Str);                                                                                    //
    TimeDate := '.'+Copy(Str, 1, P1-1);                                                                       //
    P2 := PosEx('-', Str, P1+1);                                                                              //
    TimeDate := '.'+Copy(Str, P1+1, P2-P1-1)+TimeDate;                                                        //
    P3 := PosEx(' ', Str, P2+1);                                                                              //
    TimeDate := Copy(Str, P2+1, P3-P2-1)+TimeDate;                                                            //
    Prober := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['PROBE_DEVICE_NAME'].Text);        //
    Direct := 2; // ��� ����� 6290                                                                            //
    Str := XMLDoc1.DocumentElement.ChildNodes['WAFER_MAP'].Text;                                              //
    SetLength(TestsParams, 0);                                                                                //
                                                                                                              //
    SetLength(Chip, 0, 0);                                                                                    //
    SetLength(Chip, Y, X);                                                                                    //
    n := 1;                                                                                                   //
    for Y := 0 to Length(Chip)-1 do                                                                           //
      for X := 0 to Length(Chip[0])-1 do                                                                      //
      begin                                                                                                   //
        Chip[Y, X].Status := 2;                                                                               //
                                                                                                              //
        case Str[n] of                                                                                        //
          '.': Chip[Y, X].Status := 2;                                                                        //
          ':': Chip[Y, X].Status := 3;                                                                        //
          'X': Chip[Y, X].Status := 2000;                                                                     //
          '1': Chip[Y, X].Status := 1;                                                                        //
          '-': Chip[Y, X].Status := 4;                                                                        //
          '/': Chip[Y, X].Status := 4;                                                                        //
          'a': begin                                                                                          //
                 Chip[Y, X].Status := 10; // ������� ����� ���������                                          //
                 BaseChip.X := X;                                                                             //
                 BaseChip.Y := Y;                                                                             //
               end;                                                                                           //
        end;                                                                                                  //
        SetLength(Chip[Y, X].ChipParams, 0);                                                                  //
                                                                                                              //
        Inc(n);                                                                                               //
      end;                                                                                                    //
                                                                                                              //
    XMLDoc1.Active := False;                                                                                  //
    SL.Free;                                                                                                  //
  except                                                                                                      //
    XMLDoc1.Active := False;                                                                                  //
    SL.Free;                                                                                                  //
    ErrMess(Handle, '������ �������� �����!');                                                                //
//    Init;                                                                                                     //
    Exit;                                                                                                     //
  end;                                                                                                        //
                                                                                                              //
  Result := True;                                                                                             //
                                                                                                              //
  Normalize;                                                                                                  //
  SetChipsID;                                                                                                 //
  CalcChips;                                                                                                  //
end;                                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.AddXML(const XMLfName: TFileName): Boolean;                                                   //
var                                                                                                           //
  n, ErrCount: DWORD;                                                                                         //
  Str: String;                                                                                                //
  X, Y: WORD;                                                                                                 //
  P1, P2, P3: byte;                                                                                           //
  XMLDoc1: IXMLDocument;                                                                                      //
  SL: TStringList;                                                                                            //
  tmpWafer: TWafer;                                                                                           //
begin                                                                                                         //
  Result := False;                                                                                            //
                                                                                                              //
  tmpWafer := TWafer.Create(Handle);                                                                          //
  tmpWafer.fName := XMLfName;                                                                                 //
                                                                                                              //
  FormatSettings.DecimalSeparator := '.';                                                                     //
                                                                                                              //
  try                                                                                                         //
    SL := TStringList.Create;                                                                                 //
    SL.LoadFromFile(tmpWafer.fName);                                                                          //
    SL.Strings[0] := '<?xml version="1.0" encoding="windows-1251"?>';                                         //
                                                                                                              //
    XMLDoc1 := TXMLDocument.Create(nil);                                                                      //
    XMLDoc1.XML := SL;                                                                                        //
    XMLDoc1.Active := True;                                                                                   //
                                                                                                              //
    with tmpWafer do                                                                                          //
    begin                                                                                                     //
      Device := XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['WAFER_BATCH_ID'].Text;               //
      Str := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['WAFER_OCR_ID'].Text);              //
      P1 := Pos('-', Str);                                                                                    //
      if P1 <> 0 then                                                                                         //
      begin                                                                                                   //
        NLot := Copy(Str, 1, P1-1);                                                                           //
        Num  := Copy(Str, P1+1, Length(Str)-P1);                                                              //
      end                                                                                                     //
      else NLot := Str;                                                                                       //
      Diameter := StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['WAFER_SIZE'].Text)); //
      case StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['FLAT_LOCATION'].Text)) of   //
          0: CutSide := 3;                                                                                    //
         90: CutSide := 2;                                                                                    //
        180: CutSide := 1;                                                                                    //
        270: CutSide := 4;                                                                                    //
      end;                                                                                                    //
      LDiameter := Diameter;                                                                                  //
      Radius  := Diameter/2;                          //                                                      //
      LRadius := Radius-(Diameter-LDiameter);         //                                                      //
      Chord   := Sqrt(Radius*Radius-LRadius*LRadius); //                                                      //
                                                                                                              //
      StepX := StrToFloat(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['XSTEP'].Text));       //
      StepY := StrToFloat(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['YSTEP'].Text));       //
                                                                                                              //
//      StepX := StepX/1.27;  /////////////////////////////////////////                                         //
//      StepY := StepY/1.40;  //////////////////////////////////////////////                                    //
                                                                                                              //
      X := StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['COLUMN_COUNT'].Text));      //
      Y := StrToInt(Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['ROW_COUNT'].Text));         //
      Str := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['FIRST_DIE'].Text);                 //
      P1 := Pos(',', Str);                                                                                    //
      if P1 <> 0 then                                                                                         //
      begin                                                                                                   //
        BaseChip.X := StrToInt(Copy(Str, 1, P1-1));                                                           //
        BaseChip.Y := StrToInt(Copy(Str, P1+1, Length(Str)-P1));                                              //
      end;                                                                                                    //
      Str := Trim(XMLDoc1.DocumentElement.ChildNodes['HEADER'].ChildNodes['PROBE_START_DATETIME'].Text);      //
      P1 := PosEx('-', Str);                                                                                  //
      TimeDate := '.'+Copy(Str, 1, P1-1);                                                                     //
      P2 := PosEx('-', Str, P1+1);                                                                            //
      TimeDate := '.'+Copy(Str, P1+1, P2-P1-1)+TimeDate;                                                      //
      P3 := PosEx(' ', Str, P2+1);                                                                            //
      TimeDate := Copy(Str, P2+1, P3-P2-1)+TimeDate;                                                          //
      Direct := 2; // ��� ����� 6510                                                                          //
      Str := XMLDoc1.DocumentElement.ChildNodes['WAFER_MAP'].Text;                                            //
      SetLength(TestsParams, 0);                                                                              //
                                                                                                              //
      SetLength(Chip, 0, 0);                                                                                  //
      SetLength(Chip, Y, X);                                                                                  //
      n := 1;                                                                                                 //
      for Y := 0 to Length(Chip)-1 do                                                                         //
        for X := 0 to Length(Chip[0])-1 do                                                                    //
        begin                                                                                                 //
          Chip[Y, X].Status := 2;                                                                             //
                                                                                                              //
          case Str[n] of                                                                                      //
            '.': Chip[Y, X].Status := 2;                                                                      //
            ':': Chip[Y, X].Status := 3;                                                                      //
            'X': Chip[Y, X].Status := 2000;                                                                   //
            '1': Chip[Y, X].Status := 1;                                                                      //
            '-': Chip[Y, X].Status := 4;                                                                      //
            '/': Chip[Y, X].Status := 4;                                                                      //
            'a': begin                                                                                        //
//                   Chip[Y, X].Status := 0;                                                                    //
//                   BaseChip.X := X;                                                                           //
//                   BaseChip.Y := Y;                                                                           //
                 end;                                                                                         //
          end;                                                                                                //
          SetLength(Chip[Y, X].ChipParams, 0);                                                                //
                                                                                                              //
          Inc(n);                                                                                             //
        end;                                                                                                  //
    end;                                                                                                      //
                                                                                                              //
    XMLDoc1.Active := False;                                                                                  //
    SL.Free;                                                                                                  //
  except                                                                                                      //
    XMLDoc1.Active := False;                                                                                  //
    SL.Free;                                                                                                  //
    ErrMess(Handle, '������ �������� �����!');                                                                //
//    Init;                                                                                                     //
    Exit;                                                                                                     //
  end;                                                                                                        //
                                                                                                              //
  Result := True;                                                                                             //
                                                                                                              //
  tmpWafer.Normalize;                                                                                         //
  tmpWafer.SetChipsID;                                                                                        //
  tmpWafer.CalcChips;                                                                                         //
                                                                                                              //
  if tmpWafer.NMeased <> NMeased then                                                                         //
    if QuestMess(Handle, '����� '+IntToStr(NTotal)+' ����������, �������� '+IntToStr(tmpWafer.NMeased)+#13#10+'��� ����� ����������?') = IDNO then
    begin
      tmpWafer.Free;
      Exit;
    end;

  ErrCount := 0;
  for n := 0 to Length(tmpWafer.ChipN)-1 do
    if n < Length(ChipN) then
    begin
      Y := tmpWafer.ChipN[n].Y;
      X := tmpWafer.ChipN[n].X;

      if not EqualStatus(tmpWafer.Chip[Y, X].Status, Chip[ChipN[n].Y,  ChipN[n].X].Status) then Inc(ErrCount);

      tmpWafer.Chip[tmpWafer.ChipN[n].Y, tmpWafer.ChipN[n].X].Status := Chip[ChipN[n].Y, ChipN[n].X].Status;
      tmpWafer.Chip[Y, X].ChipParams := Chip[ChipN[n].Y,  ChipN[n].X].ChipParams;
    end;
  if ErrCount > 0 then ErrMess(Handle, IntToStr(ErrCount)+' ������������!');

  Chip := tmpWafer.Chip;                                                                                //
  Diameter  := tmpWafer.Diameter;                                                                       //
  LDiameter := tmpWafer.LDiameter;                                                                      //
  Radius    := tmpWafer.Radius;                                                                         //
  LRadius   := tmpWafer.LRadius;                                                                        //
  Chord     := tmpWafer.Chord;                                                                          //
  StepX     := tmpWafer.StepX;                                                                          //
  StepY     := tmpWafer.StepY;                                                                          //
  CutSide   := tmpWafer.CutSide;                                                                        //
  Direct    := tmpWafer.Direct;                                                                         //

//  TestsParams   := tmpWafer.TestsParams;
//  StatusNamesSL := tmpWafer.StatusNamesSL;
//  TestsParamsSL := tmpWafer.TestsParamsSL;
//  ColorParams   := tmpWafer.ColorParams;

  NLot := tmpWafer.NLot;
  Num  := tmpWafer.Num;

  SetChipsID;                                                                                           //
//  CalcChips; // ����������������!!!!
                                                                                                              //
  tmpWafer.Free;                                                                                              //
end;                                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadAGL(const AGLfName: TFileName): Boolean;                                  //
var                                                                                           //
  SL: TStringList;                                                                            //
  m, Y, X: DWORD;                                                                             //
  Str: string;                                                                                //
  n, NFC, NKK: WORD;                                                                          //
  OK_param: Boolean;                                                                          //
begin                                                                                         //
  Result := False;                                                                            //
                                                                                              //
  fName := AGLfName;                                                                          //
                                                                                              //
  if not LoadAGLHeader then                                                                   //
  begin                                                                                       //
    ErrMess(Handle, '������ �������� �����!');                                                //
//    Init;                                                                                     //
    Exit;                                                                                     //
  end;                                                                                        //
                                                                                              //
  SL := TStringList.Create;                                                                   //
  SL.LoadFromFile(fName);                                                                     //
                                                                                              //
  m := 0;                                                                                     //
  for Y := 0 to Length(Chip)-1 do                                                             //
    for X := 0 to Length(Chip[0])-1 do                                                        //
    begin                                                                                     //
      if m >= SL.Count then Break;                                                            //
                                                                                              //
      repeat                                                                                  //
        Str := Trim(SL.Strings[m]);                                                           //
        Inc(m);                                                                               //
                                                                                              //
        if m = SL.Count then Break;                                                           //
      until Pos('TESTFLOW STARTED', UpperCase(Str)) <> 0;                                     //
                                                                                              //
      if m >= SL.Count then Break;                                                            //
                                                                                              //
      NFC := 0;                                                                               //
      NKK := 0;                                                                               //
      n := 0;                                                                                 //
      OK_param := False;                                                                      //
      repeat                                                                                  //
        Str := Trim(SL.Strings[m]);                                                           //
        Inc(m);                                                                               //
                                                                                              //
        if Str = '' then Continue;                                                            //
                                                                                              //
        if Str[1] = '1' then                                                                  //
        begin                                                                                 //
          if Pos('FAILED', UpperCase(Str)) <> 0 then OK_param := False // �������� ������     //
                                                else OK_param := True; // �������� ����       //
                                                                                              //
          if (Pos('CONTINUITY', UpperCase(Str)) <> 0) or                                      //
             (Pos('CONTAKT',    UpperCase(Str)) <> 0) or                                      //
             (Pos('CONTACT',    UpperCase(Str)) <> 0) then                                    //
          begin                                                                               //
            if not OK_param then Chip[Y, X].Status := 10+NKK;                                 //
            Inc(NKK);                                                                         //
            Continue;                                                                         //
          end;                                                                                //
          if (Pos('FUNCTIONAL', UpperCase(Str)) <> 0) or                                      //
             (Pos('FUNCTION',   UpperCase(Str)) <> 0) or                                      //
             (Pos('FK',         UpperCase(Str)) <> 0) then                                    //
          begin                                                                               //
            if not OK_param then Chip[Y, X].Status := 3500+NFC;                               //
            Inc(NFC);                                                                         //
            Continue; // ������� ����� ������, ����� ������ �� � ����������
          end;                                                                                //
                                                                                              //
          if n < Length(Chip[Y, X].ChipParams) then                                           //
          begin                                                                               //
            Delete(Str, 1, Pos('`', Str)); // ������ ����� �����                              //
            Delete(Str, 1, Pos('`', Str)); // ������ �������� ���������                       //
            Delete(Str, 1, Pos('`', Str)); // ������ ������ ��� ���������                     //
                                                                                              //
            Delete(Str, 1, Pos('`', Str)); // ������ passed/FAILED                            //
            Delete(Str, 1, Pos('`', Str)); // ������ ������ ������                            //
            Str := Trim(Str);                                                                 //
            Str := Trim(Copy(Str, 1, Pos(' ', Str)-1));                                       //
            try                                                                               //
              Chip[Y, X].ChipParams[n].Value := StrToFloat(Str);                              //
            except                                                                            //
              Chip[Y, X].ChipParams[n].Value := NotSpec;                                      //
            end;                                                                              //
                                                                                              //
            if Chip[Y, X].ChipParams[n].Value <> NotSpec then                                 //
            begin                                                                             //
              if Chip[Y, X].Status < 10 then // ���� �� ���� NK � FK                          //
                if (Chip[Y, X].ChipParams[n].Value < TestsParams[n].Norma.Min) or             //
                   (Chip[Y, X].ChipParams[n].Value > TestsParams[n].Norma.Max)                //
                then Chip[Y, X].Status := 2000+n                                              //
                else Chip[Y, X].Status := 1;                                                  //

              Chip[Y, X].ChipParams[n].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[n].Value, TestsParams[n].Norma.Min, TestsParams[n].Norma.Max);
            end;


            if not OK_param then
              if ((Chip[Y, X].ChipParams[n].Value = 0.0) and
                  (TestsParams[n].Norma.Min = 0.0) and
                  (TestsParams[n].Norma.Max = 0.0))
                  or
                 ((Chip[Y, X].ChipParams[n].Value = 0.0) and
                  (TestsParams[n].Norma.Min = -NotSpec) and
                  (TestsParams[n].Norma.Max = NotSpec)) then
              begin
                Chip[Y, X].Status := 3500+NFC; // �������, ��� ��� ���� ��
                Inc(NFC);
                Continue;// ?
              end;
          end;                                                                                //
                                                                                              //
          Inc(n);                                                                             //
        end;                                                                                  //
      until Pos('TESTFLOW ENDED', UpperCase(Str)) <> 0;                                       //
    end;                                                                                      //
                                                                                              //
  Result := True;                                                                             //
                                                                                              //
  SetChipsID;                                                                                 //
  CalcChips;                                                                                  //
                                                                                              //
  SL.Free;                                                                                    //
end;                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.AddAGL(const AGLfName: TFileName): Boolean;                                   //
var                                                                                           //
  SL: TStringList;                                                                            //
  m, Y, X, Nm: DWORD;                                                                         //
  Str: string;                                                                                //
  n, NFC, NKK: WORD;                                                                          //
begin                                                                                         //
  Result := False;                                                                            //
                                                                                              //
  fName := AGLfName;                                                                          //
                                                                                              //
  if not AddAGLHeader then                                                                    //
  begin                                                                                       //
    ErrMess(Handle, '������ �������� �����!');                                                //
//    Init;                                                                                     //
    Exit;                                                                                     //
  end;                                                                                        //
                                                                                              //
  SL := TStringList.Create;                                                                   //
  SL.LoadFromFile(AGLfName);                                                                  //
                                                                                              //
  m := 0;                                                                                     //
  for Nm := 0 to Length(ChipN)-1 do                                                           //
  begin                                                                                       //
    X := ChipN[Nm].X;                                                                         //
    Y := ChipN[Nm].Y;                                                                         //
                                                                                              //
    if m >= SL.Count then Break;                                                              //
                                                                                              //
//      if not IsChip(Wafer.Chip[Y, X].Status) then Continue;                                   //
                                                                                              //
    repeat                                                                                    //
      Str := Trim(SL.Strings[m]);                                                             //
      Inc(m);                                                                                 //
                                                                                              //
      if m = SL.Count then Break;                                                             //
    until Pos('TESTFLOW STARTED', UpperCase(Str)) <> 0;                                       //
                                                                                              //
    if m >= SL.Count then Break;                                                              //
                                                                                              //
    NFC := 0;                                                                                 //
    NKK := 0;                                                                                 //
    n := 0;                                                                                   //
    repeat                                                                                    //
      Str := Trim(SL.Strings[m]);                                                             //
      Inc(m);                                                                                 //
                                                                                              //
      if Str = '' then Continue;                                                              //
                                                                                              //
      if Str[1] = '1' then                                                                    //
      begin                                                                                   //
        if (Pos('CONTINUITY', UpperCase(Str)) <> 0) or                                        //
           (Pos('CONTAKT',    UpperCase(Str)) <> 0) or                                        //
           (Pos('CONTACT',    UpperCase(Str)) <> 0) then                                      //
        begin                                                                                 //
          if Pos('FAILED', UpperCase(Str)) <> 0 then Chip[Y, X].Status := 10+NKK;             //
          Inc(NKK);                                                                           //
          Continue;                                                                           //
        end;                                                                                  //
        if (Pos('FUNCTIONAL', UpperCase(Str)) <> 0) or                                        //
           (Pos('FUNCTION',   UpperCase(Str)) <> 0) or                                        //
           (Pos('FK',         UpperCase(Str)) <> 0) then                                      //
        begin                                                                                 //
          if Pos('FAILED', UpperCase(Str)) <> 0 then Chip[Y, X].Status := 3500+NFC;           //
          Inc(NFC);                                                                           //
          Continue;                                                                           //
        end;                                                                                  //
                                                                                              //
        if n < Length(Chip[Y, X].ChipParams) then                                             //
        begin                                                                                 //
          Delete(Str, 1, Pos('`', Str)); // ������ ����� �����                                //
          Delete(Str, 1, Pos('`', Str)); // ������ �������� ���������                         //
          Delete(Str, 1, Pos('`', Str)); // ������ ������ ��� ���������                       //
                                                                                              //
          Delete(Str, 1, Pos('`', Str)); // ������ PASSED/FAILED                              //
          Delete(Str, 1, Pos('`', Str)); // ������ ������ ������                              //
          Str := Trim(Str);                                                                   //
          Str := Trim(Copy(Str, 1, Pos(' ', Str)-1));                                         //
          try                                                                                 //
            Chip[Y, X].ChipParams[n].Value := StrToFloat(Str);                                //
          except                                                                              //
            Chip[Y, X].ChipParams[n].Value := NotSpec;                                        //
          end;                                                                                //
                                                                                              //
          if Chip[Y, X].ChipParams[n].Value <> NotSpec then                                   //
          begin                                                                               //
            if Chip[Y, X].Status < 10 then // ���� �� ���� NK � FK                            //
              if (Chip[Y, X].ChipParams[n].Value < TestsParams[n].Norma.Min) or               //
                 (Chip[Y, X].ChipParams[n].Value > TestsParams[n].Norma.Max)                  //
              then Chip[Y, X].Status := 2000+n                                                //
              else Chip[Y, X].Status := 1;                                                    //

            Chip[Y, X].ChipParams[n].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[n].Value, TestsParams[n].Norma.Min, TestsParams[n].Norma.Max);
          end;
        end;                                                                                  //
                                                                                              //
        Inc(n);                                                                               //
      end;                                                                                    //
    until Pos('TESTFLOW ENDED', UpperCase(Str)) <> 0;                                         //
  end;                                                                                        //
                                                                                              //
  Result := True;                                                                             //
                                                                                              //
  SetChipsID;                                                                                 //
  CalcChips;                                                                                  //
                                                                                              //
  SL.Free;                                                                                    //
end;                                                                                          //
////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.DetectXLS(const XLSfName: TFileName): byte;                                                                 //
var                                                                                                                         //
  Ap: OleVariant;                                                                                                           //
begin                                                                                                                       //
  Result := 0;                                                                                                              //
                                                                                                                            //
//  if CLSIDFromProgID(PWideChar(WideString(GetExcelAppName2)), ClassID) <> S_OK then                    //
//  begin                                                                                                //
//    if Assigned(OnEvent) then OnEvent(evError, 'Excel �� ������!');                                    //
//    Exit;                                                                                              //
//  end;                                                                                                 //

//  try                                                                                                  //
//    Excel := GetActiveOleObject(GetExcelAppName2);                                                     //
//  except                                                                                               //
//    Excel := CreateOleObject(GetExcelAppName2);                                                        //
//  end;                                                                                                 //
//  Excel.DisplayAlerts := False; // �������� ����� ��������������                                       //

  try                                                                                                                       //
    Ap := CreateOleObject('Excel.Application'); // �� ��������� (��������� ����)
  except                                                                                                                    //
    ErrMess(Handle, '�� ������� ��������� MS Excel.');                                                                      //
    Exit;                                                                                                                   //
  end;                                                                                                                      //
                                                                                                                            //
  Ap.DisplayAlerts := False;                                                                                                //
  Ap.Workbooks.Open(XLSfName, 0, True);                                                                                     //
                                                                                                                            //
  if AnsiLowerCase(Ap.Workbooks[1].Sheets[1].UsedRange.Cells[1, 1].Value) = 'pixan' then Result := 6                        //
  else                                                                                                                      //
    if (AnsiLowerCase(Ap.Workbooks[1].Sheets[1].UsedRange.Cells[1, 2].Value) = '�������') and                               //
       (AnsiLowerCase(Ap.Workbooks[1].Sheets[1].UsedRange.Cells[1, 8].Value) = '��������') then Result := 5; // Formula HF3 //
                                                                                                                            //
  Ap.Quit;                                                                                                                  //
end;                                                                                                                        //
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadXLS(const XLSfName: TFileName): Boolean;
var
  Ap, ActSheet, FData: OleVariant;
  aRows, SheetsCount, i, m, k, n, nStr, StrCount, ContCount, ChipsCount, X, Y, X1, Y1: Integer;
  TmpStr: string;
begin
  Result := False;

//  if CLSIDFromProgID(PWideChar(WideString(GetExcelAppName2)), ClassID) <> S_OK then                    //
//  begin                                                                                                //
//    if Assigned(OnEvent) then OnEvent(evError, 'Excel �� ������!');                                    //
//    Exit;                                                                                              //
//  end;                                                                                                 //

//  try                                                                                                  //
//    Excel := GetActiveOleObject(GetExcelAppName2);                                                     //
//  except                                                                                               //
//    Excel := CreateOleObject(GetExcelAppName2);                                                        //
//  end;                                                                                                 //
//  Excel.DisplayAlerts := False; // �������� ����� ��������������                                       //

  try
    Ap := CreateOleObject('Excel.Application'); // �� ��������� (��������� ����)
  except
    ErrMess(Handle, '�� ������� ��������� MS Excel!');
    Exit;
  end;

  fName := XLSfName;
  MeasSystem := 'Formula HF3';
  Direct := 2;

  Ap.DisplayAlerts := False;
  Ap.Workbooks.Open(XLSfName, 0, True);

  X := 0;
  Y := 0;
  StrCount := 0;
  SheetsCount := Ap.Workbooks[1].Sheets.Count; // ���-�� ������
  for i := 0 to SheetsCount-1 do
  begin
    ActSheet := Ap.Workbooks[1].Sheets[i+1];
    aRows := ActSheet.UsedRange.Rows.Count;  // ���-�� �����
    FData := ActSheet.UsedRange.Value; // ������� ���� � ������

    Device   := FData[2, 13];
    NLot     := FData[2, 15];
    Num      := FData[2, 16];
    TimeDate := FData[2, 20];

    nStr := 0;
    ContCount := 0;
    n := 0;
    StrCount := 0;
    k := 0;
    ChipsCount := 0;
    for m := 2 to aRows do
    begin
      if Pos('�������', AnsiLowerCase(FData[m, 3])) <> 0 then Inc(n); // ������� ���������������
      Inc(k); // ������� ����������

      if FData[m, 1] = '1' then
      begin
        if n > ContCount then ContCount := n; // ������ ����. ���-��
        n := 0;                               // ���������������

        Inc(ChipsCount); // ������ ���-�� �����

        if k > StrCount then
        begin
          StrCount := k; // ������ ����. ���-�� ����������
          nStr := m; // � ����� ������ � ������� �� �������
        end;
        k := 0;
      end;
    end;

    NTotal := ChipsCount; // ���-�� �����

    SetLength(TestsParams, StrCount-1);
    for m := nStr-StrCount to nStr do
    begin
      TestsParams[m-nStr+StrCount].Name := FData[m, 3];

      TmpStr := Trim(FData[m, 4]);
      if TmpStr <> '' then TestsParams[m-nStr+StrCount].Norma.Min := StrToFloat(TmpStr)
                      else TestsParams[m-nStr+StrCount].Norma.Min := -NotSpec;
      TmpStr := Trim(FData[m, 5]);
      if TmpStr <> '' then TestsParams[m-nStr+StrCount].Norma.Max := StrToFloat(TmpStr)
                      else TestsParams[m-nStr+StrCount].Norma.Max := NotSpec;
    end;

    X1 := Ceil(sqrt(NTotal));
    Y1 := X1;
    SetLength(Chip, 0, 0);
    SetLength(Chip, Y1, X1);
    for Y1 := 0 to Length(Chip)-1 do      // �������
      for X1 := 0 to Length(Chip[0])-1 do // ������
      begin                                     // �����
        Chip[Y1, X1].Status := 2;         //
        Chip[Y1, X1].ID     := 0;         //
//        Chip[Y1, X1].ShowGr := 0;         //
        SetLength(Chip[Y1, X1].ChipParams, Length(TestsParams));
      end;

    n := 0;
    for m := 3 to aRows do
    begin
      if FData[m, 1] = '1' then
      begin
        n := 0;

      end;



    end;



      m := 2;
      while m < aRows do
      begin
        if Device = '' then Device := FData[m, 13]
        else
          if Device <> FData[m, 13] then
          begin
            ErrMess(Handle, '������������ ������� �� ����� �'+IntToStr(i+1));

            Ap.Quit;
//            Init;
            Exit;
          end;

        Inc(m);

        for k := 0 to Length(TestsParams)-1 do
        begin
          if (X = 0) and (Y = 0) then
          begin
            TestsParams[m-3].Name := FData[m, 3];

            TmpStr := Trim(FData[m, 4]);
            if TmpStr <> '' then TestsParams[m-3].Norma.Min := StrToFloat(TmpStr)
                            else TestsParams[m-3].Norma.Min := -NotSpec;
            TmpStr := Trim(FData[m, 5]);
            if TmpStr <> '' then TestsParams[m-3].Norma.Max := StrToFloat(TmpStr)
                            else TestsParams[m-3].Norma.Max := NotSpec;

          end;
          {
          try
            Chip[Y, X].ChipParams[k].Value := FData[m, 6];
          except
            Chip[Y, X].ChipParams[k].Value := NotSpec;
          end;
          }
          TmpStr := Trim(FData[m, 6]);
          if TmpStr <> '' then Chip[Y, X].ChipParams[k].Value := StrToFloat(TmpStr)
                          else Chip[Y, X].ChipParams[k].Value := NotSpec;

          if Chip[Y, X].ChipParams[k].Value <> NotSpec then
          begin
            if Chip[Y, X].Status < 2000 then
              if (Chip[Y, X].ChipParams[k].Value < TestsParams[k].Norma.Min) or
                 (Chip[Y, X].ChipParams[k].Value > TestsParams[k].Norma.Max)
              then Chip[Y, X].Status := 2000+k
              else Chip[Y, X].Status := 1;

            Chip[Y, X].ChipParams[k].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[k].Value, TestsParams[k].Norma.Min, TestsParams[k].Norma.Max);
          end;

          Inc(m);
        end;

        if X = Length(Chip[0])-1 then
        begin
          Inc(Y);
          X := 0;
        end
        else Inc(X);
      end;
  end;

//  Str := Ap.Range['B1'];
//  Str := FData[1, 2];
//  MessageBox(0, PChar(Str), '123', MB_OK);
//  MessageBox(0, PChar('Sheets = '+IntToStr(SheetsCount)+'    Rows = '+IntToStr(aRows)+'    Columns = '+IntToStr(aColumns)), '123', MB_OK);

  Ap.Quit;

  Result := True;

  SetChipsID;
  CalcChips;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.AddXLS(const XLSfName: TFileName): Boolean;
begin
  Result := False;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////////////
function TWafer.LoadXLSPxn(const XLSfName: TFileName): Boolean;
var
  Ap, ActSheet, FData: OleVariant;
  aRows, aColumn, n, nChip, MinPos, MaxPos, X, Y: DWORD;
  tmpStr1, tmpStr2: string;
  P: byte;
begin
  Result := False;


//  if CLSIDFromProgID(PWideChar(WideString(GetExcelAppName2)), ClassID) <> S_OK then                    //
//  begin                                                                                                //
//    if Assigned(OnEvent) then OnEvent(evError, 'Excel �� ������!');                                    //
//    Exit;                                                                                              //
//  end;                                                                                                 //

//  try                                                                                                  //
//    Excel := GetActiveOleObject(GetExcelAppName2);                                                     //
//  except                                                                                               //
//    Excel := CreateOleObject(GetExcelAppName2);                                                        //
//  end;                                                                                                 //
//  Excel.DisplayAlerts := False; // �������� ����� ��������������                                       //


  try
    Ap := CreateOleObject('Excel.Application'); // �� ��������� (��������� - ����)
  except
    ErrMess(Handle, '�� ������� ��������� MS Excel!');
    Exit;
  end;

  fName := XLSfName;
  MeasSystem := '������';
//  Direct := 2;
  TimeDate := DateToStr(FileDateToDateTime(FileAge(XLSfName)));

  tmpStr1 := ExtractFileName(XLSfName);
  Delete(tmpStr1, Pos('.', tmpStr1), Length(tmpStr1));
  P := LastDelimiter('-', tmpStr1);
  if P > 0 then
    NLot := Copy(tmpStr1, 1, P-1);
  for n := P+1 to Length(tmpStr1) do
    if not (tmpStr1[n] in ['0'..'9']) then Break;
  if n > P then
    Num := Copy(tmpStr1, P+1, n-P-1);

//  Wafer.Diameter := 150;
//  Wafer.Device := '';

  Ap.DisplayAlerts := False;
  Ap.Workbooks.Open(XLSfName, 0, True);

  ActSheet := Ap.Workbooks[1].Sheets[1];
  aRows := ActSheet.UsedRange.Rows.Count; // ���-�� �����
  FData := ActSheet.UsedRange.Value; // ������� ���� � ������
  aColumn := ActSheet.UsedRange.Columns.Count ; // ���-�� ��������

  SetLength(TestsParams, aColumn-3);
  MinPos := 2;
  MaxPos := 3;
  tmpStr1 := FData[2, 1];
  tmpStr2 := FData[3, 1];
  if Pos('max', AnsiLowerCase(tmpStr1)) <> 0 then // ������ Max
  begin
    MaxPos := 2;
    MinPos := 3;
  end;

  for n := 0 to Length(TestsParams)-1 do
  begin
    TestsParams[n].Name := FData[4, n+3]; // ������ � ���������� ���������� (x � y ��������)

    TmpStr1 := Trim(FData[MinPos, n+3]); // ���. ����� ���������
    try
      TestsParams[n].Norma.Min := StrToFloat(TmpStr1);
    except
      TestsParams[n].Norma.Min := -NotSpec;
    end;

    TmpStr2 := Trim(FData[MaxPos, n+3]);
    try
      TestsParams[n].Norma.Max := StrToFloat(TmpStr2); // ����. ����� ���������
    except
      TestsParams[n].Norma.Max := NotSpec;
    end;
  end;

  NTotal := aRows-4; // ���-�� �����

  X := Ceil(sqrt(NTotal)); // ������� ����������
  Y := X;                        // ����� ������
  SetLength(Chip, 0, 0);
  SetLength(Chip, Y, X);
  nChip := 0;
  for Y := 0 to Length(Chip)-1 do      // �������
    for X := 0 to Length(Chip[0])-1 do // ������
    begin                              // �����
      Chip[Y, X].Status := 2;          //
      Chip[Y, X].ID     := 0;          //
//      Chip[Y, X].ShowGr := 0;          //
      Inc(nChip);
      if nChip <= NTotal then
        SetLength(Chip[Y, X].ChipParams, Length(TestsParams));
    end;

  nChip := 0;
  for Y := 0 to Length(Chip)-1 do
    for X := 0 to Length(Chip[0])-1 do
    begin
      if nChip < NTotal then
        for n := 0 to Length(TestsParams)-1 do
        begin
          TmpStr1 := Trim(FData[nChip+5, n+3]);
          try
            Chip[Y, X].ChipParams[n].Value := StrToFloat(TmpStr1);
          except
            Chip[Y, X].ChipParams[n].Value := NotSpec;
          end;

          Chip[Y, X].ChipParams[n].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[n].Value, TestsParams[n].Norma.Min, TestsParams[n].Norma.Max);

          if Chip[Y, X].Status < 2000 then
            if Chip[Y, X].ChipParams[n].Stat <> 1 then Chip[Y, X].Status := 2000+n
            else Chip[Y, X].Status := 1;
        end;

      Inc(nChip);
    end;

//  MessageBox(Handle, PAnsiChar(DateToStr(FileDateToDateTime(FileAge(XLSfName)))), '111!', MB_OK);

  Ap.Quit;

  Result := True;

  SetChipsID;
  CalcChips;
end;
///////////////////////////////////////////////////////////////////////////////////////////////////////


function TWafer.AddNorms(tParams: TTestsParams): Boolean;
var
  X, Y, n: DWORD;
  Str1, Str2: string;
begin
  Result := False;

  if Length(tParams) <> Length(TestsParams) then
  begin
    ErrMess(Handle, '�� ��������� ���-�� ������');
    Exit;
  end;

  for n := 0 to Length(TestsParams)-1 do
  begin
    Str1 := Copy(TestsParams[n].Name, Pos(' ', TestsParams[n].Name), Length(TestsParams[n].Name)); // ����� ����� �����
    Str1 := StringReplace(Str1, ' ', '', [rfReplaceAll, rfIgnoreCase]); // ������ �������
    Str1 := StringReplace(Str1, ',', '', [rfReplaceAll, rfIgnoreCase]); // ������ �������
    Str1 := StringReplace(Str1, '.', '', [rfReplaceAll, rfIgnoreCase]); // ������ �����

    Str2 := StringReplace(tParams[n].Name, ' ', '', [rfReplaceAll, rfIgnoreCase]); // ������ �������
    Str2 := StringReplace(Str1, ',', '', [rfReplaceAll, rfIgnoreCase]); // ������ �������
    Str2 := StringReplace(Str1, '.', '', [rfReplaceAll, rfIgnoreCase]); // ������ �����
    if Str1 <> Str2 then
    begin
      ErrMess(Handle, '�� ��������� ����� ������');
      Exit;
    end;

    TestsParams[n] := tParams[n];
  end;

  for Y := 0 to Length(Chip)-1 do
    for X := 0 to Length(Chip[0])-1 do
      if Chip[Y, X].Status <> 2 then
      begin
        if Chip[Y, X].Status = 2000 then Chip[Y, X].Status := 2; // ������� ������

        for n := 0 to Length(TestsParams)-1 do
        begin
          if Chip[Y, X].ChipParams[n].Stat <> 4 then
            Chip[Y, X].ChipParams[n].Stat := GetChipParamsStat(Chip[Y, X].ChipParams[n].Value,
                                                               TestsParams[n].Norma.Min,
                                                               TestsParams[n].Norma.Max);
          if Chip[Y, X].Status < 2000 then
            if Chip[Y, X].ChipParams[n].Stat <> 1 then Chip[Y, X].Status := 2000+n
                                                  else Chip[Y, X].Status := 1;
        end;
      end;

  CalcChips();

  Result := True;
end;



////////////////////////////////////////////////////////////////////////
function TWafer.GetChipParamsStat(const Val, Min, Max: Single): byte; //
begin                                                                 //
  Result := 1;                                                        //
                                                                      //
  if (Min = NotSpec)  and (Max = NotSpec) then Result := 0;           //
                                                                      //
  if (Min = NotSpec)  and (Max <> NotSpec) then                       //
    if Val > Max then Result := 3;                                    //
                                                                      //
  if (Min <> NotSpec) and (Max = NotSpec)  then                       //
    if Val < Min then Result := 2;                                    //
                                                                      //
  if (Min <> NotSpec) and (Max <> NotSpec) then                       //
  begin                                                               //
    if Val < Min then Result := 2;                                    //
    if Val > Max then Result := 3;                                    //
  end;                                                                //
end;                                                                  //
////////////////////////////////////////////////////////////////////////


{ TLot }

////////////////////////////////////////////////////////
constructor TLot.Create(Hndl: THandle);               //
begin                                                 //
  Handle := Hndl;                                     //
                                                      //
  BlankWafer := TWafer.Create(Handle);                //
end;                                                  //
////////////////////////////////////////////////////////
////////////////////////////////////////////////////////
destructor TLot.Destroy;                              //
var                                                   //
  n: byte;                                            //
begin                                                 //
  if Length(Wafer) > 0 then                           //
    for n := 0 to Length(Wafer)-1 do Wafer[n].Free(); //
                                                      //
  BlankWafer.Free();                                  //
                                                      //
  inherited;                                          //
end;                                                  //
////////////////////////////////////////////////////////

////////////////////////////////////////////////////////
procedure TLot.Init;                                  //
var                                                   //
  n: byte;                                            //
begin                                                 //
  if Length(Wafer) > 0 then                           //
    for n := 0 to Length(Wafer)-1 do Wafer[n].Free(); //
                                                      //
  SetLength(Wafer, 0);                                //
                                                      //
  LName   := '';                                      //
  LfName  := '';                                      //
                                                      //
  LDevice := '';                                      //
  LConfig := '';                                      //
end;                                                  //
////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
function TLot.LoadHistGroupsFromIni(): Boolean;                              //
var                                                                          //
  Buffer: array[0..MAX_PATH] of Char;                                        //
  IniFile: TIniFile;                                                         //
  InifName: TFileName;                                                       //
  SL: TStringList;                                                           //
  n, P: WORD;                                                                //
begin                                                                        //
  Result := False;                                                           //
                                                                             //
  GetModuleFileName(0, Buffer, MAX_PATH);                                    //
  InifName := ExtractFilePath(Buffer)+'Templates\Groups.ini';                //
  if not FileExists(InifName) then                                           //
  begin                                                                      //
    if Assigned(OnEvent) then OnEvent(evError, '�� ������ ���� Groups.ini'); //
//    Exit;                                                                  //
  end;                                                                       //
                                                                             //
  SetLength(HistGroup, 0);                                                   //
                                                                             //
  SL := TStringList.Create();                                                //
  IniFile := TIniFile.Create(InifName);                                      //
  try                                                                        //
    IniFile.ReadSectionValues('Groups', SL);                                 //
    if SL.Count > 0 then                                                     //
    begin                                                                    //
      SetLength(HistGroup, SL.Count);                                        //
                                                                             //
      for n := 0 to SL.Count-1 do                                            //
      begin                                                                  //
        P := Pos('=', SL[n]);                                                //
        HistGroup[n].Name := AnsiUpperCase(Copy(SL[n], 1, P-1));             //
        try                                                                  //
          HistGroup[n].Num := StrToInt(Copy(SL[n], P+1, Length(SL[n])));     //
        except                                                               //
          HistGroup[n].Num := 0;                                             //
        end;                                                                 //
      end;                                                                   //
    end;                                                                     //
  except                                                                     //
    IniFile.Free();                                                          //
  end;                                                                       //
  SL.Free();                                                                 //
                                                                             //
  Result := True;                                                            //
end;                                                                         //
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////
function TLot.CreateBlankWafer(const NWafer: WORD): Boolean;                             //
var                                                                                      //
  n, X, Y: DWORD;                                                                        //
begin                                                                                    //
  Result := False;                                                                       //
                                                                                         //
  if Wafer[NWafer] = nil then Exit;                                                      //
  if Wafer[NWafer].Diameter = 0 then Exit;                                               //
                                                                                         //
  BlankWafer.MeasSystem := Wafer[NWafer].MeasSystem;                                     //
  BlankWafer.Diameter   := Wafer[NWafer].Diameter;                                       //
  with BlankWafer do                                                                     //
  begin                                                                                  //
    if Diameter = 150 then LDiameter := 144.25                                           //
                      else LDiameter := Diameter;                                        //
    Radius  := Diameter/2;                                                               //
    LRadius := Radius-(Diameter-LDiameter);                                              //
    Chord   := Sqrt(Radius*Radius-LRadius*LRadius);                                      //
  end;                                                                                   //
  BlankWafer.Direct     := Wafer[NWafer].Direct;                                         //
  BlankWafer.CutSide    := Wafer[NWafer].CutSide;                                        //
  BlankWafer.NTotal     := Wafer[NWafer].NTotal;                                         //
                                                                                         //
  SetLength(BlankWafer.ChipN, 0);                                                        //
  SetLength(BlankWafer.ChipN, Length(Wafer[NWafer].ChipN));                              //
  for n := 0 to Wafer[NWafer].NTotal-1 do BlankWafer.ChipN[n] := Wafer[NWafer].ChipN[n]; //
                                                                                         //
                                                                                         //
  SetLength(BlankWafer.Chip, 0, 0);                                                      //
  SetLength(BlankWafer.Chip, Length(Wafer[NWafer].Chip), Length(Wafer[NWafer].Chip[0])); //
  for Y := 0 to Length(BlankWafer.Chip)-1 do                                             //
    for X := 0 to Length(BlankWafer.Chip[0])-1 do                                        //
    begin                                                                                //
      BlankWafer.Chip[Y, X].Status := Wafer[NWafer].Chip[Y, X].Status;                   //
      BlankWafer.Chip[Y, X].ID     := Wafer[NWafer].Chip[Y, X].ID;                       //
    end;                                                                                 //
                                                                                         //
  Result := True;                                                                        //
end;                                                                                     //
///////////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////////////////////////
function TLot.SaveXLS(const ToFirstFail, MapByParams: Boolean): Boolean;                               //
var                                                                                                    //
  Buffer: array[0..MAX_PATH] of Char;                                                                  //
  tmpfName: TFileName; // ���� �������                                                                 //
  n, m, i, Nm, X, Y, X1, Y1: DWORD;                                                                    //
  Excel, WorkBook1, Range1, Range2, tmpRange, Chart1, Chart2, Sheet, Cell1, Cell2: OleVariant;         //
  Chart: array of OleVariant;                                                                          //
  VarMass1, VarMass2: OleVariant;                                                                      //
  XLSfName: TFileName;                                                                                 //
  ClassID: TCLSID;                                                                                     //
  AvrSum, MinSum, MaxSum, StdSum, Qrt1Sum, MedSum, Qrt3Sum: array of Single;                           //
  HistParams: array of THistParams;                                                                    //
  QuantSum, OKSum, FailsSum, MeasSum: array of DWORD;                                                  //
  Col: TColor;                                                                                         //
  Str: string;                                                                                         //
  StX, LenX, StY, P: WORD;                                                                             //
  MinLength, MaxLength: array of DWORD;                                                                //
  GroupDiap: Single;                                                                                   //
  BottomSheet: WORD;                                                                                   //
begin                                                                                                  //
  Result := False;                                                                                     //
                                                                                                       //
  if Length(Wafer) = 0 then                                                                            //
  begin                                                                                                //
    if Assigned(OnEvent) then OnEvent(evError, '��� �������!');                                        //
    Exit;                                                                                              //
  end;                                                                                                 //
                                                                                                       //
  if CLSIDFromProgID(PWideChar(WideString(GetExcelAppName2)), ClassID) <> S_OK then                    //
  begin                                                                                                //
    if Assigned(OnEvent) then OnEvent(evError, 'Excel �� ������!');                                    //
    Exit;                                                                                              //
  end;                                                                                                 //
                                                                                                       //
  try                                                                                                  //
    Excel := GetActiveOleObject(GetExcelAppName2);                                                     //
  except                                                                                               //
    Excel := CreateOleObject(GetExcelAppName2);                                                        //
  end;                                                                                                 //
  Excel.DisplayAlerts := False; // �������� ����� ��������������                                       //
                                                                                                       //
  GetModuleFileName(0, Buffer, MAX_PATH);                                                              //
  tmpfName := ExtractFilePath(Buffer)+'Templates\TmpX3.xlsx';                                          //
  if not FileExists(tmpfName) then                                                                     //
  begin                                                                                                //
    if Assigned(OnEvent) then OnEvent(evError, '�� ������ ���� �������!');                             //
    Exit;                                                                                              //
  end;                                                                                                 //

////////////////////////////////////

  SetLength(HistParams, Length(Wafer[0].TestsParams)); // ������� ���-�� ����� ��� ����������

  for i := 0 to Length(HistParams)-1 do
  begin
    HistParams[i].Name := Wafer[0].TestsParams[i].Name;
    P := Pos(' ', HistParams[i].Name);
    if P > 0 then
      HistParams[i].ShortName := AnsiUpperCase(Copy(HistParams[i].Name, 1, P-1));

    HistParams[i].NGroups := 30;
                                                                                  //
    if Length(HistGroup) > 0 then                                                 //
      for m := 0 to Length(HistGroup) do                                          //
        if HistParams[i].ShortName = HistGroup[m].Name then                       //
        begin                                                                     //
          if HistGroup[m].Num > 0 then HistParams[i].NGroups := HistGroup[m].Num; //
          Break;                                                                  //
        end;                                                                      //
  end;

/////////////////////////////////                                                                      //
                                                                                                       //
  Workbook1 := Excel.WorkBooks.Open(tmpfName);                                                         //
                                                                                                       //
  FormatSettings.DecimalSeparator := ',';                                                              //
                                                                                                       //
////////////////////////////////////////////////////////////////////////////////                       //
//////////////////////////// * ���� "����� � ������" * /////////////////////////                       //
////////////////////////////////////////////////////////////////////////////////                       //
                                                                                                       //
  try                                                                                                  //
    Workbook1.Sheets['����� � ������'].Activate;                                                       //
  except                                                                                               //
    if Assigned(OnEvent) then OnEvent(evError, '� ������� ����������� ���� <����� � ������>');         //

  end;

  for i := 0 to Length(Wafer[0].TestsParams)-1 do
  begin
    WorkBook1.ActiveSheet.Cells[1, 2+i] := Wafer[0].TestsParams[i].Name;
    WorkBook1.ActiveSheet.Cells[1, 2+i].Interior.Color := clSkyBlue;
    WorkBook1.ActiveSheet.Cells[1, 2+i].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);

    if Wafer[0].TestsParams[i].Norma.Min <> -NotSpec then
    WorkBook1.ActiveSheet.Cells[2, 2+i] := Wafer[0].TestsParams[i].Norma.Min;
    WorkBook1.ActiveSheet.Cells[2, 2+i].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);

    if Wafer[0].TestsParams[i].Norma.Max <> NotSpec then
    WorkBook1.ActiveSheet.Cells[3, 2+i] := Wafer[0].TestsParams[i].Norma.Max;
    WorkBook1.ActiveSheet.Cells[3, 2+i].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);

    WorkBook1.ActiveSheet.Cells[4, 2+i] := Wafer[0].TestsParams[i].PUnit;
    WorkBook1.ActiveSheet.Cells[4, 2+i].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);

    WorkBook1.ActiveSheet.Cells[6, 2+i] := Wafer[0].TestsParams[i].PMode;
    WorkBook1.ActiveSheet.Cells[6, 2+i].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
  end;
  Cell1 := WorkBook1.ActiveSheet.Cells[1, 1];
  Cell2 := WorkBook1.ActiveSheet.Cells[2, Length(Wafer[0].TestsParams)+1];
  WorkBook1.ActiveSheet.Range[Cell1, Cell2].EntireColumn.AutoFit;


////////////////////////////////////////////////////////////////////////////////
/////////////////////////////// * ����� �������  * /////////////////////////////
////////////////////////////////////////////////////////////////////////////////

  try
    Workbook1.Sheets['��������'].Activate;
  except
    if Assigned(OnEvent) then OnEvent(evError, '� ������� ����������� ���� <��������>');

  end;

  if Assigned(OnEvent) then OnEvent(evInfo, '...��� ���������');

  Range1 := WorkBook1.ActiveSheet.Range['D1', 'D1' ]; // �������� ������ ��������
  Range2 := WorkBook1.ActiveSheet.Range['D4', 'D12']; //

  for n := 0 to Length(Wafer)-1 do
    with Wafer[n] do
    begin
      if n = 0 then
  //////////////////// * �������� ������ ��� ������� ������� * ///////////////////

      begin
        for i := 0 to Length(TestsParams)-1 do
        begin
          if i <> 0 then // ��������� ������ �� ���-�� ������ � ������� �� �����
          begin
            Range1.Copy(WorkBook1.ActiveSheet.Cells[1, i+4]); // �������� ��� ����� 1-� ������
            Range2.Copy(WorkBook1.ActiveSheet.Cells[4, i+4]); // (��� ��� ����)
          end;
          WorkBook1.ActiveSheet.Cells[ 1, i+4] := TestsParams[i].Name; // ��������
                                                                       // ��������
          WorkBook1.ActiveSheet.Cells[13, i+4] := TestsParams[i].Name; // ������
        end;

        tmpRange := WorkBook1.ActiveSheet; // �������� ���� ��� �����������
      end;

////////////////////////////////////////////////////////////////////////////////

      tmpRange.Copy(WorkBook1.ActiveSheet); // ��������� ����� ����
      WorkBook1.ActiveSheet.Name := Wafer[n].Num;

      Cell1 := WorkBook1.ActiveSheet.Cells[3, 1];                      //
      Cell2 := WorkBook1.ActiveSheet.Cells[17, 3+Length(TestsParams)]; // ��������
      Range1 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];             // ����
      Range1.Copy(WorkBook1.ActiveSheet.Cells[2+NTotal, 1]);           // ����������
      Range1.Clear;                                                    //

      Chart1 := Workbook1.ActiveSheet.ChartObjects(1);         // �������� ������
      Chart1.Top := WorkBook1.ActiveSheet.Rows[NTotal+17].Top; // ����
      Cell1 := WorkBook1.ActiveSheet.Cells[13+NTotal-1, 3];                     // �������
      Cell2 := WorkBook1.ActiveSheet.Cells[15+NTotal-1, 3+Length(TestsParams)]; // ��������
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                      // ��������
      Chart1.Chart.SetSourceData(Range2);                                       // ��� �������

      BottomSheet := NTotal+17+27; // ����� ������ ������

      SetLength(CalcsParams, Length(TestsParams));
      for i := 0 to Length(CalcsParams)-1 do
        with CalcsParams[i] do
        begin
          AvrVal    := 0.0;
          MinVal    :=  MaxSingle;
          MaxVal    := -MaxSingle;
          StdVal    := 0.0;
          ASum      := 0.0;
          QSum      := 0.0;
          Qrt1Val   := 0.0;
          MedVal    := 0.0;
          Qrt3Val   := 0.0;
          NOKVal    := 0;
          NFailsVal := 0;
          ValCount  := 0;
          SetLength(ValMass, NTotal);
        end;

//////////// ������� � variant �������

      VarMass1 := VarArrayCreate([1, NTotal, 1, Length(TestsParams)], varVariant); // ������ ��� ��������
      VarMass2 := VarArrayCreate([1, NTotal, 1, 3], varVariant); // ������ ��� ������ ���������, ������ � �/�

      for Nm := 0 to NTotal-1 do
      begin
        Y := ChipN[Nm].Y;
        X := ChipN[Nm].X;

        VarMass2[Nm+1, 1] := Nm+1;   // ����� ���������
        if Chip[Y, X].Status <> 1 then
        begin
          VarMass2[Nm+1, 2] := '�'; // �/�
          VarMass2[Nm+1, 3] := Chip[Y, X].Status-1999; // ������
        end;

        if Length(Chip[Y, X].ChipParams) > 0 then
          if Chip[Y, X].Status = 1 then // ���� ������ �������� //
          begin
            for i := 0 to Length(TestsParams)-1 do
            begin
              VarMass1[Nm+1, i+1] := Chip[Y, X].ChipParams[i].Value;

              with CalcsParams[i] do                                                                      //
              begin                                                                                       //
                if MinVal > Chip[Y, X].ChipParams[i].Value then MinVal := Chip[Y, X].ChipParams[i].Value; // ���.
                if MaxVal < Chip[Y, X].ChipParams[i].Value then MaxVal := Chip[Y, X].ChipParams[i].Value; // ����.
                                                                                                          //
                ASum := ASum+Chip[Y, X].ChipParams[i].Value;                                              // �������
                QSum := QSum+Sqr(Chip[Y, X].ChipParams[i].Value);                                         // �����

                ValMass[ValCount] := Chip[Y, X].ChipParams[i].Value;                                      // ��� ������, ���������

                Inc(NOKVal);                                                                              // ���-�� ������

                Inc(ValCount);
              end;
            end;
          end
          else                          // ���� ����������� �������� (�� �� ���� ����������) //
          begin
            if not ToFirstFail then // ���� �� �� 1-�� �����
              for i := 0 to Length(TestsParams)-1 do
              begin
                // if Chip[Y, X].ChipParams[i].Value = NotSpec then Continue; // ���� ����������� ��������
                if Chip[Y, X].ChipParams[i].Value = NotSpec then //
                begin                                            // ��������
                  VarMass1[Nm+1, i+1] := '';                     // ��
                  Chip[Y, X].ChipParams[i].Stat := 4;            // ���������
                  Continue;                                      //
                end;

                VarMass1[Nm+1, i+1] := Chip[Y, X].ChipParams[i].Value;

                with CalcsParams[i] do                                                                      //
                begin                                                                                       //
                  if MinVal > Chip[Y, X].ChipParams[i].Value then MinVal := Chip[Y, X].ChipParams[i].Value; // ���.
                  if MaxVal < Chip[Y, X].ChipParams[i].Value then MaxVal := Chip[Y, X].ChipParams[i].Value; // ����.
                                                                                                            //
                  ASum := ASum+Chip[Y, X].ChipParams[i].Value;                                              // �������
                  QSum := QSum+Sqr(Chip[Y, X].ChipParams[i].Value);                                         // �����

                  ValMass[ValCount] := Chip[Y, X].ChipParams[i].Value;                                      // ��� ������, ���������

                  if Chip[Y, X].ChipParams[i].Stat = 1 then Inc(NOKVal)                                     // ���-�� ������
                                                       else Inc(NFailsVal);                                 // ���-�� �����
                  Inc(ValCount);
                end;
              end
            else                    // ���� �� 1-�� �����
              for i := 0 to Length(TestsParams)-1 do
                if i < Chip[Y, X].Status-1999 then
                begin
                  if Chip[Y, X].ChipParams[i].Value = NotSpec then Continue; // ���� ����������� ��������

                  VarMass1[Nm+1, i+1] := Chip[Y, X].ChipParams[i].Value;

                  with CalcsParams[i] do
                  begin
                    if i < Chip[Y, X].Status-2000 then // ��� ����������� �������� �� ��������� ���., ����. � �����.
                    begin
                      if MinVal > Chip[Y, X].ChipParams[i].Value then MinVal := Chip[Y, X].ChipParams[i].Value; // ���.
                      if MaxVal < Chip[Y, X].ChipParams[i].Value then MaxVal := Chip[Y, X].ChipParams[i].Value; // ����.
                                                                                                                //
                      ASum := ASum+Chip[Y, X].ChipParams[i].Value;                                              // �������
                      QSum := QSum+Sqr(Chip[Y, X].ChipParams[i].Value);                                         // �����

                      ValMass[ValCount] := Chip[Y, X].ChipParams[i].Value;                                      // ��� ������, ���������

                      Inc(ValCount);
                    end;

                    if Chip[Y, X].ChipParams[i].Stat = 1 then Inc(CalcsParams[i].NOKVal)                        // ���-�� ������
                                                         else Inc(CalcsParams[i].NFailsVal);                    // ���-�� �����
                  end;
                end
                else
                begin
                  VarMass1[Nm+1, i+1] := '';
                  Chip[Y, X].ChipParams[i].Stat := 4; // �������� �� ���������
                end;
          end;
      end; // for Nm = 0 to Total


      for i := 0 to Length(CalcsParams)-1 do
      begin
        with Wafer[n].CalcsParams[i] do
        begin
          if ValCount > 0 then /////////////
          begin
            StdVal := Sqrt((ValCount*QSum-Sqr(ASum))/(ValCount*(ValCount-1))); // ����������� ����������
            AvrVal := ASum/ValCount;                                           // ������� ��������

            SetLength(ValMass, ValCount);
            SortMassByValue(ValMass); // ����������� ������� ��� ������ ...

            if Odd(ValCount) then MedVal :=  ValMass[ValCount div 2]                                 // �����
                             else MedVal := (ValMass[(ValCount div 2)-1]+ValMass[ValCount div 2])/2; // �������

            if Odd(ValCount div 2) then                                                   //
            begin                                                                         //
              Qrt1Val := ValMass[ValCount div 4];                                         // �����
              Qrt3Val := ValMass[3*(ValCount div 4)];                                     //
            end                                                                           // 1-� � 3-�
            else                                                                          //
            begin                                                                         // ��������
              Qrt1Val := (ValMass[(ValCount div 4)-1]+ValMass[ValCount div 4])/2;         //
              Qrt3Val := (ValMass[(3*(ValCount div 4))-1]+ValMass[3*(ValCount div 4)])/2; //
            end;                                                                          //

/////////////////////////////

            with HistParams[i] do // ��� �����������
            begin
              GroupDiap := (MaxVal-MinVal)/NGroups;

              SetLength(Group, NGroups);
              for m := 0 to Length(Group)-1 do            //
              begin                                       //
                Group[m].Num := 0;                        //
                Group[m].Min := MinVal+GroupDiap*m;       // ��������
                if m < Length(Group)-1 then               // ������
                  Group[m].Max := Group[m].Min+GroupDiap  // ���
                else                                      // ����������
                  Group[m].Max := MaxVal;                 //
              end;                                        //

              Nm := 0;                                           //
              for m := 0 to Length(ValMass)-1 do                 //
              begin                                              // �����������
                if ValMass[m] > Group[Nm].Max then               // ���������
                  while (ValMass[m] > Group[Nm].Max) do Inc(Nm); // ��
                                                                 // �������
                Inc(Group[Nm].Num);                              //
              end;                                               //
            end;
          end;
        end;
      end;

      /////////////// ** �������� ����������� �� ���������� ** /////////////////

      SetLength(Chart, 0);
      SetLength(Chart, Length(HistParams));
      Chart[0] := Workbook1.ActiveSheet.ChartObjects(2);             // ������
      Chart[0].Top := WorkBook1.ActiveSheet.Rows[BottomSheet+3].Top; // �����������
      Chart[0].Copy;

      for i := 0 to Length(HistParams)-1 do
        with HistParams[i] do
        begin
          WorkBook1.ActiveSheet.Cells[BottomSheet+2+26*i, 5].Font.Bold := True; //
          WorkBook1.ActiveSheet.Cells[BottomSheet+2+26*i, 5].Font.Size := 12;   // ���
          WorkBook1.ActiveSheet.Cells[BottomSheet+2+26*i, 5] := Name;           // �����������

          if i > 0 then
          begin
            WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i, 1].Select;
            WorkBook1.ActiveSheet.Paste;

            Chart[i] := Workbook1.ActiveSheet.ChartObjects(i+2);
          end;

          for m := 0 to Length(Group)-1 do
          begin                                                                                                                                   // �������
            WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i+m, 13+2*i] := FormatFloat('0.000',Group[m].Min)+'-'+FormatFloat('0.000',Group[m].Max); // �������
            WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i+m, 14+2*i] := Group[m].Num;                                                            // � ���-��
          end;

          Cell1 := WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i, 13+2*i];           //
          Cell2 := WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i+NGroups-1, 14+2*i]; // ������
          Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                        // ������
          Chart[i].Chart.SetSourceData(Range2);                                       //

          Range2.Font.Color := clWhite; // ������� ������ ����������
        end;

      //////////////////////////////////////////////////////////////////////////

      Cell1 := WorkBook1.ActiveSheet.Cells[2, 4];                            // �����
      Cell2 := WorkBook1.ActiveSheet.Cells[NTotal+1, 3+Length(TestsParams)]; // ����������
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                   // ���������
      Range2.NumberFormat := '0,000';                                        // �� �����
      Range2.Value := VarMass1;                                              // ����

      Cell1 := WorkBook1.ActiveSheet.Cells[2, 1];          // �����
      Cell2 := WorkBook1.ActiveSheet.Cells[NTotal+1, 3];   // �����, ������, �/�
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2]; // �� �����
      Range2.Value := VarMass2;                            // ����

      Cell1 := WorkBook1.ActiveSheet.Cells[NTotal+3, 4];                     // ������ ��� �����
      Cell2 := WorkBook1.ActiveSheet.Cells[NTotal+9, 3+Length(TestsParams)]; // �������, 1-� ��������,
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                   // �������, 3-� ��������,
      Range2.NumberFormat := '0,000';                                        // ���., ����., �����

      Cell1 := WorkBook1.ActiveSheet.Cells[NTotal+10, 4];                     // ������ ��� �����
      Cell2 := WorkBook1.ActiveSheet.Cells[NTotal+11, 3+Length(TestsParams)]; // ����,
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                    // %������
      Range2.NumberFormat := '0';                                             //

      Cell1 := WorkBook1.ActiveSheet.Cells[NTotal+13, 4];                     // ������ ��� �����
      Cell2 := WorkBook1.ActiveSheet.Cells[NTotal+14, 3+Length(TestsParams)]; // ������, ����,
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                    // ����� ��������,
      Range2.NumberFormat := '0';                                             // ���������� ������

      for i := 0 to Length(TestsParams)-1 do
      begin
        WorkBook1.ActiveSheet.Cells[NTotal+3,  4+i] := CalcsParams[i].AvrVal;  // �������
        WorkBook1.ActiveSheet.Cells[NTotal+4,  4+i] := CalcsParams[i].Qrt1Val; // 1-� ��������
        WorkBook1.ActiveSheet.Cells[NTotal+5,  4+i] := CalcsParams[i].MedVal;  // �������
        WorkBook1.ActiveSheet.Cells[NTotal+6,  4+i] := CalcsParams[i].Qrt3Val; // 3-� ��������
        if CalcsParams[i].MinVal <> MaxSingle then
          WorkBook1.ActiveSheet.Cells[NTotal+7,  4+i] := CalcsParams[i].MinVal;  // ���.
        if CalcsParams[i].MaxVal <> -MaxSingle then
          WorkBook1.ActiveSheet.Cells[NTotal+8,  4+i] := CalcsParams[i].MaxVal;  // ���.
        WorkBook1.ActiveSheet.Cells[NTotal+9,  4+i] := CalcsParams[i].StdVal;  // �����
        WorkBook1.ActiveSheet.Cells[NTotal+10, 4+i] := Wafer[n].NOK;           // ���� ???????????????????

        WorkBook1.ActiveSheet.Cells[NTotal+11, 4+i] := (Wafer[n].NOK*100)/NMeased; // %������ ??????????????
        WorkBook1.ActiveSheet.Cells[NTotal+13, 4+i] := CalcsParams[i].NOKVal;      // ������
        WorkBook1.ActiveSheet.Cells[NTotal+14, 4+i] := CalcsParams[i].NFailsVal;   // ����

        WorkBook1.ActiveSheet.Columns[4+i].EntireColumn.AutoFit; /////
      end;

      WorkBook1.ActiveSheet.Cells[NTotal+15, 4] := Wafer[n].NMeased; // ����� ��������
      WorkBook1.ActiveSheet.Cells[NTotal+16, 4] := Wafer[n].NOK;     // ����� ������

      if Assigned(OnEvent) then OnEvent(evOK, '���������� �������� �: '+Num);

      VarClear(VarMass1); // �������
      VarClear(VarMass2); // ���������� �������
    end; // for Wafer[n]

  try
    Workbook1.Sheets['��������'].Delete; // ������ ���� �������
  except

  end;

////////////////////////////////////////////////////////////////////////////////
//////////////////////////////// * ���� "�����" * //////////////////////////////
////////////////////////////////////////////////////////////////////////////////

  try
    Workbook1.Sheets['�����'].Activate;
  except
    if Assigned(OnEvent) then OnEvent(evError, '� ������� ����������� ���� <�����>');

  end;

  for n := 0 to Length(Wafer)-1 do
    with Wafer[n] do
    begin
      if n = 0 then
      begin
        tmpRange := WorkBook1.ActiveSheet.Range['C1', 'C25']; // �������� ������ ��������

        for i := 0 to Length(TestsParams)-1 do
        begin
          if i > 0 then // ��������� ������ �� ���-�� ������ � ������� �� �����
            tmpRange.Copy(WorkBook1.ActiveSheet.Cells[1, i+3]); // �������� ��� ����� 1-� ������

          WorkBook1.ActiveSheet.Cells[3, i+3] := TestsParams[i].Name;  // ��������
          WorkBook1.ActiveSheet.Cells[3, i+3].Columns.Autofit;         // ��������
                                                                       //
          WorkBook1.ActiveSheet.Cells[22, i+3] := TestsParams[i].Name; // ������
        end;

        Cell1 := WorkBook1.ActiveSheet.Cells[22+9*(Length(Wafer)-1), 2];                     // �������
        Cell2 := WorkBook1.ActiveSheet.Cells[24+9*(Length(Wafer)-1), 2+Length(TestsParams)]; // ��������
        Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                                 // ��������
        Chart1 := Workbook1.ActiveSheet.ChartObjects(1);                                     // ��� �������
        Chart1.Chart.SetSourceData(Range2);                                                  //

        if Length(Wafer) > 1 then // ��������, ���� ������ 1-� ��������
        begin
          Cell1 := WorkBook1.ActiveSheet.Cells[13, 1];                     //
          Cell2 := WorkBook1.ActiveSheet.Cells[25, 2+Length(TestsParams)]; // ��������
          Range1 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];             //
          if Length(Wafer) = 2 then
          begin                                                              // ���� �����
            Range1.Copy(WorkBook1.ActiveSheet.Cells[35, 1]);                 // ���������
            Range1.Clear;                                                    // � ������ �����
            Cell1 := WorkBook1.ActiveSheet.Cells[35, 1];                     // �����
            Cell2 := WorkBook1.ActiveSheet.Cells[47, 2+Length(TestsParams)]; // ������ �����������
            Range1 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];             // (������������ ������)
          end;
          Range1.Copy(WorkBook1.ActiveSheet.Cells[13+9*(Length(Wafer)-1), 1]); // ��������
          Range1.Clear;                                                        // � �����

          Chart1.Top := WorkBook1.ActiveSheet.Rows[27+9*(Length(Wafer)-1)].Top; // �������� ������ ����

          Cell1 := WorkBook1.ActiveSheet.Cells[4,  1];                     // ��������
          Cell2 := WorkBook1.ActiveSheet.Cells[12, 2+Length(TestsParams)]; // ������
          Range1 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];             // ��� �����������

          for i := 0 to Length(Wafer)-2 do                       // ���������
            Range1.Copy(WorkBook1.ActiveSheet.Cells[13+9*i, 1]); // ��� ���� �������
        end;
      end;

      BottomSheet := 27+9*(Length(Wafer)-1)+26; // ����� ������ ������ //////// +27

      WorkBook1.ActiveSheet.Cells[4+9*n, 1] := Wafer[n].Num;
      WorkBook1.ActiveSheet.Columns[1].EntireColumn.AutoFit;
      for i := 0 to Length(TestsParams)-1 do
      begin
        WorkBook1.ActiveSheet.Cells[4+ 9*n, 3+i] := Wafer[n].CalcsParams[i].AvrVal;   // �������
        WorkBook1.ActiveSheet.Cells[5+ 9*n, 3+i] := Wafer[n].CalcsParams[i].Qrt1Val;  // 1-� ��������
        WorkBook1.ActiveSheet.Cells[6+ 9*n, 3+i] := Wafer[n].CalcsParams[i].MedVal;   // �������
        WorkBook1.ActiveSheet.Cells[7+ 9*n, 3+i] := Wafer[n].CalcsParams[i].Qrt3Val;  // 3-� ��������
        if Wafer[n].CalcsParams[i].MinVal <> MaxSingle then
          WorkBook1.ActiveSheet.Cells[8+ 9*n, 3+i] := Wafer[n].CalcsParams[i].MinVal; // ���.
        if Wafer[n].CalcsParams[i].MaxVal <> -MaxSingle then
          WorkBook1.ActiveSheet.Cells[9+ 9*n, 3+i] := Wafer[n].CalcsParams[i].MaxVal; // ����.
        WorkBook1.ActiveSheet.Cells[10+9*n, 3+i] := Wafer[n].CalcsParams[i].StdVal;   // �����

        WorkBook1.ActiveSheet.Cells[11+9*n, 3+i] := Wafer[n].NOK;      // ���� ??????
        WorkBook1.ActiveSheet.Cells[12+9*n, 3+i] := (NOK*100)/NMeased; // %������ ??????????????

        WorkBook1.ActiveSheet.Columns[3+i].EntireColumn.AutoFit;
      end;

    end;

  SetLength(AvrSum,   0);
  SetLength(AvrSum,   Length(Wafer[0].TestsParams));
  SetLength(Qrt1Sum,  0);
  SetLength(Qrt1Sum,  Length(Wafer[0].TestsParams));
  SetLength(MedSum,   0);
  SetLength(MedSum,   Length(Wafer[0].TestsParams));
  SetLength(Qrt3Sum,  0);
  SetLength(Qrt3Sum,  Length(Wafer[0].TestsParams));
  SetLength(MinSum,   0);
  SetLength(MinSum,   Length(Wafer[0].TestsParams));
  SetLength(MaxSum,   0);
  SetLength(MaxSum,   Length(Wafer[0].TestsParams));
  SetLength(StdSum,   0);
  SetLength(StdSum,   Length(Wafer[0].TestsParams));
  SetLength(QuantSum, 0);
  SetLength(QuantSum, Length(Wafer[0].TestsParams));
  SetLength(OKSum,    0);
  SetLength(OKSum,    Length(Wafer[0].TestsParams));
  SetLength(FailsSum, 0);
  SetLength(FailsSum, Length(Wafer[0].TestsParams));
  SetLength(MeasSum,  0);
  SetLength(MeasSum,  Length(Wafer[0].TestsParams));

  SetLength(MinLength,  0);
  SetLength(MinLength,  Length(Wafer[0].TestsParams));
  SetLength(MaxLength,  0);
  SetLength(MaxLength,  Length(Wafer[0].TestsParams));

  for i := 0 to Length(Wafer[0].TestsParams)-1 do
  begin
    AvrSum[i]   := 0.0;
    Qrt1Sum[i]  := 0.0;
    MedSum[i]   := 0.0;
    Qrt3Sum[i]  := 0.0;
    MinSum[i]   := 0.0;
    MaxSum[i]   := 0.0;
    StdSum [i]  := 0.0;
    QuantSum[i] := 0;
    OKSum[i]    := 0;
    FailsSum[i] := 0;
    MeasSum[i]  := 0;

    MinLength[i] := 0; // ���� ����
    MaxLength[i] := 0; // �� �� ��������
  end;

  for i := 0 to Length(Wafer[0].TestsParams)-1 do
    for n := 0 to Length(Wafer)-1 do
      with Wafer[n] do
      begin
        AvrSum[i]   := AvrSum[i] +CalcsParams[i].AvrVal;
        Qrt1Sum[i]  := Qrt1Sum[i]+CalcsParams[i].Qrt1Val;
        MedSum[i]   := MedSum[i] +CalcsParams[i].MedVal;
        Qrt3Sum[i]  := Qrt3Sum[i]+CalcsParams[i].Qrt3Val;
        if CalcsParams[i].MinVal <> MaxSingle then
          MinSum[i] := MinSum[i]+CalcsParams[i].MinVal
        else
          Inc(MinLength[i]);

        if CalcsParams[i].MaxVal <> -MaxSingle then
          MaxSum[i] := MaxSum[i]+CalcsParams[i].MaxVal
        else
          Inc(MaxLength[i]);

        StdSum[i]   := StdSum[i] +CalcsParams[i].StdVal;
        QuantSum[i] := QuantSum[i]+NOK;                  // ���� ???????????????
        OKSum[i]    := OKSum[i]+CalcsParams[i].NOKVal;
        FailsSum[i] := FailsSum[i]+CalcsParams[i].NFailsVal;
        MeasSum[i]  := MeasSum[i]+NMeased;
      end;

  for i := 0 to Length(Wafer[0].TestsParams)-1 do
  begin
//////////////////////////////// * ������� * ///////////////////////////////////

    X := 4+9*Length(Wafer);
    WorkBook1.ActiveSheet.Cells[0+X, i+3] := AvrSum[i]/Length(Wafer);      // �������
    WorkBook1.ActiveSheet.Cells[1+X, i+3] := Qrt1Sum[i]/Length(Wafer);     // 1-� ��������
    WorkBook1.ActiveSheet.Cells[2+X, i+3] := MedSum[i]/Length(Wafer);      // �������
    WorkBook1.ActiveSheet.Cells[3+X, i+3] := Qrt3Sum[i]/Length(Wafer);     // 1-� ��������
    if Length(Wafer) <> MinLength[i] then
      WorkBook1.ActiveSheet.Cells[4+X, i+3] := MinSum[i]/(Length(Wafer)-MinLength[i]); // ���.
    if Length(Wafer) <> MaxLength[i] then
      WorkBook1.ActiveSheet.Cells[5+X, i+3] := MaxSum[i]/(Length(Wafer)-MaxLength[i]); // ����.
    WorkBook1.ActiveSheet.Cells[6+X, i+3] := StdSum[i]/Length(Wafer);      // �����
    WorkBook1.ActiveSheet.Cells[7+X, i+3] := QuantSum[i]/Length(Wafer);    // ���� ??????????
    WorkBook1.ActiveSheet.Cells[8+X, i+3] := (QuantSum[i]*100)/MeasSum[i]; // %������ ???????

/////////////////////////////// * �� ������ * //////////////////////////////////

    WorkBook1.ActiveSheet.Cells[10+X, i+3] := OKSum[i];             // ������
    WorkBook1.ActiveSheet.Cells[11+X, i+3] := FailsSum[i];          // ����
    WorkBook1.ActiveSheet.Cells[12+X, i+3] := OKSum[i]+FailsSum[i]; // �����
  end;

  for i := 0 to Length(HistParams)-1 do                        //
  begin                                                        // �������������
    SetLength(HistParams[i].AllValMass, 0);                    // �����
                                                               //
    HistParams[i].AllMinVal := Wafer[0].CalcsParams[i].MinVal; //
    HistParams[i].AllMaxVal := Wafer[0].CalcsParams[i].MaxVal; //

    SetLength(HistParams[i].Group, 0);
  end;

//////////////////// * ����� ����������� �� ���������� * ///////////////////////

  for i := 0 to Length(HistParams)-1 do                        //
  begin                                                        // �������������
    SetLength(HistParams[i].AllValMass, 0);                    // �����
                                                               //
    HistParams[i].AllMinVal := Wafer[0].CalcsParams[i].MinVal; //
    HistParams[i].AllMaxVal := Wafer[0].CalcsParams[i].MaxVal; //

    SetLength(HistParams[i].Group, 0);
  end;

  for n := 0 to Length(Wafer)-1 do                                                         //
    for i := 0 to Length(HistParams)-1 do                                                  //
      with Wafer[n].CalcsParams[i] do                                                      // ������
      begin                                                                                // ���������
        Nm := Length(HistParams[i].AllValMass);                                            // �� ����
        SetLength(HistParams[i].AllValMass, Nm+ValCount);                                  // ���������
          for m := 0 to Length(ValMass)-1 do HistParams[i].AllValMass[Nm+m] := ValMass[m]; //
                                                                                           //
        if HistParams[i].AllMinVal > MinVal then HistParams[i].AllMinVal := MinVal;        // ����� ���
        if HistParams[i].AllMaxVal < MaxVal then HistParams[i].AllMaxVal := MaxVal;        // � ����
      end;


  for i := 0 to Length(HistParams)-1 do
    with HistParams[i] do
    begin
      SortMassByValue(AllValMass); // ����������� �������

      GroupDiap := (AllMaxVal-AllMinVal)/NGroups; // �������� ������

      SetLength(Group, NGroups);
      for m := 0 to Length(Group)-1 do            //
      begin                                       //
        Group[m].Num := 0;                        //
        Group[m].Min := AllMinVal+GroupDiap*m;    // ��������
        if m < Length(Group)-1 then               // ������
          Group[m].Max := Group[m].Min+GroupDiap  // ���
        else                                      // ����������
          Group[m].Max := AllMaxVal;              //
      end;                                        //

      Nm := 0;                                              //
      for m := 0 to Length(AllValMass)-1 do                 //
      begin                                                 // �����������
        if AllValMass[m] > Group[Nm].Max then               // ���������
          while (AllValMass[m] > Group[Nm].Max) do Inc(Nm); // ��
                                                            // �������
        Inc(Group[Nm].Num);                                 //
      end;                                                  //
    end;                                                    //


  SetLength(Chart, 0);
  SetLength(Chart, Length(HistParams));
  Chart[0] := Workbook1.ActiveSheet.ChartObjects(2);             // ������
  Chart[0].Top := WorkBook1.ActiveSheet.Rows[BottomSheet+3].Top; // �����������
  Chart[0].Copy;

  for i := 0 to Length(HistParams)-1 do
    with HistParams[i] do
    begin
      WorkBook1.ActiveSheet.Cells[BottomSheet+2+26*i, 5].Font.Bold := True; //
      WorkBook1.ActiveSheet.Cells[BottomSheet+2+26*i, 5].Font.Size := 12;   // ���
      WorkBook1.ActiveSheet.Cells[BottomSheet+2+26*i, 5] := Name;           // �����������

      if i > 0 then
      begin
        WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i, 1].Select;
        WorkBook1.ActiveSheet.Paste;

        Chart[i] := Workbook1.ActiveSheet.ChartObjects(i+2);
      end;

      for m := 0 to Length(Group)-1 do
      begin                                                                                                                                   // �������
        WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i+m, 13+2*i] := FormatFloat('0.000',Group[m].Min)+'-'+FormatFloat('0.000',Group[m].Max); // �������
        WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i+m, 14+2*i] := Group[m].Num;                                                            // � ���-��
      end;

      Cell1 := WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i, 13+2*i];           //
      Cell2 := WorkBook1.ActiveSheet.Cells[BottomSheet+3+26*i+NGroups-1, 14+2*i]; // ������
      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];                        // ������
      Chart[i].Chart.SetSourceData(Range2);                                       //

      Range2.Font.Color := clWhite; // ������� ������ ����������
    end;


  if Assigned(OnEvent) then OnEvent(evOK, '>>> ���������� ��� ��������!');

////////////////////////////////////////////////////////////////////////////////
//////////////////////////// * ���� "�����������" * ////////////////////////////
////////////////////////////////////////////////////////////////////////////////

//  try
//    Workbook1.Sheets['�����������'].Activate;
//  except
//    if Assigned(OnEvent) then OnEvent(evError, '� ������� ����������� ���� <�����������>');
//
//  end;
//
//  for i := 0 to Length(HistParams)-1 do                        //
//  begin                                                        // �������������
//    SetLength(HistParams[i].AllValMass, 0);                    // �����
//                                                               //
//    HistParams[i].AllMinVal := Wafer[0].CalcsParams[i].MinVal; //
//    HistParams[i].AllMaxVal := Wafer[0].CalcsParams[i].MaxVal; //
//
//    SetLength(HistParams[i].Group, 0);
//  end;
//
//  for n := 0 to Length(Wafer)-1 do                                                         //
//    for i := 0 to Length(HistParams)-1 do                                                  //
//      with Wafer[n].CalcsParams[i] do                                                      // ������
//      begin                                                                                // ���������
//        Nm := Length(HistParams[i].AllValMass);                                            // �� ����
//        SetLength(HistParams[i].AllValMass, Nm+ValCount);                                  // ���������
//          for m := 0 to Length(ValMass)-1 do HistParams[i].AllValMass[Nm+m] := ValMass[m]; //
//                                                                                           //
//        if HistParams[i].AllMinVal > MinVal then HistParams[i].AllMinVal := MinVal;        // ����� ���
//        if HistParams[i].AllMaxVal < MaxVal then HistParams[i].AllMaxVal := MaxVal;        // � ����
//      end;
//
//
//  for i := 0 to Length(HistParams)-1 do
//    with HistParams[i] do
//    begin
//      SortMassByValue(AllValMass); // ����������� �������
//
//      GroupDiap := (AllMaxVal-AllMinVal)/NGroups; // �������� ������
//
//      SetLength(Group, NGroups);
//      for m := 0 to Length(Group)-1 do            //
//      begin                                       //
//        Group[m].Num := 0;                        //
//        Group[m].Min := AllMinVal+GroupDiap*m;    // ��������
//        if m < Length(Group)-1 then               // ������
//          Group[m].Max := Group[m].Min+GroupDiap  // ���
//        else                                      // ����������
//          Group[m].Max := AllMaxVal;              //
//      end;                                        //
//
//      Nm := 0;                                              //
//      for m := 0 to Length(AllValMass)-1 do                 //
//      begin                                                 // �����������
//        if AllValMass[m] > Group[Nm].Max then               // ���������
//          while (AllValMass[m] > Group[Nm].Max) do Inc(Nm); // ��
//                                                            // �������
//        Inc(Group[Nm].Num);                                 //
//      end;                                                  //
//    end;                                                    //
//
//  Chart[0] := Workbook1.ActiveSheet.ChartObjects(1);
//
//  Cell1 := WorkBook1.ActiveSheet.Cells[2, 1];
//  Cell2 := WorkBook1.ActiveSheet.Cells[23, 12];
//  Range1 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];
//
//  for i := 0 to Length(HistParams)-1 do
//    with HistParams[i] do
//    begin
//      WorkBook1.ActiveSheet.Cells[1+25*i, 6] := Name;
//
//      if i > 0 then
//      begin
//        Range1.Copy(WorkBook1.ActiveSheet.Cells[2+25*i, 1]); // ��������� ����
//
//        Chart[i] := Workbook1.ActiveSheet.ChartObjects(i+1);
//      end;
//
//      for m := 0 to Length(Group)-1 do
//      begin
//        WorkBook1.ActiveSheet.Cells[2+m, 16+i*3] := FormatFloat('0.000',Group[m].Min)+'-'+FormatFloat('0.000',Group[m].Max);
//        WorkBook1.ActiveSheet.Cells[2+m, 17+i*3] := Group[m].Num;
//      end;
//
//      Cell1 := WorkBook1.ActiveSheet.Cells[2, 16+i*3];           //
//      Cell2 := WorkBook1.ActiveSheet.Cells[2+NGroups-1, 17+i*3]; // ������
//      Range2 := WorkBook1.ActiveSheet.Range[Cell1, Cell2];       // ������
//      Chart[i].Chart.SetSourceData(Range2);                      //
//    end;

////////////////////////////////////////////////////////////////////////////////
//////////////////////////// * ���� "����� ������" * ///////////////////////////
////////////////////////////////////////////////////////////////////////////////

  if BlankWafer.NTotal > 0 then // ���� ���� ����� ��������
  begin
    try
      Workbook1.Sheets['����� ������'].Activate;
    except
      if Assigned(OnEvent) then OnEvent(evError, '� ������� ����������� ���� <����� ������>');

    end;

    if Assigned(OnEvent) then OnEvent(evInfo, '... ��� ������ � ����!');

    StX := 11; // ���������
    StY := 2;  // 1-� ��������

    if MapByParams then // ���� ����� �������� �� ����������
    begin
      WorkBook1.ActiveSheet.Cells[2, 1].Interior.Color := clRed;
      WorkBook1.ActiveSheet.Cells[2, 1].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
      WorkBook1.ActiveSheet.Cells[2, 2] := '����';
      WorkBook1.ActiveSheet.Cells[3, 1].Interior.Color := clMaroon; ////
      WorkBook1.ActiveSheet.Cells[3, 1].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
      WorkBook1.ActiveSheet.Cells[3, 2] := '�������� ���� �����';
      WorkBook1.ActiveSheet.Cells[4, 1].Interior.Color := clYellow;
      WorkBook1.ActiveSheet.Cells[4, 1].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
      WorkBook1.ActiveSheet.Cells[4, 2] := '�������� ���� �����';
      WorkBook1.ActiveSheet.Cells[5, 1].Interior.Color := clWhite;
      WorkBook1.ActiveSheet.Cells[5, 1].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
      WorkBook1.ActiveSheet.Cells[5, 2] := '�������� �� ���������';

      Str := '����������';
      case BlankWafer.CutSide of
        1: Str := '������';
        2: Str := '�����';
        3: Str := '�����';
        4: Str := '������';
      end;
      WorkBook1.ActiveSheet.Cells[7, 1] := '���� ��������: '+Str;

      WorkBook1.ActiveSheet.Columns[StX].Orientation := 90; // ��� �����������
      WorkBook1.ActiveSheet.Columns[StX].Font.Bold := True; // ������
      WorkBook1.ActiveSheet.Columns[StX].Font.Size := 12;   // ��������

      LenX := Length(BlankWafer.Chip[0])+1; // ��� �� � ����� ����������

      for n := 0 to Length(Wafer)-1 do
        with Wafer[n] do
        begin
          StX := 11;

          if NTotal > BlankWafer.NTotal then
            if Assigned(OnEvent) then OnEvent(evError, ' ��������: '+Num+' - ���������� ������, ��� � ������!');
          if NTotal < BlankWafer.NTotal then
            if Assigned(OnEvent) then OnEvent(evError, ' ��������: '+Num+' - ���������� ������, ��� � ������!');

          Cell1 := WorkBook1.ActiveSheet.Cells[StY, StX];                           // ���������
          Cell2 := WorkBook1.ActiveSheet.Cells[StY+Length(BlankWafer.Chip)-1, StX]; // ������ ���
          WorkBook1.ActiveSheet.Range[Cell1, Cell2].Merge;                          // �����������
          WorkBook1.ActiveSheet.Cells[StY, StX] := Num; // ����� ��������           // ������� �������

          //////////////////////////// * ����� ���� * //////////////////////////

          WorkBook1.ActiveSheet.Cells[1, StX+1+LenX div 2] := '����� �����';

          for Nm := 0 to NTotal-1 do
          begin
            if Nm > BlankWafer.NTotal-1 then Break;

            X := BlankWafer.ChipN[Nm].X;
            Y := BlankWafer.ChipN[Nm].Y;
            X1 := ChipN[Nm].X;
            Y1 := ChipN[Nm].Y;
            WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2] := Nm+1;
            if Chip[Y1, X1].Status = 1 then WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2].Interior.Color := clLime
                                       else WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2].Interior.Color := clRed;
            WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
          end;

          //////////////////////// * ���� �� ���������� * //////////////////////

          StX := StX+LenX;

          for Nm := 0 to NTotal-1 do
          begin
            if Nm > BlankWafer.NTotal-1 then Break;

            X := BlankWafer.ChipN[Nm].X;
            Y := BlankWafer.ChipN[Nm].Y;
            X1 := ChipN[Nm].X;
            Y1 := ChipN[Nm].Y;

            for i := 0 to Length(TestsParams)-1 do
            begin
              case Chip[Y1, X1].ChipParams[i].Stat of
                0: Col := clBlack;
                1: Col := clLime;
                2: Col := clYellow;
                3: Col := clMaroon; ////
                4: Col := clWhite;
              end;

              if n = 0 then                                                                             // ��������
                if NM = 0 then WorkBook1.ActiveSheet.Cells[StY-1, StX+7+i*LenX] := TestsParams[i].Name; // ������

              WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2+i*LenX] := Nm+1;
              WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2+i*LenX].Interior.Color := Col;
              WorkBook1.ActiveSheet.Cells[Y+StY, X+StX+2+i*LenX].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
            end;
          end;

          StY := StY+Length(BlankWafer.Chip)+2;
        end;
    end
    else               // ���� ����� �������� �� �/�
    begin
      for i := 1 to Length(Wafer[0].TestsParams)-1 do
      begin
        WorkBook1.ActiveSheet.Cells[i+1, 1].Interior.Color := GetColorByStatus(i+1999);
        WorkBook1.ActiveSheet.Cells[i+1, 1].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
        WorkBook1.ActiveSheet.Cells[i+1, 2] := '���� �� '+Wafer[0].TestsParams[i].Name;
      end;

      Str := '����������';
      case BlankWafer.CutSide of
        1: Str := '������';
        2: Str := '�����';
        3: Str := '�����';
        4: Str := '������';
      end;
      WorkBook1.ActiveSheet.Cells[Length(Wafer[0].TestsParams)+2, 1] := '���� ��������: '+Str;

      for n := 0 to Length(Wafer)-1 do
        with Wafer[n] do
        begin
          if NTotal > BlankWafer.NTotal then
            if Assigned(OnEvent) then OnEvent(evError, ' ��������: '+Num+' - ���������� ������, ��� � ������!');
          if NTotal < BlankWafer.NTotal then
            if Assigned(OnEvent) then OnEvent(evError, ' ��������: '+Num+' - ���������� ������, ��� � ������!');

          WorkBook1.ActiveSheet.Cells[StY-1, StX+5] := '��������: '+Num;
          for Nm := 0 to NTotal-1 do
          begin
            if Nm > BlankWafer.NTotal-1 then Break;

            X := BlankWafer.ChipN[Nm].X;
            Y := BlankWafer.ChipN[Nm].Y;
            X1 := ChipN[Nm].X;
            Y1 := ChipN[Nm].Y;
            WorkBook1.ActiveSheet.Cells[Y+StY, X+StX] := Nm+1;
            WorkBook1.ActiveSheet.Cells[Y+StY, X+StX].Interior.Color := GetColorByStatus(Chip[Y1, X1].Status);
            WorkBook1.ActiveSheet.Cells[Y+StY, X+StX].BorderAround(xlContinuous, xlThin, xlAutomatic, xlAutomatic);
          end;

          if ((n+1) mod 5) <> 0 then
          begin
            StX := StX+Length(BlankWafer.Chip[0])+2;
          end
          else
          begin
            StX := 11;
            StY := StY+Length(BlankWafer.Chip)+2;

            WorkBook1.ActiveSheet.Rows[StY-1].Font.Bold := True; // ��������
            WorkBook1.ActiveSheet.Rows[StY-1].Font.Size := 12;   // ������� (����. �����)
          end;
        end;
    end;
  end
  else // ���� ��� ����� ������
  try
    Workbook1.Sheets['����� ������'].Delete; // ������ ���� <����� ������> /////////////////////////////////////////////////////////////
  except
  end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

  XLSfName := ChangeFileExt(LfName, '')+'.xlsx';
  Workbook1.SaveAs(GetFreeFileName(XLSfName));

  if Assigned(OnEvent) then OnEvent(evCreate, '-------------------------------------------');
  if Assigned(OnEvent) then OnEvent(evCreate, '������ ����: '+ExtractFileName(XLSfName)); // ������ ��������� � �������� �����
                                                                                                                     //
  FormatSettings.DecimalSeparator := '.';                                                                            //
                                                                                                                     //
/////////////////////////////////                                                                                    //
                                                                                                                     //
  VarMass1 := UnAssigned;                                                                                            //
  VarMass2 := UnAssigned;                                                                                            //
  Workbook1.Close;                                                                                                   //
  if not VarIsEmpty(Excel) then                                                                                      //
  begin                                                                                                              //
    if (Excel.WorkBooks.Count > 0) and (not Excel.Visible) then                                                      //
    begin                                                                                                            //
      Excel.WindowState := Excel.xlMinimized;                                                                        //
      Excel.Visible := True;                                                                                         //
    end                                                                                                              //
    else Excel.Quit;                                                                                                 //
                                                                                                                     //
    Workbook1 := UnAssigned;                                                                                         //
    Excel     := UnAssigned;                                                                                         //
  end;                                                                                                               //
                                                                                                                     //
  Result := True;                                                                                                    //
end;                                                                                                                 //
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////
function TLot.GetColorByStatus(const Stat: WORD): TColor; //
begin                                                     //
  Result := clSilver;                                     //
                                                          //
  case Stat of                                            //
       1: Result := clLime;                               //
    2000: Result := clRed;                                //
    2001: Result := clOlive;                              //
    2002: Result := clBlue;                               //
    2003: Result := clMaroon;                             //
    2004: Result := clSkyBlue;                            //
    2005: Result := clNavy;                               //
    2006: Result := clHighlight;                          //
    2007: Result := clYellow;                             //
    2008: Result := clFuchsia;                            //
    2009: Result := clPurple;                             //
    2010: Result := clAqua;                               //
    2011: Result := clGray;                               //
  end;                                                    //
end;                                                      //
////////////////////////////////////////////////////////////


end.
