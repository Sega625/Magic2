unit Structs;

interface

uses
  Windows, Forms, System.Classes, SysUtils, ExtCtrls, Controls, Registry;

const
  xlAutomatic = $FFFFEFF7;
  xlContinuous = $00000001;
  xlHairline = 1; 	  // ����� ������ �������
  xlMedium   = -4138; //	������� �������
  xlThick    = 4;     //	������� �������
  xlThin     = 2;     //	������ �������


  NotSpec: Real=987654321.0;

  csTop    = 1;
  csLeft   = 2;
  csBottom = 3;
  csRight  = 4;


  dURightToLeft = 0;
  dURightSnake  = 1;
  dULeftToRight = 2;
  dULeftSnake   = 3;

  dLUpToDown    = 4;
  dLUpSnake     = 5;
  dLDownToUp    = 6;
  dLDownSnake   = 7;

  dDLeftToRight = 8;  //
  dDLeftSnake   = 9;  //
  dDRightToLeft = 10; //
  dDRightSnake  = 11; //

  dRDownToUp    = 12;
  dRDownSnake   = 13;
  dRUpToDown    = 14;
  dRUpSnake     = 15;

{
  ConnStr = 'Jet OLEDB:System database="";'+
            'Jet OLEDB:Registry Path="";'+
            'Jet OLEDB:Database Password="";'+
            'Jet OLEDB:Engine Type=5;'+
            'Jet OLEDB:Database Locking Mode=1;'+
            'Jet OLEDB:Global Partial Bulk Ops=2;'+
            'Jet OLEDB:Global Bulk Transactions=1;'+
            'Jet OLEDB:New Database Password="";'+
            'Jet OLEDB:Create System Database=False;'+
            'Jet OLEDB:Encrypt Database=False;'+
            'Jet OLEDB:Don''t Copy Locale on Compact=False;'+
            'Jet OLEDB:Compact Without Replica Repair=False;'+
            'Jet OLEDB:SFP=False;'+
            'Provider=Microsoft.Jet.OLEDB.4.0;'+
            'User ID=Admin;'+
            'Mode=Share Deny None;'+
            'Extended Properties="";';
}
  ConnStr = 'Provider=Microsoft.Jet.OLEDB.4.0;'+
            'Jet OLEDB:Create System Database=true;'+
            'Data Source=';

type
  TEventType = (evError, evOK, evInfo, evSave, evCreate);

  TDynArray = array of Real;

  TCadre = record
    StartX: WORD;
    StartY: WORD;
    ScaleX: WORD;
    ScaleY: WORD;
  end;

  TZone2 = record
    Y1: WORD;
    Y2: WORD;
    X3: WORD;
    X4: WORD;
  end;
  TZones2 = array of TZone2;

/////////////////////////////////////////

  TChipParams = record
    Value: Real;
    Stat : byte; // 0-?; 1-������; 2-���� �����; 3-���� �����; 4-�� ���������
  end;

  TChip = record
    ID        : DWORD;
    Status    : WORD;
    ChipParams: array of TChipParams;
  end;
  PChip = ^TChip;
  TChips = array of array of TChip;

/////////////////////////////////////////

  TChipsN = array of TPoint; // ���������� ����� � ������� ���������

/////////////////////////////////////////

  TFail = record
    Status  : WORD;
    Name    : string;
    Quantity: DWORD;
//    Col     : TColor;
  end;
  TFails = array of TFail;

/////////////////////////////////////////

  TNorma = record
    Min: Real;
    Max: Real;
  end;

  TTestParams = record
    Name  : string;
    Norma : TNorma;
    PUnit : string;
    PMode : string;
    Status: WORD;
  end;
  TTestsParams = array of TTestParams;

/////////////////////////////////////////

  TCalcParams = record
    Asum: Real;       // ����� ��� ��������
    Qsum: Real;       // ����� ��������� ��� ���. ����������
    ValMass: TDynArray; // ������ ���������� ��� ������� � ���������
    ValCount: DWORD;    // ���-�� ����������

    NOKVal   : DWORD;
    NFailsVal: DWORD;
    
    AvrVal : Real; // �������
    MinVal : Real; // �����������
    MaxVal : Real; // ������������
    StdVal : Real; // ���. ����������
    MedVal : Real; // �������
    Qrt1Val: Real; // 1-� ��������
    Qrt3Val: Real; // 3-� ��������
  end;
  TCalcsParams = array of TCalcParams;

/////////////////////////////////////////

  THistGroup = record //
    Name: string;     // ��� ����������
    Num: DWORD;       // �� �����
  end;                //

  TGroup = record
    Min: Real;
    Max: Real;
    Num: DWORD;
  end;

  THistParams = record
    Name: string;
    ShortName: string;
//    AllValMass: TDynArray; // ������ ���������� ��� ����������
//    AllMinVal : Real;    // ����������� ��������
//    AllMaxVal : Real;    // ������������ ��������

    NGroups: DWORD; // ���-�� �����
    Group: array of TGroup; // ������
  end;

/////////////////////////////////////////

  procedure SortMassByValue(var Mass: TDynArray);
  function  GetExcelAppName2(): string;
  function  GetFreeFileName(fName: TFileName): TFileName;
  function  IsChip    (const Status: WORD): Boolean;
  function  IsFailChip(const Status: WORD): Boolean;
  function  EqualStatus(const Status1, Status2 : WORD): Boolean;
  procedure ErrMess(Handle: THandle; const ErrMes: String);
  function  QuestMess(Handle: THandle; const QStr: String): Integer;
  procedure InfoMess(Handle: THandle; const IStr: string);

implementation

//////////////////////////////////////////////////
procedure SortMassByValue(var Mass: TDynArray); //
var                                             //
  n, m, b_m: DWORD;                             //
  b_val, tmpVal: Real;                          //
begin                                           //
  if Length(Mass) < 2 then Exit;                //
                                                //
  for n := 0 to Length(Mass)-2 do               //
  begin                                         //
    b_val := Mass[n];                           //
    b_m := n;                                   //
    for m := n+1 to Length(Mass)-1 do           //
      if Mass[m] < b_val then                   //
      begin                                     //
        b_val := Mass[m];                       //
        b_m := m;                               //
      end;                                      //
    tmpVal    := Mass[b_m];                     //
    Mass[b_m] := Mass[n];                       //
    Mass[n]   := tmpVal;                        //
  end;                                          //
end;                                            //
//////////////////////////////////////////////////

/////////////////////////////////////////////////////
function GetExcelAppName2(): string;               //
var                                                //
  reg: TRegistry;                                  //
  SL: TStringList;                                 //
  n: DWORD;                                        //
begin                                              //
  Result := '';                                    //
                                                   //
  reg := TRegistry.Create;                         //
  reg.RootKey := HKEY_CLASSES_ROOT;                //
  try                                              //
    if reg.OpenKeyReadOnly('') then                //
    begin                                          //
      SL := TStringList.Create();                  //
      reg.GetKeyNames(SL);                         //
      reg.CloseKey;                                //
    end                                            //
  finally                                          //
    reg.Free;                                      //
  end;                                             //
                                                   //
  if SL.Count > 0 then                             //
    for n := 0 to SL.Count-1 do                    //
      if Pos('Excel.Application', SL[n]) <> 0 then //
      begin                                        //
        Result := SL[n];                           //
        Break;                                     //
      end;                                         //
  SL.Free();                                       //
end;                                               //
/////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////
function GetFreeFileName(fName: TFileName): TFileName;      //
var                                                         //
  n, P: WORD;                                               //
  tPath, tfName, tExt: TFileName;                           //
begin                                                       //
  tPath  := ExtractFilePath(fName);                         //
  Result := ExtractFileName(fName);                         //
  tExt   := ExtractFileExt(fName);                          //
  Result := ChangeFileExt(Result, '');                      //
                                                            //
  n := 0;                                                   //
  while FileExists(tPath+Result+tExt) do                    //
  begin                                                     //
    P := LastDelimiter('(', Result);                        //
    if P <> 0 then Delete(Result, P, (Length(Result)-P)+1); //
    Inc(n);                                                 //
    Result := Result+'('+IntToStr(n)+')';                   //
  end;                                                      //
                                                            //
  Result := tPath+Result+tExt;                              //
end;                                                        //
//////////////////////////////////////////////////////////////

////////////////////////////////////////////////////
function IsChip(const Status: WORD): Boolean;     //
begin                                             //
  Result := False;                                //
                                                  //
  case Status of                                  //
    0         : Result := True;                   //
    1         : Result := True;                   //
    2         : ;                                 //
    3         : ;                                 //
    4         : ;//Result := True;                   //
    5         : ;                                 //
    7         : Result := True;                   //
    10..1500  : Result := True;                   //
    2000..3000: Result := True;                   //
    3500..4500: Result := True;                   //
  end;                                            //
end;                                              //
////////////////////////////////////////////////////
////////////////////////////////////////////////////
function IsFailChip(const Status: WORD): Boolean; //
begin                                             //
  Result := False;                                //
                                                  //
  case Status of                                  //
    0         : ;                                 //
    1         : ;                                 //
    2         : ;                                 //
    3         : ;                                 //
    4         : ;                                 //
    5         : ;                                 //
    7         : ;                                 //
    10..1500  : Result := True;                   //
    2000..3000: Result := True;                   //
    3500..4500: Result := True;                   //
  end;                                            //
end;                                              //
////////////////////////////////////////////////////
{
/////////////////////////////////////////////////////
function GetMainColor(const Status: WORD): TColor; //
begin                                              //
  case Status of                                   //
    0         : Result := clNotTested;             //
    1         : Result := clOK;                    //
    2         : Result := clNotChip;               //
    3         : Result := clNot4Testing;           //
    4         : Result := cl4Mark;                 //
    5         : Result := clRepper;                //
    7         : Result := clMRepper;               //
    10..1500  : Result := clFailNC;                //
    2000..3000: Result := clFailSC;                //
    3500..4500: Result := clFailFC;                //
  else          Result := clGray;                  //
  end;                                             //
end;                                               //
/////////////////////////////////////////////////////
}
//////////////////////////////////////////////////////////////////////////
function EqualStatus(const Status1, Status2 : WORD): Boolean;           //
begin                                                                   //
  Result := False;                                                      //
                                                                        //
  if Status1 = Status2 then Result := True                              //
  else                                                                  //
    if IsFailChip(Status1) and IsFailChip(Status1) then Result := True; //
end;                                                                    //
//////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////
procedure ErrMess(Handle: THandle; const ErrMes: String);                                //
begin                                                                                    //
  MessageBox(Handle, PChar(ErrMes), '������!!!', MB_ICONERROR+MB_OK);                    //
end;                                                                                     //
///////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////
function QuestMess(Handle: THandle; const QStr: String): Integer;                        //
begin                                                                                    //
  Result := MessageBox(Handle, PChar(Qstr), '�������������!', MB_ICONQUESTION+MB_YESNO); //
end;                                                                                     //
///////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////
procedure InfoMess(Handle: THandle; const IStr: string);                                 //
begin                                                                                    //
  MessageBox(Handle, PChar(IStr), '��������!!!', MB_ICONINFORMATION+MB_OK);              //
end;                                                                                     //
///////////////////////////////////////////////////////////////////////////////////////////


end.
