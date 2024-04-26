unit PrefDlg;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Structs;

type
  TPrefForm = class(TForm)
    MainGroup: TGroupBox;
    ToFirstFailChB: TCheckBox;
    CreateSTSChB: TCheckBox;
    NoNormsChB: TCheckBox;
    CloseBtn: TBitBtn;
    MapByParamsChB: TCheckBox;
  private
  public

  end;

var
  PrefForm: TPrefForm;

implementation

{$R *.dfm}

end.
