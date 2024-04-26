program Magic2;

uses
  Forms,
  uMagic2 in 'uMagic2.pas' {MDBForm},
  Structs in 'Structs.pas',
  Statistica_le in 'Statistica_le.pas',
  PrefDlg in 'PrefDlg.pas' {PrefForm};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMDBForm, MDBForm);
  Application.Run;
end.
