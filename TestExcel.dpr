program TestExcel;

uses
  Vcl.Forms,
  uMainForm in 'uMainForm.pas' {MainForm},
  Excel_TLB in 'C:\Users\mika\Documents\Embarcadero\Studio\21.0\Imports\Excel_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
