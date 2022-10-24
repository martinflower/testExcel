program TestExcel;

uses
  Vcl.Forms,
  uMainForm in 'uMainForm.pas' {MainForm},
  Excel_TLB in 'import\Excel_TLB.pas',
  VBIDE_TLB in 'import\VBIDE_TLB.pas',
  Office_TLB in 'import\Office_TLB.pas';



{$R *.res}

begin
  Vcl.Forms.Application.Initialize;
  Vcl.Forms.Application.MainFormOnTaskbar := True;
  Vcl.Forms.Application.CreateForm(TMainForm, MainForm);
  Vcl.Forms.Application.Run;
end.
