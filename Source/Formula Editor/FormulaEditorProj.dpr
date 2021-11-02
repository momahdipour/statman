program FormulaEditorProj;

uses
  Forms,
  MainFormulaUnit in 'MainFormulaUnit.pas' {MainForm},
  FmlAboutUnit in 'FmlAboutUnit.pas' {FmlAboutForm};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TFmlAboutForm, FmlAboutForm);
  Application.Run;
end.
