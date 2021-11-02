program StatManProj;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  FormulaUnit in 'FormulaUnit.pas' {FormulaForm},
  TableUnit in 'TableUnit.pas' {TableForm},
  FmlAboutUnit in 'FmlAboutUnit.pas' {FmlAboutForm},
  ChartFormUnit in 'ChartFormUnit.pas' {ChartForm},
  MainAboutUnit in 'MainAboutUnit.pas' {MainAboutBox},
  SplashScreenUnit in 'SplashScreenUnit.pas' {SplashScreen},
  SysUtils,
  Procs in 'Procs.pas',
  HelpLangFormUnit in 'HelpLangFormUnit.pas' {HelpLangForm},
  TrayAnimation in 'TrayAnimation.pas',
  Languages in 'Languages.pas',
  GridFindUnit in 'GridFindUnit.pas' {GridFindForm},
  GridReplaceUnit in 'GridReplaceUnit.pas' {GridReplaceForm},
  XLSSheetNameUnit in 'XLSSheetNameUnit.pas' {XLSSheetName},
  ChartPrevUnit in 'ChartPrevUnit.pas' {ChartPrevForm},
  RandomProducerUnit in '..\Statistics\Images\RandomProducerUnit.pas' {RandomProducerForm},
  TipUnit in 'TipUnit.pas' {TipForm},
  MemberUnit in 'MemberUnit.pas' {MemberForm};

{$R *.res}
begin
  SplashScreen:=TSplashScreen.Create(Application);
  SplashScreen.Show;
  SplashScreen.Update;
  Application.Initialize;
  Application.Title := 'StatMan';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(THelpLangForm, HelpLangForm);
  Application.CreateForm(TChartForm, ChartForm);
  Application.CreateForm(TFormulaForm, FormulaForm);
  Application.CreateForm(TTableForm, TableForm);
  Application.CreateForm(TFmlAboutForm, FmlAboutForm);
  Application.CreateForm(TMainAboutBox, MainAboutBox);
  Application.CreateForm(TGridFindForm, GridFindForm);
  Application.CreateForm(TGridReplaceForm, GridReplaceForm);
  Application.CreateForm(TXLSSheetName, XLSSheetName);
  Application.CreateForm(TChartPrevForm, ChartPrevForm);
  Application.CreateForm(TRandomProducerForm, RandomProducerForm);
  Application.CreateForm(TTipForm, TipForm);
  Application.CreateForm(TMemberForm, MemberForm);
  Sleep(600);
  SplashScreen.Hide;
  SplashScreen.Free;
  Application.Run;
end.
