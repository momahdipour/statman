unit MainUnit;

interface

uses
  Windows, Messages, SysUtils,  Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ExtCtrls, Grids, ComCtrls, Buttons, StdCtrls,
  Spin, Math, AppEvnts, TeeProcs, TeEngine, Chart,
  Series, xpCheckBox, xpButton, ColorSelector, ExplBtn,
  ExTrackBar, penwidthcombo, penstylecombo, ColorButtons,
  Mask, Parser10,  StrUtils, CoolTrayIcon, TextTrayIcon,
  ToolWin, TeCanvas, ImgList, TrayAnimation,  BaseGrid, AdvGrid,
  Languages, asgprint, XPMenu, XLSSheetNameUnit, Clipbrd, RandomProducerUnit,
  AsgReplaceDialog, ExtDlgs, OfficeImageList;


const
  MaxFieldLen=20;
  MaxDesLen=120;
  MaxStringLen=50;

type
  TCharStyle=(csNone,csUppercase,csLowercase);
  SheetSettings=record
    Used:Boolean;
    FieldName:string[MaxFieldLen];
    DesIndex:Byte;
    ForceValue:Boolean;
    VListIndex:Byte;
    Filter:Boolean;
    FilterIndex:Byte;
    case TypeIndex:Byte of
      1:(DValue1:Integer;);
      2:(DValue2:Extended;);
      3:(DValue3:String[50];MaxStringLength:Byte;CharStyle:TCharStyle);
      4:(LBoundI,UBoundI:Integer;);
      5:(LBoundE,UBoundE:Extended;);
    end;
  DefaultValueType=record
    case TypeIndex:Byte of
      1:(DValue1:Integer;);
      2:(DValue2:Extended;);
      3:(DValue3:String[50];);
      4:(LBoundI,UBoundI:Integer;);
      5:(LBoundE,UBoundE:Extended;);
    end;
  WorkSettings=record
    Sheet:SheetSettings;
    ForceCount:Integer;
    ValueCount:Integer;
    StrTempCount:Integer;
    SpanCount:Integer;
    IntCount,DecCount,StrCount:Integer;
  end;
  FilteredCell=record
    Str:String;
    Row,Col:Integer;
  end;
  TLockedCell=record
    Row,Col:Integer;
  end;

  {$I ChartSettings.Inc}
  WorkFile=File of WorkSettings;
  TSortingMode=(smAscending,smDescending);
  TGridResizeKind=(grkGreatest,grkSmallest);

  TMainForm = class(TForm)
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    View1: TMenuItem;
    MemberPane1: TMenuItem;
    StatusBar1: TMenuItem;
    ListView1: TMenuItem;
    IconView1: TMenuItem;
    ToolBarView1: TMenuItem;
    MenuView1: TMenuItem;
    Help1: TMenuItem;
    About1: TMenuItem;
    MemberGroup: TGroupBox;
    Shape1: TShape;
    MemberPaneImage: TImage;
    MemberList: TPanel;
    SettingPanel: TPanel;
    SettingGrid: TStringGrid;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    GroupBox2: TGroupBox;
    HelpLabel: TLabel;
    AlertImage: TImage;
    List1: TListBox;
    ForceValue: TCheckBox;
    UseFilter: TCheckBox;
    FxLabel: TStaticText;
    StaticText2: TStaticText;
    SettingTitle: TStaticText;
    MemberButtonNormal: TImage;
    MemberButtonPressed: TImage;
    SettingSection: TSpeedButton;
    Sheet1Section: TSpeedButton;
    Sheet2Section: TSpeedButton;
    Sheet3Section: TSpeedButton;
    Sheet4Section: TSpeedButton;
    Sheet5Section: TSpeedButton;
    TableSection: TSpeedButton;
    AnalyzeSection: TSpeedButton;
    ChartSection: TSpeedButton;
    StrList1: TListBox;
    DecList4: TListBox;
    DecList5: TListBox;
    SpanList: TListBox;
    IntList1: TListBox;
    StrList5: TListBox;
    StrList4: TListBox;
    StrList3: TListBox;
    StrList2: TListBox;
    IntList2: TListBox;
    IntList3: TListBox;
    IntList4: TListBox;
    IntList5: TListBox;
    DecList1: TListBox;
    DecList2: TListBox;
    DecList3: TListBox;
    StrTemp: TListBox;
    MemberIcon: TPanel;
    MemberToolbar: TPanel;
    IconSettings: TSpeedButton;
    IconSheet1: TSpeedButton;
    IconSheet2: TSpeedButton;
    IconSheet3: TSpeedButton;
    IconSheet4: TSpeedButton;
    IconSheet5: TSpeedButton;
    IconTable: TSpeedButton;
    IconChart: TSpeedButton;
    IconAnalyze: TSpeedButton;
    SpeedButton4: TSpeedButton;
    AlertTimer: TTimer;
    ToolSettings: TSpeedButton;
    ToolSheet1: TSpeedButton;
    ToolSheet2: TSpeedButton;
    ToolSheet3: TSpeedButton;
    ToolSheet4: TSpeedButton;
    ToolSheet5: TSpeedButton;
    ToolTable: TSpeedButton;
    ToolChart: TSpeedButton;
    ToolAnalyze: TSpeedButton;
    ViewPopup: TPopupMenu;
    ListView: TMenuItem;
    IconView: TMenuItem;
    ToolView: TMenuItem;
    MenuView: TMenuItem;
    MemberMenu: TPopupMenu;
    MSettingSection: TMenuItem;
    MSheet1Section: TMenuItem;
    MSheet2Section: TMenuItem;
    MSheet3Section: TMenuItem;
    MSheet4Section: TMenuItem;
    MSheet5Section: TMenuItem;
    MTableSection: TMenuItem;
    MChartSection: TMenuItem;
    MAnalyzeSection: TMenuItem;
    PaneMenuBtn: TImage;
    S1Check: TCheckBox;
    S2Check: TCheckBox;
    S3Check: TCheckBox;
    S4Check: TCheckBox;
    S5Check: TCheckBox;
    TypeCombo: TComboBox;
    TempGrid: TStringGrid;
    Sheet2Panel: TPanel;
    Sheet1Panel: TPanel;
    S1NLabel22: TStaticText;
    S1IDEdit: TStaticText;
    GroupBox7: TGroupBox;
    Sheet1Title: TStaticText;
    S1FieldLabel: TStaticText;
    Type1Title: TStaticText;
    S1Filterbtn: TBitBtn;
    Des1Title: TStaticText;
    Filter1Title: TStaticText;
    S1DesLabel: TStaticText;
    S1FilterLabel: TStaticText;
    S1TypeLabel: TStaticText;
    S1NLabel: TStaticText;
    S1Combo: TComboBox;
    S2Combo: TComboBox;
    S2IDEdit: TStaticText;
    GroupBox3: TGroupBox;
    Sheet2Title: TStaticText;
    S2FieldLabel: TStaticText;
    Type2Title: TStaticText;
    S2Filterbtn: TBitBtn;
    Des2Title: TStaticText;
    Filter2Title: TStaticText;
    S2DesLabel: TStaticText;
    S2FilterLabel: TStaticText;
    S2TypeLabel: TStaticText;
    StaticText1: TStaticText;
    S2NLabel: TStaticText;
    Sheet3Panel: TPanel;
    GroupBox4: TGroupBox;
    Sheet3Title: TStaticText;
    S3FieldLabel: TStaticText;
    Type3Title: TStaticText;
    S3Filterbtn: TBitBtn;
    Des3Title: TStaticText;
    Filter3Title: TStaticText;
    S3DesLabel: TStaticText;
    S3FilterLabel: TStaticText;
    S3TypeLabel: TStaticText;
    S3Combo: TComboBox;
    S3IDEdit: TStaticText;
    StaticText17: TStaticText;
    S3NLabel: TStaticText;
    Sheet4Panel: TPanel;
    GroupBox5: TGroupBox;
    Sheet4Title: TStaticText;
    S4FieldLabel: TStaticText;
    Type4Title: TStaticText;
    S4Filterbtn: TBitBtn;
    Des4Title: TStaticText;
    Filter4Title: TStaticText;
    S4DesLabel: TStaticText;
    S4FilterLabel: TStaticText;
    S4TypeLabel: TStaticText;
    S4IDEdit: TStaticText;
    S4Combo: TComboBox;
    StaticText22: TStaticText;
    S4NLabel: TStaticText;
    Sheet5Panel: TPanel;
    GroupBox6: TGroupBox;
    Sheet5Title: TStaticText;
    S5FieldLabel: TStaticText;
    Type5Title: TStaticText;
    S5Filterbtn: TBitBtn;
    Des5Title: TStaticText;
    Filter5Title: TStaticText;
    S5DesLabel: TStaticText;
    S5FilterLabel: TStaticText;
    S5TypeLabel: TStaticText;
    S5Combo: TComboBox;
    S5IDEdit: TStaticText;
    StaticText27: TStaticText;
    StaticText28: TStaticText;
    S5NLabel: TStaticText;
    N5: TMenuItem;
    N6: TMenuItem;
    TablePanel: TPanel;
    TableGrid: TStringGrid;
    HeadLabel: TStaticText;
    FootLabel: TStaticText;
    PageControl2: TPageControl;
    TabSheet2: TTabSheet;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    Label1: TLabel;
    SpinEdit1: TSpinEdit;
    Label2: TLabel;
    Edit2: TEdit;
    Label3: TLabel;
    Edit3: TEdit;
    CreateTable2: TBitBtn;
    CheckBox5: TCheckBox;
    FrequencyTableTitle: TStaticText;
    ApplicationEvents1: TApplicationEvents;
    AnalyzePanel: TPanel;
    GroupBox1: TGroupBox;
    GroupBox8: TGroupBox;
    GroupBox9: TGroupBox;
    GroupBox10: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    anaDesLabel: TLabel;
    Image1: TImage;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Image2: TImage;
    Label23: TLabel;
    DataAnalyzingTitle: TStaticText;
    MeanLabel: TLabel;
    MiddleLabel: TLabel;
    ModeLabel: TLabel;
    ADLabel: TLabel;
    VarLabel: TLabel;
    SDLabel: TLabel;
    CVLabel: TLabel;
    anaMinLabel: TLabel;
    anaMaxLabel: TLabel;
    anaNLabel: TLabel;
    anaRangeLabel: TLabel;
    ModeList: TListBox;
    ModeListTimer: TTimer;
    SpeedButton1: TSpeedButton;
    ModeBtnDown: TImage;
    ModeBtnUp: TImage;
    ChartPanel: TPanel;
    ChartPageControl1: TPageControl;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    WallSheet: TTabSheet;
    AxisSheet: TTabSheet;
    Chart1: TChart;
    GroupBox11: TGroupBox;
    GroupBox12: TGroupBox;
    Label16: TLabel;
    Panel1: TPanel;
    ChartLabel: TLabel;
    ChartsTitle: TStaticText;
    RadioButton1: TRadioButton;
    UseColors: TRadioButton;
    ColorOptions: TGroupBox;
    bf: TxpCheckBox;
    cg: TxpCheckBox;
    UseBackImage: TxpCheckBox;
    bc: TxpCheckBox;
    ColorD1: TColorDialog;
    GLabel: TLabel;
    StartColor: TColorBox;
    EndColor: TColorBox;
    backimage: TImage;
    Style: TGroupBox;
    PutInside: TxpCheckBox;
    BrowseImage: TxpButton;
    TileRadio: TRadioButton;
    StretchRadio: TRadioButton;
    CenterRadio: TRadioButton;
    BackColor: TColorSelector;
    FrameColor: TColorSelector;
    xpCheckBox6: TxpCheckBox;
    View3D: TGroupBox;
    Chart3DTrack: TExTrackBar;
    Label26: TLabel;
    Label27: TLabel;
    GroupBox16: TGroupBox;
    ZoomTrack: TExTrackBar;
    NormalView: TRadioButton;
    CustomizedView: TRadioButton;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    PerspectiveTrack: TExTrackBar;
    ZR: TExTrackBar;
    YR: TExTrackBar;
    XR: TExTrackBar;
    Image4: TImage;
    xpCheckBox7: TxpCheckBox;
    LegendOptions: TPageControl;
    TabSheet10: TTabSheet;
    TabSheet11: TTabSheet;
    GroupBox17: TGroupBox;
    GroupBox18: TGroupBox;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    LegendPos: TComboBox;
    LegendStyle: TComboBox;
    ColorSelector3: TColorSelector;
    Legendfont: TStaticText;
    xpButton2: TxpButton;
    xpCheckBox8: TxpCheckBox;
    xpCheckBox9: TxpCheckBox;
    GroupBox19: TGroupBox;
    xpCheckBox10: TxpCheckBox;
    GroupBox20: TGroupBox;
    GroupBox21: TGroupBox;
    Label37: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    xpCheckBox11: TxpCheckBox;
    Label43: TLabel;
    lcolor: TColorSelector;
    ShadowColor: TColorSelector;
    HorizMarg: TSpinEdit;
    VertMarg: TSpinEdit;
    ShadowSize: TSpinEdit;
    View3DCheck: TxpCheckBox;
    LeftWall: TGroupBox;
    Label45: TLabel;
    Label46: TLabel;
    GroupBox23: TGroupBox;
    Label44: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    SpinEdit5: TSpinEdit;
    xpCheckBox13: TxpCheckBox;
    ColorSelector6: TColorSelector;
    ColorSelector7: TColorSelector;
    penwidthcombo3: Tpenwidthcombo;
    penstylecombo2: Tpenstylecombo;
    BackWall: TGroupBox;
    Label49: TLabel;
    Label50: TLabel;
    GroupBox25: TGroupBox;
    Label51: TLabel;
    Label52: TLabel;
    Label53: TLabel;
    xpCheckBox14: TxpCheckBox;
    ColorSelector8: TColorSelector;
    penwidthcombo4: Tpenwidthcombo;
    penstylecombo3: Tpenstylecombo;
    SpinEdit6: TSpinEdit;
    ColorSelector9: TColorSelector;
    PageControl5: TPageControl;
    TabSheet12: TTabSheet;
    TabSheet13: TTabSheet;
    xpCheckBox15: TxpCheckBox;
    ResizeT: TxpCheckBox;
    Label54: TLabel;
    GroupBox26: TGroupBox;
    TitleText: TGroupBox;
    Label55: TLabel;
    Label56: TLabel;
    Label57: TLabel;
    xpCheckBox16: TxpCheckBox;
    bwidth: Tpenwidthcombo;
    bstyle: Tpenstylecombo;
    bColor: TColorSelector;
    ColorT: TColorSelector;
    ToolBar3: TToolBar;
    ColorBtn1: TColorButton;
    UnderBtn1: TSpeedButton;
    BoldBtn1: TSpeedButton;
    ItalicBtn1: TSpeedButton;
    RightAl: TSpeedButton;
    LeftAl: TSpeedButton;
    CenterAl: TSpeedButton;
    SpeedButton9: TSpeedButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    SizeCombo: TComboBox;
    ChartTitleEdit: TMemo;
    FD1: TFontDialog;
    xpCheckBox17: TxpCheckBox;
    ResizeF: TxpCheckBox;
    Label58: TLabel;
    ColorF: TColorSelector;
    GroupBox28: TGroupBox;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    xpCheckBox20: TxpCheckBox;
    bwidthf: Tpenwidthcombo;
    bstylef: Tpenstylecombo;
    bcolorf: TColorSelector;
    FooterText: TGroupBox;
    ChartFooterEdit: TMemo;
    ToolBar4: TToolBar;
    BoldBtn2: TSpeedButton;
    ItalicBtn2: TSpeedButton;
    UnderBtn2: TSpeedButton;
    ToolButton4: TToolButton;
    LeftAl2: TSpeedButton;
    CenterAl2: TSpeedButton;
    RightAl2: TSpeedButton;
    ToolButton5: TToolButton;
    SizeCombo2: TComboBox;
    ColorBtn2: TColorButton;
    ToolButton6: TToolButton;
    SpeedButton10: TSpeedButton;
    LeftAxis: TPageControl;
    TabSheet14: TTabSheet;
    TabSheet15: TTabSheet;
    TabSheet16: TTabSheet;
    TabSheet17: TTabSheet;
    RadioButton8: TRadioButton;
    RadioButton9: TRadioButton;
    Automatic: TGroupBox;
    Label62: TLabel;
    Label63: TLabel;
    Label64: TLabel;
    GroupBox31: TGroupBox;
    Label65: TLabel;
    Label66: TLabel;
    Label67: TLabel;
    xpCheckBox21: TxpCheckBox;
    minspin: TSpinEdit;
    MaxSpin: TSpinEdit;
    Increment: TSpinEdit;
    SpinEdit10: TSpinEdit;
    SpinEdit11: TSpinEdit;
    SpinEdit12: TSpinEdit;
    xpCheckBox22: TxpCheckBox;
    GroupBox32: TGroupBox;
    GroupBox33: TGroupBox;
    xpCheckBox23: TxpCheckBox;
    xpCheckBox24: TxpCheckBox;
    mmlabel: TxpCheckBox;
    Label68: TLabel;
    Label69: TLabel;
    Label70: TLabel;
    labelcustom: TxpButton;
    mmtrack: TExTrackBar;
    LabelsFont: TStaticText;
    AxisTitle: TGroupBox;
    AxisTitleEdit: TMemo;
    ToolBar5: TToolBar;
    BoldBtn3: TSpeedButton;
    ItalicBtn3: TSpeedButton;
    UnderBtn3: TSpeedButton;
    ToolButton8: TToolButton;
    SizeCombo3: TComboBox;
    ColorBtn3: TColorButton;
    ToolButton9: TToolButton;
    SpeedButton11: TSpeedButton;
    GroupBox36: TGroupBox;
    Label75: TLabel;
    miCount: TSpinEdit;
    GroupBox37: TGroupBox;
    Label88: TLabel;
    Label89: TLabel;
    Label90: TLabel;
    xpCheckBox29: TxpCheckBox;
    axcolor: TColorSelector;
    axstyle: Tpenstylecombo;
    axwidth: Tpenwidthcombo;
    ChartToolbar: TToolBar;
    ExplorerButton9: TExplorerButton;
    ToolButton10: TToolButton;
    ChartBtn: TExplorerButton;
    SaveAsPopup: TPopupMenu;
    SaveAsMetafile1: TMenuItem;
    PrintBtn: TSpeedButton;
    ExplorerPopup1: TExplorerPopup;
    DetachBtn: TSpeedButton;
    ToolButton21: TToolButton;
    Atachbmp: TImage;
    Detachbmp: TImage;
    LegendWidth: TSpinEdit;
    FD2: TFontDialog;
    xpCheckBox27: TxpCheckBox;
    miLength: TSpinEdit;
    Label76: TLabel;
    Label77: TLabel;
    miColor: TColorSelector;
    Label78: TLabel;
    miWidth: Tpenwidthcombo;
    Label83: TLabel;
    miStyle: Tpenstylecombo;
    PageControl3: TPageControl;
    TabSheet37: TTabSheet;
    Label72: TLabel;
    Label80: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    Label82: TLabel;
    xpCheckBox26: TxpCheckBox;
    MLength: TSpinEdit;
    MColor: TColorSelector;
    MWidth: Tpenwidthcombo;
    MStyle: Tpenstylecombo;
    TabSheet38: TTabSheet;
    Label81: TLabel;
    Label84: TLabel;
    Label85: TLabel;
    Label86: TLabel;
    Label87: TLabel;
    xpCheckBox28: TxpCheckBox;
    ILength: TSpinEdit;
    IColor: TColorSelector;
    IWidth: Tpenwidthcombo;
    IStyle: Tpenstylecombo;
    Series1: TLineSeries;
    Series2: TBarSeries;
    Series3: THorizBarSeries;
    Series4: TAreaSeries;
    Series5: TPointSeries;
    Series6: TPieSeries;
    Series7: TFastLineSeries;
    StaticText25: TStaticText;
    TDefBorder: TStaticText;
    StaticText26: TStaticText;
    StaticText30: TStaticText;
    StaticText31: TStaticText;
    Label24: TLabel;
    TitleSize: TSpinEdit;
    PageControl4: TPageControl;
    TabSheet8: TTabSheet;
    xpCheckBox30: TxpCheckBox;
    Label91: TLabel;
    gridcolor: TColorSelector;
    Label92: TLabel;
    gridwidth: Tpenwidthcombo;
    Label93: TLabel;
    gridstyle: Tpenstylecombo;
    TabSheet9: TTabSheet;
    xpCheckBox1: TxpCheckBox;
    Label25: TLabel;
    gridcolorv: TColorSelector;
    Label71: TLabel;
    gridwidthv: Tpenwidthcombo;
    Label79: TLabel;
    gridstylev: Tpenstylecombo;
    PSD1: TPrinterSetupDialog;
    CopyPopup: TPopupMenu;
    CopyAsBitmap1: TMenuItem;
    CopyAsMetafile1: TMenuItem;
    ExplorerButton8: TExplorerButton;
    AnalyzeReportbtn: TExplorerButton;
    ReportPopup: TExplorerPopup;
    ReportEdit: TRichEdit;
    xpButton1: TxpButton;
    ToolBar7: TToolBar;
    RBold: TSpeedButton;
    RItalic: TSpeedButton;
    RUnder: TSpeedButton;
    ToolButton7: TToolButton;
    RLeft: TSpeedButton;
    RCenter: TSpeedButton;
    RRight: TSpeedButton;
    ToolButton11: TToolButton;
    RSizeCombo: TComboBox;
    RColorBtn: TColorButton;
    ToolButton12: TToolButton;
    RBuild: TSpeedButton;
    FD3: TFontDialog;
    xpButton3: TxpButton;
    ExplorerButton11: TExplorerButton;
    SaveGridPopup: TPopupMenu;
    Load1: TMenuItem;
    N7: TMenuItem;
    Save1: TMenuItem;
    OpenGrid: TOpenDialog;
    SaveGrid: TSaveDialog;
    ExplorerButton12: TExplorerButton;
    ExplorerButton13: TExplorerButton;
    ExplorerButton14: TExplorerButton;
    ExplorerButton15: TExplorerButton;
    FilterParser: TParser;
    Precision4Title: TLabel;
    S4PSpin: TSpinEdit;
    Precision1Title: TLabel;
    S1PSpin: TSpinEdit;
    Precision3Title: TLabel;
    S3PSpin: TSpinEdit;
    Precision5Title: TLabel;
    S5PSpin: TSpinEdit;
    Precision2Title: TLabel;
    S2PSpin: TSpinEdit;
    DVCheck: TCheckBox;
    Modebtn: TSpeedButton;
    SD1: TSaveDialog;
    S4Clrbtn: TSpeedButton;
    S1Clrbtn: TSpeedButton;
    S3Clrbtn: TSpeedButton;
    S5Clrbtn: TSpeedButton;
    S2Clrbtn: TSpeedButton;
    Fml: TImage;
    CVFml: TImage;
    S2Fml: TImage;
    SFml: TImage;
    ADFml: TImage;
    LineBtn: TSpeedButton;
    BarBtn: TSpeedButton;
    HBarBtn: TSpeedButton;
    AreaBtn: TSpeedButton;
    PointBtn: TSpeedButton;
    PieBtn: TSpeedButton;
    FastBtn: TSpeedButton;
    ChartPopup: TPopupMenu;
    Print1: TMenuItem;
    N8: TMenuItem;
    SaveAsPicture2: TMenuItem;
    SaveAsMetafile2: TMenuItem;
    N9: TMenuItem;
    CopyAsPicture1: TMenuItem;
    CopyAsMetafile2: TMenuItem;
    ToolButton13: TToolButton;
    iwidth2: Tpenwidthcombo;
    istyle2: Tpenstylecombo;
    GridPopup1: TPopupMenu;
    G1Sort: TMenuItem;
    GridPopup5: TPopupMenu;
    G5Sort: TMenuItem;
    GridPopup4: TPopupMenu;
    G4Sort: TMenuItem;
    GridPopup3: TPopupMenu;
    G3Sort: TMenuItem;
    GridPopup2: TPopupMenu;
    G2Sort: TMenuItem;
    StaticText32: TStaticText;
    StaticText33: TStaticText;
    StaticText34: TStaticText;
    StaticText35: TStaticText;
    StaticText36: TStaticText;
    Chart3: TMenuItem;
    c3: TMenuItem;
    c4: TMenuItem;
    N10: TMenuItem;
    c5: TMenuItem;
    c6: TMenuItem;
    N11: TMenuItem;
    c1: TMenuItem;
    StaticText37: TStaticText;
    Toolbars: TMenuItem;
    ChartToolbar1: TMenuItem;
    ResizeToDefault1: TMenuItem;
    ResizeToDefault2: TMenuItem;
    ResizeToDefault3: TMenuItem;
    ResizeToDefault4: TMenuItem;
    ResizeToDefault5: TMenuItem;
    StatManHelp1: TMenuItem;
    N1: TMenuItem;
    Default1: TMenuItem;
    N2: TMenuItem;
    GreatestColumnWidth1: TMenuItem;
    Smallest1: TMenuItem;
    GreatestColumnWidth2: TMenuItem;
    N4: TMenuItem;
    Default2: TMenuItem;
    N12: TMenuItem;
    Default3: TMenuItem;
    N13: TMenuItem;
    Default4: TMenuItem;
    N14: TMenuItem;
    Default5: TMenuItem;
    GreatestColumnWidth3: TMenuItem;
    GreatestColumnWidth4: TMenuItem;
    GreatestColumnWidth5: TMenuItem;
    Smallest2: TMenuItem;
    Smallest3: TMenuItem;
    Smallest4: TMenuItem;
    Smallest5: TMenuItem;
    AppendFromFile1: TMenuItem;
    AppendValuesFromDataSheet1: TMenuItem;
    AppendEmpty: TMenuItem;
    AppendS1: TMenuItem;
    AppendS2: TMenuItem;
    AppendS3: TMenuItem;
    AppendS4: TMenuItem;
    AppendS5: TMenuItem;
    SDExt: TSaveDialog;
    ODExt: TOpenDialog;
    OPD1: TOpenPictureDialog;
    Title0: TRadioButton;
    Title90: TRadioButton;
    Title180: TRadioButton;
    Title270: TRadioButton;
    Title360: TRadioButton;
    ExplorerButton1: TExplorerButton;
    AxisTitleRotation: TExplorerPopup;
    AxisTitleR: TExTrackBar;
    StatManHelp2: TMenuItem;
    DefaultHelpLanguage1: TMenuItem;
    EnglishHelp: TMenuItem;
    PersianHelp: TMenuItem;
    N15: TMenuItem;
    HideWhenMinimized1: TMenuItem;
    N16: TMenuItem;
    TTIcon: TTextTrayIcon;
    TrayPopup: TPopupMenu;
    OpenStatMan1: TMenuItem;
    About2: TMenuItem;
    N17: TMenuItem;
    Exit2: TMenuItem;
    Help2: TMenuItem;
    N18: TMenuItem;
    TrayIconImage: TImage;
    TrayTimer: TTimer;
    MemberPane2: TMenuItem;
    Edit1: TMenuItem;
    MCut: TMenuItem;
    MCopy: TMenuItem;
    MPaste: TMenuItem;
    HelpStyle1: TMenuItem;
    PromptForLan: TMenuItem;
    ShowDefaultLan: TMenuItem;
    Tools1: TMenuItem;
    Languages1: TMenuItem;
    English1: TMenuItem;
    Persian1: TMenuItem;
    LimitationTab: TTabSheet;
    StrLimitsTable: TStringGrid;
    CharStyleCombo: TComboBox;
    SGrid4: TAdvStringGrid;
    SGrid1: TAdvStringGrid;
    SGrid3: TAdvStringGrid;
    SGrid5: TAdvStringGrid;
    SGrid2: TAdvStringGrid;
    NewWorkProject1: TMenuItem;
    N19: TMenuItem;
    FreqReportPopup: TExplorerPopup;
    FreqReportEdit: TRichEdit;
    xpButton4: TxpButton;
    ToolBar1: TToolBar;
    FRBold: TSpeedButton;
    FRItalic: TSpeedButton;
    FRUnder: TSpeedButton;
    ToolButton14: TToolButton;
    FRLeft: TSpeedButton;
    FRCenter: TSpeedButton;
    FRRight: TSpeedButton;
    ToolButton15: TToolButton;
    FRSizeCombo: TComboBox;
    FRColorBtn: TColorButton;
    ToolButton16: TToolButton;
    SpeedButton12: TSpeedButton;
    xpButton5: TxpButton;
    N20: TMenuItem;
    G4Find: TMenuItem;
    G4Print: TMenuItem;
    N21: TMenuItem;
    Style1: TMenuItem;
    G4Font: TMenuItem;
    G4ShowBackImage: TMenuItem;
    G4ImageSub: TMenuItem;
    G4Default: TMenuItem;
    G4ChoosePicture: TMenuItem;
    G4SClassic: TMenuItem;
    G4SFlat: TMenuItem;
    G4SDefault: TMenuItem;
    GridBackImage: TImage;
    DataSheetMainMenu: TMenuItem;
    G4Replace: TMenuItem;
    G4ReplaceD: TAdvGridReplaceDialog;
    G4PrintSettingsD: TAdvGridPrintSettingsDialog;
    G1ReplaceD: TAdvGridReplaceDialog;
    G2ReplaceD: TAdvGridReplaceDialog;
    G3ReplaceD: TAdvGridReplaceDialog;
    G5ReplaceD: TAdvGridReplaceDialog;
    G1PrintSettingsD: TAdvGridPrintSettingsDialog;
    G2PrintSettingsD: TAdvGridPrintSettingsDialog;
    G3PrintSettingsD: TAdvGridPrintSettingsDialog;
    G5PrintSettingsD: TAdvGridPrintSettingsDialog;
    N23: TMenuItem;
    G1Find: TMenuItem;
    G1Replace: TMenuItem;
    G1Print: TMenuItem;
    N24: TMenuItem;
    Style2: TMenuItem;
    G1SDefault: TMenuItem;
    G1SClassic: TMenuItem;
    G1SFlat: TMenuItem;
    G1Font: TMenuItem;
    G1ShowBackImage: TMenuItem;
    G1ImageSub: TMenuItem;
    G1Default: TMenuItem;
    G1ChoosePicture: TMenuItem;
    N25: TMenuItem;
    G2Find: TMenuItem;
    G2Replace: TMenuItem;
    G2Print: TMenuItem;
    N26: TMenuItem;
    Style3: TMenuItem;
    G2SDefault: TMenuItem;
    G2SClassic: TMenuItem;
    G2SFlat: TMenuItem;
    G2Font: TMenuItem;
    G2ShowBackImage: TMenuItem;
    G2ImageSub: TMenuItem;
    G2Default: TMenuItem;
    G2ChoosePicture: TMenuItem;
    N27: TMenuItem;
    G3Find: TMenuItem;
    G3Replace: TMenuItem;
    G3Print: TMenuItem;
    N28: TMenuItem;
    Style4: TMenuItem;
    G3SDefault: TMenuItem;
    G3sClassic: TMenuItem;
    G3SFlat: TMenuItem;
    G3Font: TMenuItem;
    G3ShowBackImage: TMenuItem;
    G3ImageSub: TMenuItem;
    G3Default: TMenuItem;
    G3ChoosePicture: TMenuItem;
    N29: TMenuItem;
    G5Find: TMenuItem;
    G5Replace: TMenuItem;
    G5Print: TMenuItem;
    N30: TMenuItem;
    Style5: TMenuItem;
    G5SDefault: TMenuItem;
    G5SClassic: TMenuItem;
    G5SFlat: TMenuItem;
    G5Font: TMenuItem;
    G5ShowBackImage: TMenuItem;
    G5ImageSub: TMenuItem;
    G5Default: TMenuItem;
    G5ChoosePicture: TMenuItem;
    MSort: TMenuItem;
    MResize: TMenuItem;
    MGretestColumn: TMenuItem;
    MSmallestColumn: TMenuItem;
    N22: TMenuItem;
    MDefault: TMenuItem;
    N31: TMenuItem;
    MFind: TMenuItem;
    MReplace: TMenuItem;
    MPrint: TMenuItem;
    N32: TMenuItem;
    MGStyle: TMenuItem;
    MStyleDefault: TMenuItem;
    MStyleClassic: TMenuItem;
    MStyleFlat: TMenuItem;
    MFont: TMenuItem;
    N33: TMenuItem;
    G1NoBack: TMenuItem;
    N34: TMenuItem;
    G2NoBack: TMenuItem;
    N35: TMenuItem;
    G3NoBack: TMenuItem;
    N36: TMenuItem;
    G4NoBack: TMenuItem;
    N37: TMenuItem;
    G5NoBack: TMenuItem;
    XPMenu1: TXPMenu;
    N38: TMenuItem;
    PrinterSetup1: TMenuItem;
    OpenWorkProject1: TMenuItem;
    FileList: TListBox;
    N39: TMenuItem;
    SaveWorkProject1: TMenuItem;
    SaveWorkAs: TMenuItem;
    LoadDataSheet1: TMenuItem;
    AppendDataSheet1: TMenuItem;
    N40: TMenuItem;
    N41: TMenuItem;
    SaveDataSheetAs1: TMenuItem;
    HTMLFile1: TMenuItem;
    MicrosoftExcell1: TMenuItem;
    MicrosoftExcel2: TMenuItem;
    TextFile1: TMenuItem;
    WordDocument1: TMenuItem;
    N42: TMenuItem;
    LoadDataSheetXLS: TMenuItem;
    Notes1: TMenuItem;
    AnimItem: TMenuItem;
    FValueList: TListBox;
    Print2: TMenuItem;
    PrintDataSheet: TMenuItem;
    PrintChart: TMenuItem;
    N3: TMenuItem;
    PrintChart1: TMenuItem;
    TabSheet18: TTabSheet;
    ShowChartMarks: TxpCheckBox;
    ApplyChartGrid: TxpButton;
    ChartGrid: TStringGrid;
    ChartGridColorSelector: TColorSelector;
    xpButton6: TxpButton;
    xpButton7: TxpButton;
    GroupBox13: TGroupBox;
    RadioButton2: TRadioButton;
    RadioButton6: TRadioButton;
    RadioButton10: TRadioButton;
    RadioButton11: TRadioButton;
    RadioButton12: TRadioButton;
    RadioButton13: TRadioButton;
    RadioButton14: TRadioButton;
    ColorGrid: TStringGrid;
    FastCalcPopup: TExplorerPopup;
    ControlBar1: TControlBar;
    FastCalcAllMenu: TPopupMenu;
    Calculate1: TMenuItem;
    CopyToClipboard1: TMenuItem;
    FastCalc1: TMenuItem;
    StaticText3: TStaticText;
    CalcResultbtn: TSpeedButton;
    GroupBox14: TGroupBox;
    GroupBox15: TGroupBox;
    GroupBox22: TGroupBox;
    SpeedButton3: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    SpeedButton16: TSpeedButton;
    SpeedButton17: TSpeedButton;
    SpeedButton18: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton20: TSpeedButton;
    SpeedButton21: TSpeedButton;
    SpeedButton22: TSpeedButton;
    SpeedButton26: TSpeedButton;
    SpeedButton27: TSpeedButton;
    SpeedButton28: TSpeedButton;
    SpeedButton29: TSpeedButton;
    SpeedButton23: TSpeedButton;
    SpeedButton24: TSpeedButton;
    SpeedButton25: TSpeedButton;
    SpeedButton30: TSpeedButton;
    SpeedButton31: TSpeedButton;
    SpeedButton32: TSpeedButton;
    FastCalcParser: TParser;
    SpeedButton33: TSpeedButton;
    SaveAsPicture1: TMenuItem;
    HistoryBtn: TExplorerButton;
    ViewAsChartImage: TImage;
    N45: TMenuItem;
    RandomNumberProducer1: TMenuItem;
    GridCellPopup: TPopupMenu;
    CutCell: TMenuItem;
    CopyCell: TMenuItem;
    PasteCell: TMenuItem;
    N46: TMenuItem;
    DeleteCell: TMenuItem;
    ResizeCell: TMenuItem;
    SearchCell: TMenuItem;
    N47: TMenuItem;
    ColorCell: TMenuItem;
    FreqCell: TMenuItem;
    LockUnlockCell: TMenuItem;
    UnlockAllCell: TMenuItem;
    FilterCell: TMenuItem;
    N48: TMenuItem;
    FreqCellValue: TMenuItem;
    FilterCellValue: TMenuItem;
    UndoCellFilter: TMenuItem;
    Image3: TImage;
    OfficeImageList1: TOfficeImageList;
    N49: TMenuItem;
    TipOfTheDay: TMenuItem;
    FormulaEditor1: TMenuItem;
    ChartHistoryList: TExplorerPopup;
    HistoryChart1: TChart;
    Series8: TBarSeries;
    HistoryChart2: TChart;
    BarSeries1: TBarSeries;
    HistoryChart3: TChart;
    BarSeries2: TBarSeries;
    HistoryChart4: TChart;
    BarSeries3: TBarSeries;
    HistoryChart5: TChart;
    BarSeries4: TBarSeries;
    ChartHistoryPopup: TPopupMenu;
    Load2: TMenuItem;
    Append1: TMenuItem;
    N50: TMenuItem;
    Preview1: TMenuItem;
    HistoryLabel1: TLabel;
    HistoryLabel2: TLabel;
    HistoryLabel3: TLabel;
    HistoryLabel4: TLabel;
    HistoryLabel5: TLabel;
    N51: TMenuItem;
    ViewAsChart1: TMenuItem;
    ChartItem: TMenuItem;
    DesItem: TMenuItem;
    N54: TMenuItem;
    N55: TMenuItem;
    PrintViewChart: TMenuItem;
    TempChart: TChart;
    BarSeries5: TBarSeries;
    N52: TMenuItem;
    AsPicture1: TMenuItem;
    AsMetafile1: TMenuItem;
    AsPicture2: TMenuItem;
    AsMetafile2: TMenuItem;
    SendToChartViewer1: TMenuItem;
    N53: TMenuItem;
    ViewAsChart2: TMenuItem;
    N57: TMenuItem;
    ViewAsChart3: TMenuItem;
    N58: TMenuItem;
    N59: TMenuItem;
    ViewAsChart5: TMenuItem;
    ViewAsChart4: TMenuItem;
    AsChart2: TMenuItem;
    DesItem2: TMenuItem;
    Copy2: TMenuItem;
    Save2: TMenuItem;
    Print3: TMenuItem;
    N62: TMenuItem;
    AsChart3: TMenuItem;
    DesItem3: TMenuItem;
    Copy3: TMenuItem;
    Save3: TMenuItem;
    Print4: TMenuItem;
    N65: TMenuItem;
    AsChart4: TMenuItem;
    DesItem4: TMenuItem;
    Copy4: TMenuItem;
    Save4: TMenuItem;
    Print5: TMenuItem;
    N68: TMenuItem;
    AsChart5: TMenuItem;
    DesItem5: TMenuItem;
    Copy5: TMenuItem;
    Save5: TMenuItem;
    Print6: TMenuItem;
    N71: TMenuItem;
    SendToChartViewer2: TMenuItem;
    SendToChartViewer3: TMenuItem;
    SendToChartViewer4: TMenuItem;
    SendToChartViewer5: TMenuItem;
    AsPicture3: TMenuItem;
    AsMetafile3: TMenuItem;
    AsPicture4: TMenuItem;
    AsMetafile4: TMenuItem;
    AsPicture5: TMenuItem;
    AsMetafile5: TMenuItem;
    AsPicture6: TMenuItem;
    AsMetafile6: TMenuItem;
    AsPicture7: TMenuItem;
    AsMetafile7: TMenuItem;
    AsPicture8: TMenuItem;
    AsMetafile8: TMenuItem;
    AsPicture9: TMenuItem;
    AsMetafile9: TMenuItem;
    AsPicture10: TMenuItem;
    AsMetafile10: TMenuItem;
    GlobalToolbar: TToolBar;
    CutToolbtn: TToolButton;
    CopyToolbtn: TToolButton;
    PasteToolbtn: TToolButton;
    ToolButton20: TToolButton;
    FindToolbtn: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton26: TToolButton;
    PrintToolbtn: TToolButton;
    ToolButton28: TToolButton;
    ToolButton29: TToolButton;
    PrintToolPopup: TPopupMenu;
    DataSheet2: TMenuItem;
    Chart4: TMenuItem;
    GlobalToolbar1: TMenuItem;
    DettachAttachPane: TMenuItem;
    TempAdvGrid: TAdvStringGrid;
    StatAnim1: TImageList;
    Image5: TImage;
    AnimTimer: TTimer;
    FreqExportPopup: TPopupMenu;
    MicrosoftExcelSpreadSheet1: TMenuItem;
    WordDocument2: TMenuItem;
    HTMLFile2: TMenuItem;
    extFile1: TMenuItem;
    TempPrintSettingsD: TAdvGridPrintSettingsDialog;
    FrequencyTableToolbar1: TMenuItem;
    Image6: TImage;
    Image7: TImage;
    Image8: TImage;
    Image9: TImage;
    Image10: TImage;
    Image11: TImage;
    Image12: TImage;
    Image13: TImage;
    Image14: TImage;
    Image15: TImage;
    Image16: TImage;
    Image17: TImage;
    Image18: TImage;
    Image19: TImage;
    Image20: TImage;
    Image21: TImage;
    Image22: TImage;
    Image23: TImage;
    Image24: TImage;
    Image25: TImage;
    Image26: TImage;
    Image27: TImage;
    Image28: TImage;
    N43: TMenuItem;
    PromptAsMenu: TMenuItem;
    PromptAsList: TMenuItem;
    HelpLangMenu: TPopupMenu;
    English2: TMenuItem;
    Persian2: TMenuItem;
    N60: TMenuItem;
    Add1: TxpButton;
    Del1: TxpButton;
    EditFml: TxpButton;
    VEdit1: TEdit;
    Label4: TLabel;
    StatusControlBar: TControlBar;
    StatusBar: TStatusBar;
    FastCalcToolbar: TToolBar;
    FastExpression: TEdit;
    FastCalcbtn: TExplorerButton;
    FastResult: TStaticText;
    FastLabel: TLabel;
    FreqControlBar: TControlBar;
    FreqToolbar: TToolBar;
    ReportFreqTablebtn: TExplorerButton;
    FreqExportBtn: TExplorerButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    DettachFreqToolbar: TSpeedButton;
    Label94: TLabel;
    S4EditFast: TEdit;
    S4AddFastbtn: TBitBtn;
    Label95: TLabel;
    S1EditFast: TEdit;
    S1AddFastbtn: TBitBtn;
    Label96: TLabel;
    S3EditFast: TEdit;
    S3AddFastbtn: TBitBtn;
    Label97: TLabel;
    S5EditFast: TEdit;
    S5AddFastbtn: TBitBtn;
    Label98: TLabel;
    S2EditFast: TEdit;
    S2AddFastbtn: TBitBtn;
    procedure AlertTimerTimer(Sender: TObject);
    procedure MemberButtonNormalMouseMove(Sender: TObject;
      Shift: TShiftState; X, Y: Integer);
    procedure MemberButtonPressedMouseMove(Sender: TObject;
      Shift: TShiftState; X, Y: Integer);
    procedure MemberButtonPressedMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure EditFmlClick(Sender: TObject);
    procedure AlertImageClick(Sender: TObject);
    procedure HelpLabelMouseLeave(Sender: TObject);
    procedure PaneMenuBtnMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MemberMenuPopup(Sender: TObject);
    procedure SettingGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ForceValueClick(Sender: TObject);
    procedure S1CheckClick(Sender: TObject);
    procedure VEdit1Change(Sender: TObject);
    procedure List1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SettingGridSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure S2CheckClick(Sender: TObject);
    procedure S3CheckClick(Sender: TObject);
    procedure S4CheckClick(Sender: TObject);
    procedure S5CheckClick(Sender: TObject);
    procedure HelpLabelMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure UseFilterClick(Sender: TObject);
    procedure FxLabelDblClick(Sender: TObject);
    procedure Add1Click(Sender: TObject);
    procedure Del1Click(Sender: TObject);
    procedure VEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure TypeComboChange(Sender: TObject);
    procedure SGrid1SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure S1ComboKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure S1ComboChange(Sender: TObject);
    procedure Sheet1SectionClick(Sender: TObject);
    procedure SGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure S1ClrbtnClick(Sender: TObject);
    procedure SGrid1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SettingSectionClick(Sender: TObject);
    procedure Sheet2SectionClick(Sender: TObject);
    procedure Sheet3SectionClick(Sender: TObject);
    procedure Sheet4SectionClick(Sender: TObject);
    procedure Sheet5SectionClick(Sender: TObject);
    procedure MSettingSectionClick(Sender: TObject);
    procedure MSheet1SectionClick(Sender: TObject);
    procedure MSheet2SectionClick(Sender: TObject);
    procedure MSheet3SectionClick(Sender: TObject);
    procedure MSheet4SectionClick(Sender: TObject);
    procedure MSheet5SectionClick(Sender: TObject);
    procedure MTableSectionClick(Sender: TObject);
    procedure MChartSectionClick(Sender: TObject);
    procedure MAnalyzeSectionClick(Sender: TObject);
    procedure S5ComboChange(Sender: TObject);
    procedure S5ComboKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SGrid52KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SGrid52MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SGrid52SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure SGrid52SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure S5ClrbtnClick(Sender: TObject);
    procedure S4ComboChange(Sender: TObject);
    procedure S3ComboChange(Sender: TObject);
    procedure S2ComboChange(Sender: TObject);
    procedure SGrid4SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure SGrid3SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure SGrid2SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure SGrid4MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SGrid3MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SGrid2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure SGrid4KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SGrid3KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SGrid2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure S4ClrbtnClick(Sender: TObject);
    procedure S3ClrbtnClick(Sender: TObject);
    procedure S2ClrbtnClick(Sender: TObject);
    procedure S4ComboKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure S3ComboKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure S2ComboKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure IconViewClick(Sender: TObject);
    procedure ToolViewClick(Sender: TObject);
    procedure ListViewClick(Sender: TObject);
    procedure MenuViewClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ToolSettingsClick(Sender: TObject);
    procedure ToolSheet1Click(Sender: TObject);
    procedure ToolSheet2Click(Sender: TObject);
    procedure ToolSheet3Click(Sender: TObject);
    procedure ToolSheet4Click(Sender: TObject);
    procedure ToolSheet5Click(Sender: TObject);
    procedure ToolTableClick(Sender: TObject);
    procedure ToolChartClick(Sender: TObject);
    procedure ToolAnalyzeClick(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure Edit3Change(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure TableSectionClick(Sender: TObject);
    procedure IconSettingsMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure MemberGroupMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure CreateTable2Click(Sender: TObject);
    procedure ApplicationEvents1ShowHint(var HintStr: String;
      var CanShow: Boolean; var HintInfo: THintInfo);
    procedure AnalyzeSectionClick(Sender: TObject);
    procedure ModeListTimerTimer(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure ModeBtnClick(Sender: TObject);
    procedure ExplorerButton8Click(Sender: TObject);
    procedure Splitter2CanResize(Sender: TObject; var NewSize: Integer;
      var Accept: Boolean);
    procedure Splitter3Moved(Sender: TObject);
    procedure SizeComboChange(Sender: TObject);
    procedure ColorBtn1ColorSelected(Sender: TObject; AColor: TColor);
    procedure BoldBtn1Click(Sender: TObject);
    procedure ItalicBtn1Click(Sender: TObject);
    procedure UnderBtn1Click(Sender: TObject);
    procedure LeftAlClick(Sender: TObject);
    procedure CenterAlClick(Sender: TObject);
    procedure RightAlClick(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure FD1Apply(Sender: TObject; Wnd: HWND);
    procedure BoldBtn2Click(Sender: TObject);
    procedure ItalicBtn2Click(Sender: TObject);
    procedure UnderBtn2Click(Sender: TObject);
    procedure LeftAl2Click(Sender: TObject);
    procedure CenterAl2Click(Sender: TObject);
    procedure RightAl2Click(Sender: TObject);
    procedure SizeCombo2Change(Sender: TObject);
    procedure ColorBtn2ColorSelected(Sender: TObject; AColor: TColor);
    procedure SpeedButton10Click(Sender: TObject);
    procedure BoldBtn3Click(Sender: TObject);
    procedure ItalicBtn3Click(Sender: TObject);
    procedure UnderBtn3Click(Sender: TObject);
    procedure SizeCombo3Change(Sender: TObject);
    procedure ColorBtn3ColorSelected(Sender: TObject; AColor: TColor);
    procedure SpeedButton11Click(Sender: TObject);
    procedure ChartBtnClick(Sender: TObject);
    procedure DetachBtnClick(Sender: TObject);
    procedure xpCheckBox7Click(Sender: TObject);
    procedure LegendPosChange(Sender: TObject);
    procedure ColorSelector3ChangeColor(Sender: TObject);
    procedure SpinEdit34Change(Sender: TObject);
    procedure xpButton2Click(Sender: TObject);
    procedure LegendStyleChange(Sender: TObject);
    procedure xpCheckBox8Click(Sender: TObject);
    procedure xpCheckBox9Click(Sender: TObject);
    procedure xpCheckBox10Click(Sender: TObject);
    procedure lwidth2Change(Sender: TObject);
    procedure lstyle2Change(Sender: TObject);
    procedure HorizMargChange(Sender: TObject);
    procedure VertMargChange(Sender: TObject);
    procedure xpCheckBox11Click(Sender: TObject);
    procedure View3DCheckClick(Sender: TObject);
    procedure SpinEdit5Change(Sender: TObject);
    procedure SpinEdit6Change(Sender: TObject);
    procedure xpCheckBox13Click(Sender: TObject);
    procedure xpCheckBox14Click(Sender: TObject);
    procedure penwidthcombo3Change(Sender: TObject);
    procedure penwidthcombo4Change(Sender: TObject);
    procedure penstylecombo2Change(Sender: TObject);
    procedure penstylecombo3Change(Sender: TObject);
    procedure UseColorsClick(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure xpCheckBox6Click(Sender: TObject);
    procedure ZoomTrackChange(Sender: TObject);
    procedure Chart3DTrackChange(Sender: TObject);
    procedure XRChange(Sender: TObject);
    procedure YRChange(Sender: TObject);
    procedure ZRChange(Sender: TObject);
    procedure PerspectiveTrackChange(Sender: TObject);
    procedure NormalViewClick(Sender: TObject);
    procedure CustomizedViewClick(Sender: TObject);
    procedure xpCheckBox15Click(Sender: TObject);
    procedure xpCheckBox17Click(Sender: TObject);
    procedure ResizeTClick(Sender: TObject);
    procedure ResizeFClick(Sender: TObject);
    procedure xpCheckBox20Click(Sender: TObject);
    procedure xpCheckBox16Click(Sender: TObject);
    procedure bcolorfClick(Sender: TObject);
    procedure bwidthChange(Sender: TObject);
    procedure bwidthfChange(Sender: TObject);
    procedure bstylefChange(Sender: TObject);
    procedure bstyleChange(Sender: TObject);
    procedure ChartTitleEditChange(Sender: TObject);
    procedure ChartFooterEditChange(Sender: TObject);
    procedure RadioButton8Click(Sender: TObject);
    procedure RadioButton9Click(Sender: TObject);
    procedure SpinEdit10Change(Sender: TObject);
    procedure SpinEdit11Change(Sender: TObject);
    procedure SpinEdit12Change(Sender: TObject);
    procedure xpCheckBox21Click(Sender: TObject);
    procedure xpCheckBox22Click(Sender: TObject);
    procedure xpCheckBox23Click(Sender: TObject);
    procedure labelcustomClick(Sender: TObject);
    procedure mmtrackChange(Sender: TObject);
    procedure mmlabelClick(Sender: TObject);
    procedure xpCheckBox24Click(Sender: TObject);
    procedure AxisTitleRChange(Sender: TObject);
    procedure disSpinChange(Sender: TObject);
    procedure xpCheckBox26Click(Sender: TObject);
    procedure xpCheckBox28Click(Sender: TObject);
    procedure xpCheckBox27Click(Sender: TObject);
    procedure miCountChange(Sender: TObject);
    procedure miLengthChange(Sender: TObject);
    procedure miWidthChange(Sender: TObject);
    procedure miStyleChange(Sender: TObject);
    procedure MLengthChange(Sender: TObject);
    procedure ILengthChange(Sender: TObject);
    procedure IWidthChange(Sender: TObject);
    procedure MWidthChange(Sender: TObject);
    procedure MStyleChange(Sender: TObject);
    procedure IStyleChange(Sender: TObject);
    procedure IColorChangeColor(Sender: TObject);
    procedure MColorChangeColor(Sender: TObject);
    procedure miColorChangeColor(Sender: TObject);
    procedure MLengthKeyPress(Sender: TObject; var Key: Char);
    procedure ILengthKeyPress(Sender: TObject; var Key: Char);
    procedure miCountKeyPress(Sender: TObject; var Key: Char);
    procedure miLengthKeyPress(Sender: TObject; var Key: Char);
    procedure disSpinKeyPress(Sender: TObject; var Key: Char);
    procedure minspinKeyPress(Sender: TObject; var Key: Char);
    procedure SpinEdit5KeyPress(Sender: TObject; var Key: Char);
    procedure HorizMargKeyPress(Sender: TObject; var Key: Char);
    procedure LegendWidthKeyPress(Sender: TObject; var Key: Char);
    procedure LineBtnClick(Sender: TObject);
    procedure bcClick(Sender: TObject);
    procedure BackColorChangeColor(Sender: TObject);
    procedure bfClick(Sender: TObject);
    procedure FrameColorChangeColor(Sender: TObject);
    procedure cgClick(Sender: TObject);
    procedure StartColorChange(Sender: TObject);
    procedure EndColorChange(Sender: TObject);
    procedure UseBackImageClick(Sender: TObject);
    procedure LineBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure ColorTChangeColor(Sender: TObject);
    procedure bColorChangeColor(Sender: TObject);
    procedure StaticText25Click(Sender: TObject);
    procedure TDefBorderClick(Sender: TObject);
    procedure StaticText26Click(Sender: TObject);
    procedure StaticText29Click(Sender: TObject);
    procedure FRUnderClick(Sender: TObject);
    procedure lcolorChangeColor(Sender: TObject);
    procedure ShadowColorChangeColor(Sender: TObject);
    procedure ShadowSizeChange(Sender: TObject);
    procedure ColorSelector6ChangeColor(Sender: TObject);
    procedure ColorSelector9ChangeColor(Sender: TObject);
    procedure ColorSelector8ChangeColor(Sender: TObject);
    procedure ColorSelector7ChangeColor(Sender: TObject);
    procedure StaticText30Click(Sender: TObject);
    procedure StaticText31Click(Sender: TObject);
    procedure MaxSpinKeyPress(Sender: TObject; var Key: Char);
    procedure minspinChange(Sender: TObject);
    procedure MaxSpinChange(Sender: TObject);
    procedure IncrementChange(Sender: TObject);
    procedure AxisTitleEditChange(Sender: TObject);
    procedure TitleSizeKeyPress(Sender: TObject; var Key: Char);
    procedure TitleSizeChange(Sender: TObject);
    procedure xpCheckBox29Click(Sender: TObject);
    procedure axcolorChangeColor(Sender: TObject);
    procedure gridcolorChangeColor(Sender: TObject);
    procedure axwidthChange(Sender: TObject);
    procedure gridwidthChange(Sender: TObject);
    procedure axstyleChange(Sender: TObject);
    procedure gridstyleChange(Sender: TObject);
    procedure xpCheckBox30Click(Sender: TObject);
    procedure xpCheckBox1Click(Sender: TObject);
    procedure gridstylevChange(Sender: TObject);
    procedure gridwidthvChange(Sender: TObject);
    procedure gridcolorvChangeColor(Sender: TObject);
    procedure PrintBtnClick(Sender: TObject);
    procedure CopyAsBitmap1Click(Sender: TObject);
    procedure CopyAsMetafile1Click(Sender: TObject);
    procedure UseColors2Click(Sender: TObject);
    procedure RadioButton6Click(Sender: TObject);
    procedure ExplorerPopup1Open(Sender: TObject);
    procedure RBoldClick(Sender: TObject);
    procedure RItalicClick(Sender: TObject);
    procedure RUnderClick(Sender: TObject);
    procedure RLeftClick(Sender: TObject);
    procedure RCenterClick(Sender: TObject);
    procedure RRightClick(Sender: TObject);
    procedure RSizeComboChange(Sender: TObject);
    procedure RColorBtnColorSelected(Sender: TObject; AColor: TColor);
    procedure RBuildClick(Sender: TObject);
    procedure xpButton3Click(Sender: TObject);
    procedure xpButton1Click(Sender: TObject);
    procedure Load1Click(Sender: TObject);
    procedure Save1Click(Sender: TObject);
    procedure ExplorerButton11DropDownClick(Sender: TObject);
    procedure ExplorerButton12DropDownClick(Sender: TObject);
    procedure ExplorerButton13DropDownClick(Sender: TObject);
    procedure ExplorerButton14DropDownClick(Sender: TObject);
    procedure ExplorerButton15DropDownClick(Sender: TObject);
    procedure ChartSectionClick(Sender: TObject);
    procedure SGrid5MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure S4FilterbtnClick(Sender: TObject);
    procedure S4PSpinKeyPress(Sender: TObject; var Key: Char);
    procedure S1FilterbtnClick(Sender: TObject);
    procedure S3FilterbtnClick(Sender: TObject);
    procedure S5FilterbtnClick(Sender: TObject);
    procedure S2FilterbtnClick(Sender: TObject);
    procedure SGrid5SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure AnalyzeReportbtnDropDownClick(Sender: TObject);
    procedure SaveAsPicture1Click(Sender: TObject);
    procedure SaveAsMetafile1Click(Sender: TObject);
    procedure Add1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FxLabelMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ForceValueMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure UseFilterMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure TypeComboDropDown(Sender: TObject);
    procedure HelpLabelMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ReportPopupOpen(Sender: TObject);
    procedure Label15MouseEnter(Sender: TObject);
    procedure Label14MouseEnter(Sender: TObject);
    procedure Label13MouseEnter(Sender: TObject);
    procedure Label12MouseEnter(Sender: TObject);
    procedure Label11MouseEnter(Sender: TObject);
    procedure Label10MouseEnter(Sender: TObject);
    procedure Label9MouseEnter(Sender: TObject);
    procedure Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Image2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure BarBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure HBarBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure AreaBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure PointBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure ChartLabelMouseEnter(Sender: TObject);
    procedure PieBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure FastBtnMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure Print1Click(Sender: TObject);
    procedure CopyAsPicture1Click(Sender: TObject);
    procedure CopyAsMetafile2Click(Sender: TObject);
    procedure istyle2Change(Sender: TObject);
    procedure iwidth2Change(Sender: TObject);
    procedure G1SortClick(Sender: TObject);
    procedure G2SortClick(Sender: TObject);
    procedure G3SortClick(Sender: TObject);
    procedure G4SortClick(Sender: TObject);
    procedure G5SortClick(Sender: TObject);
    procedure StaticText32Click(Sender: TObject);
    procedure StaticText33Click(Sender: TObject);
    procedure StaticText34Click(Sender: TObject);
    procedure StaticText35Click(Sender: TObject);
    procedure StaticText36Click(Sender: TObject);
    procedure ListView1Click(Sender: TObject);
    procedure IconView1Click(Sender: TObject);
    procedure ToolBarView1Click(Sender: TObject);
    procedure MenuView1Click(Sender: TObject);
    procedure StatusBar1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure ApplicationEvents1Message(var Msg: tagMSG;
      var Handled: Boolean);
    procedure DataSettings1Click(Sender: TObject);
    procedure DataSheetMainMenuClick(Sender: TObject);
    procedure DataSheet2Click(Sender: TObject);
    procedure DataSheet3Click(Sender: TObject);
    procedure DataSheet4Click(Sender: TObject);
    procedure DataSheet5Click(Sender: TObject);
    procedure Charts1Click(Sender: TObject);
    procedure DataAnalyzing1Click(Sender: TObject);
    procedure Chart3Click(Sender: TObject);
    procedure c1Click(Sender: TObject);
    procedure PrintChart1Click(Sender: TObject);
    procedure c3Click(Sender: TObject);
    procedure c4Click(Sender: TObject);
    procedure c5Click(Sender: TObject);
    procedure c6Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure ChartToolbar1Click(Sender: TObject);
    procedure StatManHelp1Click(Sender: TObject);
    procedure Default1Click(Sender: TObject);
    procedure GreatestColumnWidth1Click(Sender: TObject);
    procedure Smallest1Click(Sender: TObject);
    procedure Default2Click(Sender: TObject);
    procedure Default3Click(Sender: TObject);
    procedure Default4Click(Sender: TObject);
    procedure Default5Click(Sender: TObject);
    procedure GreatestColumnWidth2Click(Sender: TObject);
    procedure GreatestColumnWidth3Click(Sender: TObject);
    procedure GreatestColumnWidth4Click(Sender: TObject);
    procedure GreatestColumnWidth5Click(Sender: TObject);
    procedure Smallest2Click(Sender: TObject);
    procedure Smallest3Click(Sender: TObject);
    procedure Smallest4Click(Sender: TObject);
    procedure Smallest5Click(Sender: TObject);
    procedure AppendFromFile1Click(Sender: TObject);
    procedure SaveGridPopupPopup(Sender: TObject);
    procedure AppendS5Click(Sender: TObject);
    procedure BrowseImageClick(Sender: TObject);
    procedure TileRadioClick(Sender: TObject);
    procedure PutInsideClick(Sender: TObject);
    procedure ColorFChangeColor(Sender: TObject);
    procedure Title0Click(Sender: TObject);
    procedure AxisTitleRotationOpen(Sender: TObject);
    procedure SaveAsPicture2Click(Sender: TObject);
    procedure SaveAsMetafile2Click(Sender: TObject);
    procedure SD1CanClose(Sender: TObject; var CanClose: Boolean);
    procedure EnglishHelpClick(Sender: TObject);
    procedure PersianHelpClick(Sender: TObject);
    procedure StatManHelp2Click(Sender: TObject);
    procedure HideWhenMinimized1Click(Sender: TObject);
    procedure OpenStatMan1Click(Sender: TObject);
    procedure Help2Click(Sender: TObject);
    procedure About2Click(Sender: TObject);
    procedure Exit2Click(Sender: TObject);
    procedure TTIconBalloonHintClick(Sender: TObject);
    procedure TrayTimerTimer(Sender: TObject);
    procedure TTIconClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure SettingGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SettingGridMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SettingGridMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure MemberPane2Click(Sender: TObject);
    procedure TTIconMinimizeToTray(Sender: TObject);
    procedure PromptForLanClick(Sender: TObject);
    procedure ShowDefaultLanClick(Sender: TObject);
    procedure English1Click(Sender: TObject);
    procedure Persian1Click(Sender: TObject);
    procedure StrLimitsTableSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure CharStyleComboChange(Sender: TObject);
    procedure StrLimitsTableExit(Sender: TObject);
    procedure SGrid4KeyPress(Sender: TObject; var Key: Char);
    procedure SGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure SGrid3KeyPress(Sender: TObject; var Key: Char);
    procedure SGrid5KeyPress(Sender: TObject; var Key: Char);
    procedure SGrid2KeyPress(Sender: TObject; var Key: Char);
    procedure NewWorkProject1Click(Sender: TObject);
    procedure xpButton4Click(Sender: TObject);
    procedure xpButton5Click(Sender: TObject);
    procedure FRBoldClick(Sender: TObject);
    procedure FRItalicClick(Sender: TObject);
    procedure FRLeftClick(Sender: TObject);
    procedure FRCenterClick(Sender: TObject);
    procedure FRRightClick(Sender: TObject);
    procedure FRSizeComboChange(Sender: TObject);
    procedure FRColorBtnColorSelected(Sender: TObject; AColor: TColor);
    procedure SpeedButton12Click(Sender: TObject);
    procedure ReportFreqTablebtnDropDownClick(Sender: TObject);
    procedure SGrid4SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure SGrid1SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure SGrid3SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure SGrid5SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure SGrid2SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure Label8MouseEnter(Sender: TObject);
    procedure Label5MouseEnter(Sender: TObject);
    procedure Label6MouseEnter(Sender: TObject);
    procedure Label7MouseEnter(Sender: TObject);
    procedure G4ChoosePictureClick(Sender: TObject);
    procedure G4DefaultClick(Sender: TObject);
    procedure G4ShowBackImageClick(Sender: TObject);
    procedure G4FontClick(Sender: TObject);
    procedure G4SDefaultClick(Sender: TObject);
    procedure G4SClassicClick(Sender: TObject);
    procedure G4SFlatClick(Sender: TObject);
    procedure G4ReplaceClick(Sender: TObject);
    procedure G4FindClick(Sender: TObject);
    procedure G4PrintClick(Sender: TObject);
    procedure G1DefaultClick(Sender: TObject);
    procedure G2DefaultClick(Sender: TObject);
    procedure G3DefaultClick(Sender: TObject);
    procedure G5DefaultClick(Sender: TObject);
    procedure G1ChoosePictureClick(Sender: TObject);
    procedure G2ChoosePictureClick(Sender: TObject);
    procedure G3ChoosePictureClick(Sender: TObject);
    procedure G5ChoosePictureClick(Sender: TObject);
    procedure G1ShowBackImageClick(Sender: TObject);
    procedure G2ShowBackImageClick(Sender: TObject);
    procedure G3ShowBackImageClick(Sender: TObject);
    procedure G5ShowBackImageClick(Sender: TObject);
    procedure G1FontClick(Sender: TObject);
    procedure G2FontClick(Sender: TObject);
    procedure G3FontClick(Sender: TObject);
    procedure G5FontClick(Sender: TObject);
    procedure G1PrintClick(Sender: TObject);
    procedure G2PrintClick(Sender: TObject);
    procedure G3PrintClick(Sender: TObject);
    procedure G5PrintClick(Sender: TObject);
    procedure G1FindClick(Sender: TObject);
    procedure G1ReplaceClick(Sender: TObject);
    procedure G2FindClick(Sender: TObject);
    procedure G2ReplaceClick(Sender: TObject);
    procedure G3FindClick(Sender: TObject);
    procedure G3ReplaceClick(Sender: TObject);
    procedure G5FindClick(Sender: TObject);
    procedure G5ReplaceClick(Sender: TObject);
    procedure G1SDefaultClick(Sender: TObject);
    procedure G1SClassicClick(Sender: TObject);
    procedure G1SFlatClick(Sender: TObject);
    procedure G2SDefaultClick(Sender: TObject);
    procedure G2SClassicClick(Sender: TObject);
    procedure G2SFlatClick(Sender: TObject);
    procedure G3SDefaultClick(Sender: TObject);
    procedure G3sClassicClick(Sender: TObject);
    procedure G3SFlatClick(Sender: TObject);
    procedure G5SDefaultClick(Sender: TObject);
    procedure G5SClassicClick(Sender: TObject);
    procedure G5SFlatClick(Sender: TObject);
    procedure SGrid4DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SGrid3DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SGrid5DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SGrid2DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure G5NoBackClick(Sender: TObject);
    procedure OpenWorkProject1Click(Sender: TObject);
    procedure File1Click(Sender: TObject);
    procedure MicrosoftExcell1Click(Sender: TObject);
    procedure MicrosoftExcel2Click(Sender: TObject);
    procedure WordDocument1Click(Sender: TObject);
    procedure HTMLFile1Click(Sender: TObject);
    procedure TextFile1Click(Sender: TObject);
    procedure LoadDataSheetXLSClick(Sender: TObject);
    procedure SaveWorkProject1Click(Sender: TObject);
    procedure SaveWorkAsClick(Sender: TObject);
    procedure LoadDataSheet1Click(Sender: TObject);
    procedure ApplyChartGridClick(Sender: TObject);
    procedure ChartGridKeyPress(Sender: TObject; var Key: Char);
    procedure ChartGridColorSelectorChangeColor(Sender: TObject);
    procedure ChartGridRowMoved(Sender: TObject; FromIndex,
      ToIndex: Integer);
    procedure xpButton7Click(Sender: TObject);
    procedure xpButton6Click(Sender: TObject);
    procedure ShowChartMarksClick(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure ColorGridTopLeftChanged(Sender: TObject);
    procedure ChartGridTopLeftChanged(Sender: TObject);
    procedure ColorGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure ColorGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ColorGridExit(Sender: TObject);
    procedure ChartGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ColorGridEnter(Sender: TObject);
    procedure FastLabelClick(Sender: TObject);
    procedure FastLabelMouseEnter(Sender: TObject);
    procedure FastLabelMouseLeave(Sender: TObject);
    procedure FastCalc1Click(Sender: TObject);
    procedure CalcResultbtnClick(Sender: TObject);
    procedure SpeedButton26Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton23Click(Sender: TObject);
    procedure SpeedButton33Click(Sender: TObject);
    procedure Calculate1Click(Sender: TObject);
    procedure CopyToClipboard1Click(Sender: TObject);
    procedure FastExpressionKeyPress(Sender: TObject; var Key: Char);
    procedure BitBtn3Click(Sender: TObject);
    procedure PrinterSetup1Click(Sender: TObject);
    procedure RandomNumberProducer1Click(Sender: TObject);
    procedure SGrid4MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CutCellClick(Sender: TObject);
    procedure CopyCellClick(Sender: TObject);
    procedure PasteCellClick(Sender: TObject);
    procedure GridCellPopupPopup(Sender: TObject);
    procedure DeleteCellClick(Sender: TObject);
    procedure ResizeCellClick(Sender: TObject);
    procedure ColorCellClick(Sender: TObject);
    procedure FreqCellClick(Sender: TObject);
    procedure SearchCellClick(Sender: TObject);
    procedure FilterCellValueClick(Sender: TObject);
    procedure UndoCellFilterClick(Sender: TObject);
    procedure SGrid1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SGrid3MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SGrid5MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SGrid2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure LockUnlockCellClick(Sender: TObject);
    procedure UnlockAllCellClick(Sender: TObject);
    procedure SGrid4CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure SGrid1CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure SGrid3CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure SGrid5CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure SGrid2CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure TipOfTheDayClick(Sender: TObject);
    procedure FormulaEditor1Click(Sender: TObject);
    procedure HistoryLabel1MouseEnter(Sender: TObject);
    procedure HistoryLabel1MouseLeave(Sender: TObject);
    procedure HistoryLabel1Click(Sender: TObject);
    procedure HistoryBtnDropDownClick(Sender: TObject);
    procedure Load2Click(Sender: TObject);
    procedure Append1Click(Sender: TObject);
    procedure Preview1Click(Sender: TObject);
    procedure GridPopup1Popup(Sender: TObject);
    procedure ChartItemDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; Selected: Boolean);
    procedure ChartItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
      var Width, Height: Integer);
    procedure AsPicture1Click(Sender: TObject);
    procedure AsMetafile1Click(Sender: TObject);
    procedure AsPicture2Click(Sender: TObject);
    procedure AsMetafile2Click(Sender: TObject);
    procedure PrintViewChartClick(Sender: TObject);
    procedure SendToChartViewer1Click(Sender: TObject);
    procedure DesItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
      var Width, Height: Integer);
    procedure DesItemAdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; State: TOwnerDrawState);
    procedure GridPopup2Popup(Sender: TObject);
    procedure GridPopup3Popup(Sender: TObject);
    procedure GridPopup4Popup(Sender: TObject);
    procedure GridPopup5Popup(Sender: TObject);
    procedure ViewAsChart1Click(Sender: TObject);
    procedure ViewAsChart2Click(Sender: TObject);
    procedure ViewAsChart3Click(Sender: TObject);
    procedure ViewAsChart4Click(Sender: TObject);
    procedure ViewAsChart5Click(Sender: TObject);
    procedure DesItem2AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; State: TOwnerDrawState);
    procedure DesItem3AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; State: TOwnerDrawState);
    procedure DesItem4AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; State: TOwnerDrawState);
    procedure DesItem5AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; State: TOwnerDrawState);
    procedure Chart4Click(Sender: TObject);
    procedure CutToolbtnClick(Sender: TObject);
    procedure CopyToolbtnClick(Sender: TObject);
    procedure PasteToolbtnClick(Sender: TObject);
    procedure FindToolbtnClick(Sender: TObject);
    procedure ToolButton28Click(Sender: TObject);
    procedure ToolButton24Click(Sender: TObject);
    procedure ToolButton25Click(Sender: TObject);
    procedure ToolButton23Click(Sender: TObject);
    procedure GlobalToolbar1Click(Sender: TObject);
    procedure Print2Click(Sender: TObject);
    procedure PrintDataSheetClick(Sender: TObject);
    procedure PrintChartClick(Sender: TObject);
    procedure PrintToolbtnClick(Sender: TObject);
    procedure ControlBar1Resize(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure G1NoBackClick(Sender: TObject);
    procedure G2NoBackClick(Sender: TObject);
    procedure G3NoBackClick(Sender: TObject);
    procedure G4NoBackClick(Sender: TObject);
    procedure FreqCellValueAdvancedDrawItem(Sender: TObject;
      ACanvas: TCanvas; ARect: TRect; State: TOwnerDrawState);
    procedure FreqCellValueMeasureItem(Sender: TObject; ACanvas: TCanvas;
      var Width, Height: Integer);
    procedure DettachAttachPaneClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Edit1Click(Sender: TObject);
    procedure MCutClick(Sender: TObject);
    procedure MCopyClick(Sender: TObject);
    procedure MPasteClick(Sender: TObject);
    procedure N60DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect;
      Selected: Boolean);
    procedure N63DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect;
      Selected: Boolean);
    procedure N66DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect;
      Selected: Boolean);
    procedure N69DrawItem(Sender: TObject; ACanvas: TCanvas; ARect: TRect;
      Selected: Boolean);
    procedure AnimTimerTimer(Sender: TObject);
    procedure AnimItemAdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; State: TOwnerDrawState);
    procedure CloseAnaReportClick(Sender: TObject);
    procedure AnimItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
      var Width, Height: Integer);
    procedure TTIconDblClick(Sender: TObject);
    procedure TableGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FreqExportBtnDropDownClick(Sender: TObject);
    procedure ToolButton18Click(Sender: TObject);
    procedure MicrosoftExcelSpreadSheet1Click(Sender: TObject);
    procedure WordDocument2Click(Sender: TObject);
    procedure HTMLFile2Click(Sender: TObject);
    procedure extFile1Click(Sender: TObject);
    procedure FrequencyTableToolbar1Click(Sender: TObject);
    procedure DettachFreqToolbarClick(Sender: TObject);
    procedure English2Click(Sender: TObject);
    procedure Persian2Click(Sender: TObject);
    procedure Del1Enter(Sender: TObject);
    procedure EditFmlEnter(Sender: TObject);
    procedure Add1Enter(Sender: TObject);
    procedure SGrid4ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure SGrid1ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure SGrid3ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure SGrid5ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure SGrid2ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure SGrid1TopLeftChanged(Sender: TObject);
    procedure PromptAsMenuClick(Sender: TObject);
    procedure PromptAsListClick(Sender: TObject);
    procedure ToolbarsClick(Sender: TObject);
    procedure GlobalToolbarResize(Sender: TObject);
    procedure FreqToolbarResize(Sender: TObject);
    procedure FreqControlBarResize(Sender: TObject);
    procedure S4EditFastKeyPress(Sender: TObject; var Key: Char);
    procedure S4AddFastbtnClick(Sender: TObject);
    procedure S4EditFastChange(Sender: TObject);
    procedure S1EditFastChange(Sender: TObject);
    procedure S1EditFastKeyPress(Sender: TObject; var Key: Char);
    procedure S1AddFastbtnClick(Sender: TObject);
    procedure S3EditFastChange(Sender: TObject);
    procedure S3EditFastKeyPress(Sender: TObject; var Key: Char);
    procedure S3AddFastbtnClick(Sender: TObject);
    procedure S5EditFastChange(Sender: TObject);
    procedure S5EditFastKeyPress(Sender: TObject; var Key: Char);
    procedure S5AddFastbtnClick(Sender: TObject);
    procedure S2EditFastChange(Sender: TObject);
    procedure S2EditFastKeyPress(Sender: TObject; var Key: Char);
    procedure S2AddFastbtnClick(Sender: TObject);
  private
    { Private declarations }
    GridImagePath:array[1..5] of String;
    ChartColors:array of TColor;
    Changed:Boolean;
    AnimIndex1:Integer;
    GridLAstPoint:array[1..5] of TPoint;
    procedure MoveCombo(SG:TStringGrid;var Combo:TComboBox;ACol,ARow:Integer);
    procedure EmptyLists(FieldID:Byte);
    procedure ResetGridField(FieldID:Integer;Row:Integer;Fill:Boolean);
    procedure ResetSheet(var Sheet:SheetSettings;LIndex:Integer);
    procedure UpdateSheet(FieldID:Byte);
    procedure LoadSheet(FieldID:Byte);
    procedure GridSelectCell(Sheet:SheetSettings;ID:Byte;Grid:TStringGrid;Combo:TComboBox;NLabel:TStaticText;ACol,ARow:Integer);
    procedure PrepareSheetGrid(Sheet:SheetSettings;Grid:TStringGrid;ID:Byte;SRow:Integer;FieldLabel,TypeLabel,DesLabel,FilterLabel,NLabel:TStaticText;Combo:TComboBox;Filterbtn:TBitBtn;PSpin:TSpinEdit);
    procedure DeSelectRow;
    procedure SortValueList(Mode:TSortingMode);
    procedure ShowModeList(Visible:Boolean);
    procedure AnalyzeData;
    procedure SetChartDefaults;
    procedure ReadConfiguration;
    procedure WriteConfiguration;
    procedure ArrangeCheckBoxes;
    procedure ForceSheetLimitations(Sheet:SheetSettings;SGrid:TStringGrid);
    procedure SetSheetDefaults;
    procedure GotoActiveSheet;
    procedure LoadWork(DatList:TListBox;const Rec:array of WorkSettings);
    procedure ReadSheetList(SheetID:Byte;DatList:TListBox;var Index:Integer;Count:Integer);
    procedure LoadGridFromList(DatList:TListBox;Grid:TStringGrid;var Index:Integer;Count:Integer);
    procedure ApplySheetSettings(SheetID:Byte;var Sheet:SheetSettings);
    procedure SaveWorkToFile(const FileName:String);
    procedure WriteSheetList(SheetID:Byte;List:TListBox);
    procedure AppendGridToList(List:TListBox;Grid:TStringGrid;var Count:Integer);
    procedure GridMouseUp(Grid:TStringGrid;P1,P2:TPopupMenu;X,Y:Integer);
    procedure ResizeForm;
    procedure GotoNextSection;
    procedure GotoPreviousSection;
    procedure SetMenuOwnerDraw;
    function ValidateChecks:Boolean;
    function GetDefaultValue(Row:Byte):String;
  public
    IsMinimized:Boolean;
    StartX, StartY, StartW, StartH: Integer;
    NeedChartRefresh:Boolean;
    LastFilter:array[1..5] of FilteredCell;
    LockedCells:array[1..5] of array of TLockedCell;
    ChartHistoryCount:Byte;
    procedure FillComboSpecific(var Combo:TComboBox;typeID:Byte);
    procedure ExtractSpin(var LI,UI:Integer;S:String);overload;
    procedure ExtractSpin(var LI,UI:Extended;S:String);overload;
    procedure SaveGridToFile(FName:String);
    procedure LoadGridFromFile(FName:String;AppendToGrid:Boolean;StCol:Integer=1;StRow:Integer=1);
    procedure FilterGridValues(Grid:TStringGrid;FX:String;Precision:Byte;IsDec:Boolean);
    procedure ResizeGrid(GridResizeKind:TGridResizeKind;Grid:TStringGrid);
    procedure AppendDataSheet(FromGrid:TStringGrid);
    procedure ResetChart(Chart:TChart);
    procedure FillValueList(var SGrid:TAdvStringGrid);
    procedure FillRangeList(Names:Boolean=True);
    procedure CreateTable;
    procedure CreateFrqTableColumns;
    procedure MoveColorSelector(SGrid:TStringGrid;ColorSelector:TColorSelector;ACol,ARow:Integer);
    procedure ClearGrid(Grid:TStringGrid);
    procedure RefreshSheetGrid(SG:TStringGrid);
    procedure ViewAsChart(SG:TAdvStringGrid;ViewItem:TMenuItem);
    procedure VisibleMenuHeaders(V:Boolean);
    procedure AutosizeTabelColumn;
    procedure AutoSizeGridRows(SGrid:TStringGrid);
    procedure OpenWorkProject(const FName:String);
    procedure ResetProgram;
    procedure SetHelpTopic(Language:LanguageID);
    procedure AddFastInput(SGrid:TAdvStringGrid;SheetID:Byte;const Input:String);
    function CanEditCell(ACol,ARow:Integer;ASheet:Byte):Boolean;
    function GetSGrid(No:Byte):TAdvStringGrid;
    function GetSheet(No:Byte):SheetSettings;
    function CheckGridCell(Grid:TStringGrid;TypeLabel:TStaticText):Boolean;
    function IsValidInt(S:String):Boolean;
    function IsValidDec(S:String):Boolean;
    function StrToExt(S:String):Extended;
    function IsValidSpan(Span:String):Boolean;
    function GetXLSFileName(var FName:String):Boolean;
    function GetWordDOCFileName(var FName:String):Boolean;
    function GetHTMLFileName(var FName:String):Boolean;
    { Public declarations }
  end;

  CheckSpanType=record
      Check:Boolean;
      Col,Row:Integer;
  end;
  RangeType=record
      LBound:Longint;
      UBound:Longint;
  end;
  ConfigRecord=record
    RowHeight:array[0..5] of Integer;
    ViewChartToolbar:Boolean;
    ViewGlobalToolbar:Boolean;
    ViewFastCalc:Boolean;
    ViewStatusBar:Boolean;
    HideWhenMinimized:Boolean;
    G1Style,G2Style,G3Style,G4Style,G5Style:Byte;
    G1Font,G2Font,G3Font,G4Font,G5Font:TFontSettings;
    ShowTipOfTheDay:Boolean;
    DefaultHelp:Byte;  //0:English 1:Persian
  end;

var
  MainForm: TMainForm;
  AlertCount:Byte;
  AlertString:String;
  NewAlert:Boolean;
  AlertIndex:Byte;
  Sheet1,Sheet2,Sheet3,Sheet4,Sheet5:SheetSettings;
  ActiveField:Byte=1;
  CurrentType:Byte=1;
  Checking:Boolean=False;
  ValueCount:array[1..5] of Integer;
  LastUsedRow:Integer;
  HHP,HP:Integer;
  RangeList:array[1..600] of RangeType;
  RangeNameList:array[1..600] of String;
  RangeCount:Integer;
  ValueList:array[1..1500] of Extended;
  VListCount:Integer=0;
  MaxValue,MinValue:Extended;
  ModeListVisible:Boolean=False;
  Memo:TMemo;
  FGrid:TStringGrid;
  CanCreateAChart:Boolean=False;
  ActiveSheet:Byte=0;
  RowSizing:Boolean=False;
  ExeDir:String;
  HelpTopic:Integer;
  RebuildFreqTable:Boolean=True;
  TableIsQualitative:Boolean=False;
  WorkFileName:String;

implementation

uses FormulaUnit, TableUnit, ChartFormUnit, MainAboutUnit,
  HelpLangFormUnit, GridFindUnit, GridReplaceUnit,Procs, TipUnit,
  ChartPrevUnit, MemberUnit;
const SC_AboutItem=WM_USER+1;

{$R *.dfm}


procedure TMainForm.ForceSheetLimitations(Sheet:SheetSettings;SGrid:TStringGrid);
var i,j:Integer;
begin
  if Sheet.TypeIndex=3 then
    for i:=1 to (SGrid.RowCount-2) do
      for j:=1 to (SGrid.ColCount-1) do
        if Length(SGrid.Cells[j,i])>0 then
        begin
          case Sheet.CharStyle of
            csUppercase: SGrid.Cells[j,i]:=UpperCase(SGrid.Cells[j,i]);
            csLowercase: SGrid.Cells[j,i]:=LowerCase(SGrid.Cells[j,i]);
          end;
          if Length(SGrid.Cells[j,i])>Sheet.MaxStringLength then
            SGrid.Cells[j,i]:=Copy(SGrid.Cells[j,i],1,Sheet.MaxStringLength);
        end;
end;

procedure TMainForm.SetHelpTopic(Language:LanguageID);
begin
  case ActiveSheet of
    0:HelpTopic:=30;
    1,2,3,4,5:HelpTopic:=40;
    6:HelpTopic:=50;
    7:HelpTopic:=60;
    8:HelpTopic:=70;
    9:HelpTopic:=120;
    10:HelpTopic:=130;
  end;
end;

procedure TMainForm.ArrangeCheckBoxes;
begin
  S1Check.Top:=SettingGrid.Top+SettingGrid.RowHeights[0]+SettingGrid.GridLineWidth*2+Round(((SettingGrid.RowHeights[1] / 2)-(S2check.Height/2)));
  S2Check.Top:=SettingGrid.Top+SettingGrid.RowHeights[0]+SettingGrid.RowHeights[1]+SettingGrid.GridLineWidth*3+Round(((SettingGrid.RowHeights[2] / 2)-(S2check.Height/2)));
  S3Check.Top:=SettingGrid.Top+SettingGrid.RowHeights[0]+SettingGrid.RowHeights[1]+SettingGrid.RowHeights[2]+SettingGrid.GridLineWidth*4+Round(((SettingGrid.RowHeights[3] / 2)-(S3check.Height/2)));
  S4Check.Top:=SettingGrid.Top+SettingGrid.RowHeights[0]+SettingGrid.RowHeights[1]+SettingGrid.RowHeights[2]+SettingGrid.RowHeights[3]+SettingGrid.GridLineWidth*5+Round(((SettingGrid.RowHeights[4] / 2)-(S1check.Height/2)));
  S5Check.Top:=SettingGrid.Top+SettingGrid.RowHeights[0]+SettingGrid.RowHeights[1]+SettingGrid.RowHeights[2]+SettingGrid.RowHeights[3]+SettingGrid.RowHeights[4]+SettingGrid.GridLineWidth*6+Round(((SettingGrid.RowHeights[5] / 2)-(S1check.Height/2)));
end;

procedure TMainForm.WriteConfiguration;

function SetGridStyle(SG:TAdvStringGrid):Byte;
begin
  if SG.Look=glXP then
    Result:=1
  else if SG.Look=glClassic then
    Result:=2
  else if SG.Look=glSoft then
    Result:=3;
end;

function GetGridFontSettings(SG:TAdvStringGrid):TFontSettings;
begin
  Result.Charset:=SG.Font.Charset;
  Result.Color:=SG.Font.Color;
  Result.Height:=SG.Font.Height;
  Result.Pitch:=SG.Font.Pitch;
  Result.Size:=SG.Font.Size;
  Result.Style:=SG.Font.Style;
end;

var F:File of ConfigRecord;
    Cfg:^ConfigRecord;
    i:Integer;
begin
  Cfg:=AllocMem(SizeOf(ConfigRecord));
  for i:=0 to 5 do
    Cfg^.RowHeight[i]:=SettingGrid.RowHeights[i];
  Cfg^.ViewChartToolbar:=ChartToolbar1.Checked;
  Cfg^.ViewGlobalToolbar:=GlobalToolbar1.Checked;
  Cfg^.ViewStatusBar:=StatusBar1.Checked;
  Cfg^.ViewFastCalc:=FastCalc1.Checked;
  Cfg^.HideWhenMinimized:=HideWhenMinimized1.Checked;
  Cfg^.G1Style:=SetGridStyle(SGrid1);
  Cfg^.G2Style:=SetGridStyle(SGrid2);
  Cfg^.G3Style:=SetGridStyle(SGrid3);
  Cfg^.G4Style:=SetGridStyle(SGrid4);
  Cfg^.G5Style:=SetGridStyle(SGrid5);

  Cfg^.G1Font:=GetGridFontSettings(SGrid1);
  Cfg^.G2Font:=GetGridFontSettings(SGrid2);
  Cfg^.G3Font:=GetGridFontSettings(SGrid3);
  Cfg^.G4Font:=GetGridFontSettings(SGrid4);
  Cfg^.G5Font:=GetGridFontSettings(SGrid5);
  if TipOfTheDay.Tag=1 then
    Cfg^.ShowTipOfTheDay:=True
  else
    Cfg^.ShowTipOfTheDay:=False;  
  if EnglishHelp.Checked then
    Cfg^.DefaultHelp:=0
  else if PersianHelp.Checked then
    Cfg^.DefaultHelp:=1;
  try
    AssignFile(F,'StatMan.cfg');
    Rewrite(F);
    Write(F,Cfg^);
    CloseFile(F);
    FreeMem(Cfg);
    Cfg:=nil;
  except
    if Cfg<>nil then
      FreeMem(Cfg);
  end;
end;

procedure TMainForm.ReadConfiguration;

procedure SetGridStyle(SDefault,SClassic,SFlat:TMenuItem;Style:Byte);
begin
  case Style of
    1: begin
         SDefault.Click;
         SDefault.Checked:=True;
       end;
    2: begin
         SClassic.Click;
         SClassic.Checked:=True;
       end;
    3: begin
         SFlat.Click;
         SFlat.Checked:=True;
       end;
  end;
end;

procedure SetGridFont(SG:TAdvStringGrid;const F:TFontSettings);
begin
  SG.Font.Charset:=F.Charset;
  SG.Font.Style:=F.Style;
  SG.Font.Height:=F.Height;
  SG.Font.Color:=F.Color;
  SG.Font.Pitch:=F.Pitch;
  SG.Font.Size:=F.Size;
end;

var F:File of ConfigRecord;
    Cfg:^ConfigRecord;
    i:Integer;
begin
  if FileExists('StatMan.cfg') then
  begin
    Cfg:=AllocMem(SizeOf(ConfigRecord));
    try
      AssignFile(F,'StatMan.cfg');
      Reset(F);
      Read(F,Cfg^);
      CloseFile(F);
      for i:=0 to 5 do
        SettingGrid.RowHeights[i]:=Cfg^.RowHeight[i];
      ChartToolbar1.Checked:=Cfg^.ViewChartToolbar;
      ChartToolbar.Visible:=Cfg^.ViewChartToolbar;
      GlobalToolbar1.Checked:=Cfg^.ViewGlobalToolbar;
      GlobalToolbar.Visible:=Cfg^.ViewGlobalToolbar;
      FastCalc1.Checked:=Cfg^.ViewFastCalc;
      FastCalcToolbar.Visible:=Cfg^.ViewFastCalc;
      StatusBar1.Checked:=Cfg^.ViewStatusBar;
      StatusBar.Visible:=Cfg^.ViewStatusBar;
      HideWhenMinimized1.Checked:=Cfg^.HideWhenMinimized;
      TTIcon.MinimizeToTray:=Cfg^.HideWhenMinimized;
      SetGridStyle(G1SDefault,G1SClassic,G1SFlat,Cfg^.G1Style);
      SetGridStyle(G2SDefault,G2SClassic,G2SFlat,Cfg^.G2Style);
      SetGridStyle(G3SDefault,G3SClassic,G3SFlat,Cfg^.G3Style);
      SetGridStyle(G4SDefault,G4SClassic,G4SFlat,Cfg^.G4Style);
      SetGridStyle(G5SDefault,G5SClassic,G5SFlat,Cfg^.G5Style);

      SetGridFont(SGrid1,Cfg^.G1Font);
      SetGridFont(SGrid2,Cfg^.G2Font);
      SetGridFont(SGrid3,Cfg^.G3Font);
      SetGridFont(SGrid4,Cfg^.G4Font);
      SetGridFont(SGrid5,Cfg^.G5Font);
      if Cfg^.ShowTipOfTheDay then
        TipOfTheDay.Tag:=1
      else
        TipOfTheDay.Tag:=0;  
      if Cfg^.DefaultHelp=0 then
      begin
        EnglishHelp.Checked:=True;
        PersianHelp.Checked:=False;
      end
      else if Cfg^.DefaultHelp=1 then
      begin
        PersianHelp.Checked:=True;
        EnglishHelp.Checked:=False;
      end;
      FreeMem(Cfg);
      Cfg:=nil;
    except
      ShowMessage('Unable to read StatMan config file.');
      if Cfg<>nil then
        FreeMem(Cfg);
    end;
  end;
end;

procedure TMainForm.AppendDataSheet(FromGrid:TStringGrid);
var Row,Col,i,j:Integer;
begin
  RefreshSheetGrid(FGrid);
  for Row:=1 to (FGrid.RowCount-2) do
  begin
    for Col:=1 to (FGrid.ColCount-1) do
    begin
      if Length(FGrid.Cells[Col,Row])=0 then
        Break;
    end;
    if (Col<FGrid.ColCount) and (Length(FGrid.Cells[Col,Row])=0) then
      Break;
  end;
  if (Row=FGrid.RowCount-1) and (Col=FGrid.ColCount) then
    Exit;
  Dec(Col);
  for i:=1 to (FromGrid.RowCount-2) do
    for j:=1 to (FromGrid.ColCount-1) do
      if Length(FromGrid.Cells[j,i])>0 then
      begin
        Inc(Col);
        if Col=Fgrid.ColCount then
        begin
          Inc(Row);
          Col:=1;
        end;
        if Row=(Fgrid.RowCount-1) then
          Exit;
        FGrid.Cells[Col,Row]:=FromGrid.Cells[j,i];
      end;
end;

procedure TMainForm.ResizeGrid(GridResizeKind:TGridResizeKind;Grid:TStringGrid);
var MaxColWidth,MinColWidth,i:Integer;
begin
  MaxColWidth:=0;
  MinColWidth:=Grid.ColWidths[1];
  for i:=1 to (Grid.ColCount-1) do
  begin
    if Grid.ColWidths[i]>MaxColWidth then
      MaxColWidth:=Grid.ColWidths[i];
    if Grid.ColWidths[i]<MinColWidth then
      MinColWidth:=Grid.ColWidths[i];
  end;
  if GridResizeKind=grkGreatest then
    Grid.DefaultColWidth:=MaxColWidth
  else if GridResizeKind=grkSmallest then
    Grid.DefaultColWidth:=MinColWidth;
  GRid.ColWidths[0]:=22;
end;

procedure TMainForm.SetChartDefaults;
begin
  Chart1.View3DOptions.Orthogonal:=True;
  Chart1.Foot.Font.Color:=clRed;
  bwidth.ItemIndex:=0;
  bstyle.ItemIndex:=0;
  SizeCombo.ItemIndex:=SizeCombo.Items.IndexOf(IntToStr(ChartTitleEdit.Font.Size));
  bwidthf.ItemIndex:=0;
  bstylef.ItemIndex:=0;
  SizeCombo2.ItemIndex:=SizeCombo2.Items.IndexOf(IntToStr(ChartFooterEdit.Font.Size));
  LegendPos.ItemIndex:=1;
  LegendStyle.ItemIndex:=0;
  LegendStyle.Hint:=LegendStyle.Items.Strings[LegendStyle.ItemIndex];
  iwidth.ItemIndex:=0;
  istyle.ItemIndex:=0;
  penwidthcombo3.ItemIndex:=0;
  penstylecombo2.ItemIndex:=0;
  penwidthcombo4.ItemIndex:=0;
  penstylecombo3.ItemIndex:=0;
  SizeCombo3.ItemIndex:=SizeCombo3.Items.IndexOf(IntToStr(AxisTitleEdit.Font.Size));
  MWidth.ItemIndex:=0;
  MStyle.ItemIndex:=0;
  IWidth.ItemIndex:=0;
  IStyle.ItemIndex:=0;
  iwidth2.ItemIndex:=0;
  istyle2.ItemIndex:=0;
  miWidth.ItemIndex:=0;
  miStyle.ItemIndex:=0;
  axwidth.ItemIndex:=0;
  axstyle.ItemIndex:=0;
  gridwidth.ItemIndex:=0;
  gridstyle.ItemIndex:=2;
  gridwidthv.ItemIndex:=0;
  gridstylev.ItemIndex:=2;
  ChartGrid.Cells[1,0]:='Names';
  ChartGrid.Cells[2,0]:='Values';
  ColorGrid.Cells[0,0]:='Colors';
end;

function TMainForm.GetDefaultValue(Row:Byte):String;
begin
  if SettingGrid.Cells[2,Row]='Integer' then
    Result:='0'
  else if SettingGrid.Cells[2,Row]='Decimal' then
    Result:='0.0'
  else if SettingGrid.Cells[2,Row]='String' then
    Result:='Str1'
  else if SettingGrid.Cells[2,Row]='Span' then
    Result:='[0,1]';
end;

procedure TMainForm.FilterGridValues(Grid:TStringGrid;FX:String;Precision:Byte;IsDec:Boolean);
var i,j,k:Integer; CanAdd:Boolean; S,SubS:String;
begin
  FilterParser.Expression:=FX;
  for i:=1 to (Grid.ColCount-1) do
    for j:=1 to (Grid.RowCount-2) do
      if Length(Grid.Cells[i,j])>0 then
      begin
        CanAdd:=True;
        S:=Grid.Cells[i,j];
        if IsDec then
          FilterParser.X:=StrToFloat(S)
        else
          FilterParser.X:=StrToInt(S);
        S:=FloatToStr(FilterParser.Value);
        if (Pos('E',S)>0) or (Pos('e',S)>0) then
          CanAdd:=False;
        if not(CanAdd) then Continue;
        SubS:='';
        if Pos('.',S)>0 then
        begin
          SubS:=Copy(S,Pos('.',S)+1,Length(S)-Pos('.',S));
          S:=Copy(S,1,Pos('.',S)-1);
          if Length(SubS)>Precision then
            SubS:=Copy(Subs,1,Precision);
        end;
        if Precision=0 then
        begin
          Grid.Cells[i,j]:=S;
          Continue;
        end;
        if IsDec then
          S:=S+'.'+SubS
        else
        begin
          for k:=1 to Length(SubS) do
            if SubS[k]<>'0' then
              CanAdd:=False;
        end;
        if CanAdd then
          Grid.Cells[i,j]:=S;
      end;
end;

//***********************************************************
procedure TMainForm.LoadGridFromFile(FName:String;AppendToGrid:Boolean;StCol:Integer=1;StRow:Integer=1);
var i,j,Col,Row,Count:Integer;
    F:TextFile;
    S:String;
begin
  AssignFile(F,FName);
  {$I-}
  Reset(F);
  {$I+}
  if IOResult<>0 then
  begin
    ShowMessage('Occured an error while reading from the specified file');
    Exit;
  end;
  if not(AppendToGrid) then
    for i:=1 to (Fgrid.ColCount-1) do
      for j:=1 to (Fgrid.RowCount-2) do
        Fgrid.Cells[i,j]:='';
  Col:=StCol; Row:=StRow;
  Dec(Col);
  while not(eof(f)) do
  begin
    Inc(Col);
    if Col=Fgrid.ColCount then
    begin
      Inc(Row);
      Col:=1;
    end;
    if Row=(Fgrid.RowCount-1) then
    begin
      CloseFile(F);
      Exit;
    end;
    ReadLn(F,S);
    Fgrid.Cells[Col,Row]:=S;
  end;
  CloseFile(F);
  Count:=0;
  for i:=1 to (FGrid.ColCount-1) do
    for j:=1 to (FGrid.RowCount-2) do
      if Length(FGrid.Cells[i,j])>0 then
        Inc(Count);
  if FGrid=SGrid1 then
  begin
    ValueCount[1]:=Count; S1NLabel.Caption:=IntToStr(Count);
  end
  else if FGrid=SGrid2 then
  begin
    ValueCount[2]:=Count; S2NLabel.Caption:=IntToStr(Count);
  end
  else if FGrid=SGrid3 then
  begin
    ValueCount[3]:=Count; S3NLabel.Caption:=IntToStr(Count);
  end
  else if FGrid=SGrid4 then
  begin
    ValueCount[4]:=Count; S4NLabel.Caption:=IntToStr(Count);
  end
  else if FGrid=SGrid5 then
  begin
     ValueCount[5]:=Count; S5NLabel.Caption:=IntToStr(Count);
  end;
end;

//********************************************************************
procedure TMainForm.SaveGridToFile(FName:String);
var i,j:Integer;
    F:TextFile;
    S:String;
begin
  AssignFile(F,FName);
  {$I-}
  Rewrite(F);
  {$I+}
  if IOResult<>0 then
  begin
    ShowMessage('Occured an error while writing in the specified file');
    Exit;
  end;
  for i:=1 to (Fgrid.RowCount-2) do
    for j:=1 to (Fgrid.ColCount-1) do
    begin
      S:=Fgrid.Cells[j,i];
      if Length(S)>0 then
        WriteLn(F,S);
    end;
  CloseFile(F);
end;

//********************************************************************
procedure TMainForm.AnalyzeData;
var Mean,Middle,Mode,AD,Variance,SD,CV,Value:Double;
    n,MaxFreq,MaxFreqIndex,ModeCount,i,j:Integer;  FreqList,ModeTemp:array of Integer;
    Str:String;
begin
  Mean:=0;
  n:=VListCount;
  for i:=1 to n do
    Mean:=Mean+(1/n)*ValueList[i];
  MeanLabel.Caption:=FloatToStr(Mean);
  SortValueList(smAscending);
  if (n MOD 2)=0 then
    Middle:=(ValueList[Round(n/2)]+ValueList[Round(n/2)+1])/2
  else
    Middle:=ValueList[Round((n+1)/2)];
  MiddleLabel.Caption:=FloatToStr(Middle);
  SetLength(FreqList,n);
  for i:=1 to n do
  begin
    Value:=ValueList[i];
    j:=i;
    if (i-1)>0 then
      for j:=1 to (i-1) do
        if ValueList[j]=Value then
        begin
          FreqList[i-1]:=FreqList[j-1];
          Break;
        end;
    if j<>i then
      Continue;
    FreqList[i-1]:=0;
    for j:=1 to n do
      if ValueList[j]=Value then
        Inc(FreqList[i-1]);
  end;
  MaxFreq:=FreqList[0];
  MaxFreqIndex:=0;
  for i:=1 to (n-1) do
    if FreqList[i]>MaxFreq then
    begin
      MaxFreq:=FreqList[i];
      MaxFreqIndex:=i;
    end;
  ModeCount:=0;
  for i:=0 to (n-1) do
    if FreqList[i]=MaxFreq then
    begin
      j:=ModeCount;
      if ModeCount>0 then
        for j:=0 to (ModeCount-1) do
          if ValueList[ModeTemp[j]]=ValueList[i+1] then
            Break;
      if j<>ModeCount then
        Continue;
      Inc(ModeCount);
      SetLength(ModeTemp,ModeCount);
      ModeTemp[ModeCount-1]:=i+1;
    end;
  ModeList.Visible:=False;ShowModeList(False);
  ModeBtn.Visible:=False;
  if ModeCount=1 then
  begin
    Mode:=ValueList[MaxFreqIndex+1];
    ModeLabel.Caption:=FloatToStr(Mode);
  end
  else
  begin
    ModeList.Items.Clear;
    for i:=0 to (ModeCount-1) do
      ModeList.Items.Append(FloatToStr(ValueList[ModeTemp[i]]));
    ModeLabel.Caption:='{Multiple Modes}';
    ModeBtn.Visible:=True;
  end;
  SetLength(FreqList,0);
  SetLength(ModeTemp,0);
  AD:=0;
  Variance:=0;
  for i:=1 to n do
  begin
    AD:=AD+Abs(ValueList[i]-Mean)/n;
    Variance:=Variance+Power(ValueList[i]-Mean,2)/(n-1);
  end;
  ADLabel.Caption:=FloatToStr(AD);
  VarLabel.Caption:=FloatToStr(Variance);
  SD:=sqrt(Variance);
  SDLabel.Caption:=FloatToStr(SD);
  if Mean<>0 then
    CV:=(SD/Mean)*100
  else
    CV:=0;
  Str:=FloatToStr(CV);
  if Pos('.',Str)>0 then
    Str:=Copy(Str,1,Pos('.',Str)+2);
  CVLabel.Caption:=Str+' %';
end;

//********************************************************************
procedure TMainForm.ShowModeList(Visible:Boolean);
begin
  if Visible and not(ModeList.Visible) then
  begin
    ModeBtn.Glyph:=ModeBtnDown.Picture.Bitmap;
    ModeList.Visible:=True;
    ModeListVisible:=True;
  end
  else if not(Visible) then
  begin
    ModeBtn.Glyph:=ModeBtnUp.Picture.Bitmap;
    ModeListVisible:=False;
  end;
  ModeListTimer.Enabled:=True;
end;

//********************************************************************
procedure TMainForm.SortValueList(Mode:TSortingMode);
var VIndex,i,j:Integer;  temp:Double;
begin
  if Mode=smAscending then
  begin
    for i:=1 to (VListCount-1) do
    begin
      VIndex:=i;
      for j:=(i+1) to VListCount do
        if ValueList[VIndex]>ValueList[j] then
          VIndex:=j;
      if VIndex<>i then
      begin
        temp:=ValueList[i];
        ValueList[i]:=ValueList[VIndex];
        ValueList[VIndex]:=temp;
      end;
    end;
  end
  else if Mode=smDescending then
  begin
    for i:=1 to (VListCount-1) do
    begin
      VIndex:=i;
      for j:=(i+1) to VListCount do
        if ValueList[VIndex]<ValueList[j] then
          VIndex:=j;
      if VIndex<>i then
      begin
        temp:=ValueList[i];
        ValueList[i]:=ValueList[VIndex];
        ValueList[VIndex]:=temp;
      end;
    end;
  end;
end;

//********************************************************************
procedure TMainForm.FillRangeList(Names:Boolean=True);
var I,K,BaseNum,Count:Integer; R:Extended;
begin
  SetRoundMode(rmDown);
  MaxValue:=ValueList[1];
  MinValue:=ValueList[1];
  for Count:=1 to VListCount do
  begin
    if ValueList[Count]>MaxValue then
      MaxValue:=ValueList[Count];
    if ValueList[Count]<MinValue then
      MinValue:=ValueList[Count];
  end;
  if MaxValue=MinValue then
  begin
    RangeCount:=1;
    RangeList[1].LBound:=Round(MinValue);
    RangeList[1].UBound:=RangeList[1].LBound+1;
    if Names then
      RangeNameList[1]:='['+IntToStr(RangeList[1].LBound)+','+IntToStr(RangeList[1].UBound)+')';
    Exit;
  end;
  R:=MaxValue-MinValue;
  SetRoundMode(rmDown);
  K:=Round(sqrt(VListCount));
  if K<>sqrt(VListCount) then
    K:=K+1;
//  ShowMessage('K = '+IntToStr(K));
  I:=Round(R/K)+1;
  BaseNum:=Round(MinValue);
  RangeCount:=K;
  for Count:=1 to K do
  begin
    RangeList[Count].LBound:=BaseNum;
    BaseNum:=BaseNum+I;
    RangeList[Count].UBound:=BaseNum;
  end;
  if Names then
    for Count:=1 to RangeCount do
      RangeNameList[count]:='['+IntToStr(RangeList[Count].LBound)+','+IntTostr(RangeList[Count].UBound)+')';
end;

//********************************************************************
procedure TMainForm.FillValueList(var SGrid:TAdvStringGrid);
var i,j:Integer; S:String;
begin
  VListCount:=0;
  for i:=1 to (SGrid.RowCount-2) do
    for j:=1 to (SGrid.ColCount-1) do
    begin
      S:=SGrid.Cells[j,i];
      if Length(S)>0 then
      begin
        Inc(VListCount);
        ValueList[VListCount]:=StrToFloat(SGrid.Cells[j,i]);
      end;
    end;
end;

//********************************************************************
procedure TMainForm.FillComboSpecific(var Combo:TComboBox;typeID:Byte);
begin
  if (Sheet1.Used) and (Sheet1.TypeIndex=typeID) then
    Combo.Items.Append('Data Sheet 1');
  if (Sheet2.Used) and (Sheet2.TypeIndex=typeID) then
    Combo.Items.Append('Data Sheet 2');
  if (Sheet3.Used) and (Sheet3.TypeIndex=typeID) then
    Combo.Items.Append('Data Sheet 3');
  if (Sheet4.Used) and (Sheet4.TypeIndex=typeID) then
    Combo.Items.Append('Data Sheet 4');
  if (Sheet5.Used) and (Sheet5.TypeIndex=typeID) then
    Combo.Items.Append('Data Sheet 5');
end;

//********************************************************************
procedure TMainForm.CreateTable;
var fi,Total,i,j:Integer;
begin
  TableGrid.Cells[0,0]:='Limits';
  TableGrid.RowCount:=RangeCount+2;
  TableGrid.Cells[1,0]:='fi';
  TableGrid.Cells[2,0]:='fpi';
  TableGrid.Cells[3,0]:='Pi';
  TableGrid.Cells[4,0]:='Fi';
  TableGrid.Cells[5,0]:='Fpi';
  TableGrid.Cells[6,0]:='Pci';
  TableGrid.Cells[0,TableGrid.RowCount-1]:='Total';
  Total:=0;
  for i:=1 to RangeCount do
  begin
    fi:=0;
    for j:=1 to VListCount do
      if (ValueList[j]>=RangeList[i].LBound) and (ValueList[j]<RangeList[i].UBound) then
        Inc(fi);
    TableGrid.Cells[0,i]:=RangeNameList[i];
    TableGrid.Cells[1,i]:=IntToStr(fi);
    Total:=Total+fi;
    TableGrid.Cells[4,i]:=IntToStr(Total);
  end;
  TableGrid.Cells[1,TableGrid.RowCount-1]:=IntToStr(Total);
  if Total=0 then
  begin
    for i:=1 to (TableGrid.ColCount-1) do
      for j:=1 to (TableGrid.RowCount-2) do
        TableGrid.Cells[i,j]:='0';
    for i:=2 to (TableGrid.ColCount-1) do
      TableGrid.Cells[i,TableGrid.RowCount-1]:='********';
    ShowMessage('Total number of this frequency table is ZERO.');
    Exit;
  end;
  CreateFrqTableColumns;
end;

procedure TMainForm.DeSelectRow;
var Row:Integer;
begin
  if S1Check.Checked then
    Row:=1
  else if S2Check.Checked then
    Row:=2
  else if S3Check.Checked then
    Row:=3
  else if S4Check.Checked then
    Row:=4
  else if S5Check.Checked then
    Row:=5;
  SettingGrid.Row:=Row;
end;

procedure TMainForm.ClearGrid(Grid:TStringGrid);
var i,j:Integer;
begin
  for i:=1 to (Grid.RowCount-1) do
    for j:=1 to (Grid.ColCount-1) do
      Grid.Cells[j,i]:='';
end;

procedure TMainForm.PrepareSheetGrid(Sheet:SheetSettings;Grid:TStringGrid;ID:Byte;SRow:Integer;FieldLabel,TypeLabel,DesLabel,FilterLabel,NLabel:TStaticText;Combo:TComboBox;Filterbtn:TBitBtn;PSpin:TSpinEdit);
var i,j:Integer;
begin
  LastUsedRow:=0;
  if SettingPanel.Visible then
  begin
//    UpdateSheet(ActiveField);
//    Cleared:=False;
    if SettingGrid.Cells[2,SRow]<>TypeLabel.Caption then
    begin
      TypeLabel.Caption:=SettingGrid.Cells[2,SRow];
      ClearGrid(Grid);
//      Cleared:=True;
    end;
    FieldLabel.Caption:=Sheet.FieldName;
    if Length(SettingGrid.Cells[4,SRow])>0 then
    begin
      DesLabel.Caption:=SettingGrid.Cells[4,SRow];
      DesLabel.Font.Style:=DesLabel.Font.Style-[fsItalic];
      DesLabel.Font.Style:=DesLabel.Font.Style+[fsBold];
    end
    else
    begin
      DesLabel.Caption:='{No Description}';
      DesLabel.Font.Style:=DesLabel.Font.Style+[fsItalic];
      DesLabel.Font.Style:=DesLabel.Font.Style-[fsBold];
    end;
    if Sheet.Filter then
    begin
      FilterLabel.Caption:=StrTemp.Items.Strings[Sheet.FilterIndex];
      FilterLabel.Font.Style:=FilterLabel.Font.Style-[fsItalic];
      FilterLabel.Font.Style:=FilterLabel.Font.Style+[fsBold];
      Filterbtn.Enabled:=True;
      PSpin.Enabled:=True;
    end
    else
    begin
      FilterLabel.Caption:='{No  Filter}';
      FilterLabel.Font.Style:=FilterLabel.Font.Style+[fsItalic];
      FilterLabel.Font.Style:=FilterLabel.Font.Style-[fsBold];
      Filterbtn.Enabled:=False;
      PSpin.Enabled:=False;
    end;
    if Sheet.ForceValue then
      Grid.Options:=Grid.Options-[goEditing]
    else
      Grid.Options:=Grid.Options+[goEditing];
  end;
  ValueCount[ID]:=0;
  for i:=1 to (Grid.RowCount-2) do
    for j:=1 to (Grid.ColCount-1) do
      if Grid.Cells[j,i]<>'' then
        Inc(ValueCount[ID]);
  NLabel.Caption:=IntToStr(ValueCount[ID]);
end;

procedure TMainForm.GridSelectCell(Sheet:SheetSettings;ID:Byte;Grid:TStringGrid;Combo:TComboBox;NLabel:TStaticText;ACol,ARow:Integer);
var i,j,Count:Integer;
begin
  Count:=0;
  for i:=1 to (Grid.ColCount-1) do
    for j:=1 to (Grid.RowCount-2) do
      if Length(Grid.Cells[i,j])>0 then
        Inc(Count);
  NLabel.Caption:=IntToStr(Count);
end;

function TMainForm.IsValidSpan(Span:String):Boolean;
var L,U:String;
begin
  Result:=False;
  if pos(',',Span)<=0 then
    Exit;
  if (Span[1]<>'[') or (Span[Length(Span)]<>']') then
    Exit;
  L:=Copy(Span,2,pos(',',Span)-2);
  U:=Copy(Span,pos(',',Span)+1,Length(Span)-pos(',',Span)-1);
  if IsValidInt(L) and IsValidInt(U) then
    Result:=True;
end;

procedure TMainForm.RefreshSheetGrid(SG:TStringGrid);
var NextRow,NextCol:Integer;
    i,j:Integer;
begin
  TempGrid.RowCount:=SG.RowCount;
  TempGrid.ColCount:=SG.ColCount-1;
  NextRow:=0; NextCol:=0;
  for i:=1 to (SG.RowCount-2) do
    for j:=1 to (SG.ColCount-1) do
      if Length(SG.Cells[j,i])>0 then
      begin
        TempGrid.Cells[NextCol,NextRow]:=SG.Cells[j,i];
        SG.Cells[j,i]:='';
        if NextCol=(TempGrid.ColCount-1) then
        begin
          Inc(NextRow);
          NextCol:=0;
        end
        else
          Inc(NextCol);
      end;
  for i:=0 to (NextRow-1) do
    for j:=0 to (TempGrid.ColCount-1) do
      SG.Cells[j+1,i+1]:=TempGrid.Cells[j,i];
  for i:=0 to (NextCol-1) do
    SG.Cells[i+1,NextRow+1]:=TempGrid.Cells[i,NextRow];
end;

procedure TMainForm.ExtractSpin(var LI,UI:Integer;S:String);
var v1,v2:String; Code:Integer;
begin
  v1:=Copy(S,2,pos(',',S)-2);
  v2:=Copy(S,pos(',',S)+1,Length(S)-pos(',',S)-1);
  val(v1,LI,Code);
  val(v2,UI,Code);
end;

procedure TMainForm.ExtractSpin(var LI,UI:Extended;S:String);
var v1,v2:String; Code:Integer;
begin
  v1:=Copy(S,2,pos(',',S)-2);
  v2:=Copy(S,pos(',',S)+1,Length(S)-pos(',',S)-1);
  val(v1,LI,Code);
  val(v2,UI,Code);
end;

function TMainForm.StrToExt(S:String):Extended;
var v:Extended; Code:Integer;
begin
  val(S,v,Code);
  Result:=v;
end;

procedure TMainForm.LoadSheet(FieldID:Byte);
var S:String; i:Integer;
begin
  case FieldID of
    1: begin
        if Sheet1.TypeIndex>=3 then
        begin
          UseFilter.Checked:=False;
          UseFilter.Enabled:=False;
        end
        else
        begin
          UseFilter.Enabled:=True;
          if Sheet1.Filter then
            UseFilter.Checked:=True
          else
            UseFilter.Checked:=False;
          FXLabel.Caption:=StrTemp.Items.Strings[Sheet1.FilterIndex];
        end;
        StrLimitsTable.Enabled:=False;
        case Sheet1.TypeIndex of
          1:;
          2:;
          3: begin
               StrLimitsTable.Enabled:=True;
               StrLimitsTable.Cells[1,1]:=IntToStr(Sheet1.MaxStringLength);
               case Sheet1.CharStyle of
                 csUppercase:StrLimitsTable.Cells[1,2]:='Uppercase';
                 csLowercase:StrLimitsTable.Cells[1,2]:='Lowercase';
                 csNone:StrLimitsTable.Cells[1,2]:='None';
               end;
             end;
          4:;
        end;
        if Sheet1.ForceValue then
        begin
          ForceValue.Checked:=True;
          List1.Items.Clear;
          case Sheet1.TypeIndex of
            1: begin
                 for i:=0 to (IntList1.Items.Count-1) do
                   List1.Items.Add(IntList1.Items.Strings[i]);
               end;
            2: begin
                 for i:=0 to (DecList1.Items.Count-1) do
                   List1.Items.Add(DecList1.Items.Strings[i]);
               end;
            3: begin
                 for i:=0 to (StrList1.Items.Count-1) do
                   List1.Items.Add(StrList1.Items.Strings[i]);
               end;
            4,5: begin
                   for i:=0 to (SpanList.Items.Count-1) do
                   begin
                     S:=SpanList.Items.Strings[i];
                     if S[1]='1' then
                       List1.Items.Add(Copy(S,2,Length(S)-1));
                   end;
                 end;
          end;
        end
      else
      begin
        ForceValue.Checked:=False;
        List1.Items.Clear;
      end;
   end;
    2: begin
        if Sheet2.TypeIndex>=3 then
        begin
          UseFilter.Checked:=False;
          UseFilter.Enabled:=False;
        end
        else
        begin
          UseFilter.Enabled:=True;
          if Sheet2.Filter then
            UseFilter.Checked:=True
          else
            UseFilter.Checked:=False;
          FXLabel.Caption:=StrTemp.Items.Strings[Sheet2.FilterIndex];
        end;
        StrLimitsTable.Enabled:=False;
        case Sheet2.TypeIndex of
          1:;
          2:;
          3: begin
               StrLimitsTable.Enabled:=True;
               StrLimitsTable.Cells[1,1]:=IntToStr(Sheet2.MaxStringLength);
               case Sheet2.CharStyle of
                 csUppercase:StrLimitsTable.Cells[1,2]:='Uppercase';
                 csLowercase:StrLimitsTable.Cells[1,2]:='Lowercase';
                 csNone:StrLimitsTable.Cells[1,2]:='None';
               end;
             end;
          4:;
        end;
        if Sheet2.ForceValue then
        begin
          ForceValue.Checked:=True;
          List1.Items.Clear;
          case Sheet2.TypeIndex of
            1: begin
                 for i:=0 to (IntList2.Items.Count-1) do
                   List1.Items.Add(IntList2.Items.Strings[i]);
               end;
            2: begin
                 for i:=0 to (DecList2.Items.Count-1) do
                   List1.Items.Add(DecList2.Items.Strings[i]);
               end;
            3: begin
                 for i:=0 to (StrList2.Items.Count-1) do
                   List1.Items.Add(StrList2.Items.Strings[i]);
               end;
            4,5: begin
                   for i:=0 to (SpanList.Items.Count-1) do
                   begin
                     S:=SpanList.Items.Strings[i];
                     if S[1]='2' then
                       List1.Items.Add(Copy(S,2,Length(S)-1));
                   end;
                 end;
          end;
        end
      else
      begin
        ForceValue.Checked:=False;
        List1.Items.Clear;
      end;
   end;
    3: begin
        if Sheet3.TypeIndex>=3 then
        begin
          UseFilter.Checked:=False;
          UseFilter.Enabled:=False;
        end
        else
        begin
          UseFilter.Enabled:=True;
          if Sheet3.Filter then
            UseFilter.Checked:=True
          else
            UseFilter.Checked:=False;
          FXLabel.Caption:=StrTemp.Items.Strings[Sheet3.FilterIndex];
        end;
        StrLimitsTable.Enabled:=False;
        case Sheet3.TypeIndex of
          1:;
          2:;
          3: begin
               StrLimitsTable.Enabled:=True;
               StrLimitsTable.Cells[1,1]:=IntToStr(Sheet3.MaxStringLength);
               case Sheet3.CharStyle of
                 csUppercase:StrLimitsTable.Cells[1,2]:='Uppercase';
                 csLowercase:StrLimitsTable.Cells[1,2]:='Lowercase';
                 csNone:StrLimitsTable.Cells[1,2]:='None';
               end;
             end;
          4:;
        end;
        if Sheet3.ForceValue then
        begin
          ForceValue.Checked:=True;
          List1.Items.Clear;
          case Sheet3.TypeIndex of
            1: begin
                 for i:=0 to (IntList3.Items.Count-1) do
                   List1.Items.Add(IntList3.Items.Strings[i]);
               end;
            2: begin
                 for i:=0 to (DecList3.Items.Count-1) do
                   List1.Items.Add(DecList3.Items.Strings[i]);
               end;
            3: begin
                 for i:=0 to (StrList3.Items.Count-1) do
                   List1.Items.Add(StrList3.Items.Strings[i]);
               end;
            4,5: begin
                   for i:=0 to (SpanList.Items.Count-1) do
                   begin
                     S:=SpanList.Items.Strings[i];
                     if S[1]='3' then
                       List1.Items.Add(Copy(S,2,Length(S)-1));
                   end;
                 end;
          end;
        end
      else
      begin
        ForceValue.Checked:=False;
        List1.Items.Clear;
      end;
   end;
    4: begin
        if Sheet4.TypeIndex>=3 then
        begin
          UseFilter.Checked:=False;
          UseFilter.Enabled:=False;
        end
        else
        begin
          UseFilter.Enabled:=True;
          if Sheet4.Filter then
            UseFilter.Checked:=True
          else
            UseFilter.Checked:=False;
          FXLabel.Caption:=StrTemp.Items.Strings[Sheet4.FilterIndex];
        end;
        StrLimitsTable.Enabled:=False;
        case Sheet4.TypeIndex of
          1:;
          2:;
          3: begin
               StrLimitsTable.Enabled:=True;
               StrLimitsTable.Cells[1,1]:=IntToStr(Sheet4.MaxStringLength);
               case Sheet4.CharStyle of
                 csUppercase:StrLimitsTable.Cells[1,2]:='Uppercase';
                 csLowercase:StrLimitsTable.Cells[1,2]:='Lowercase';
                 csNone:StrLimitsTable.Cells[1,2]:='None';
               end;
             end;
          4:;
        end;
        if Sheet4.ForceValue then
        begin
          ForceValue.Checked:=True;
          List1.Items.Clear;
          case Sheet4.TypeIndex of
            1: begin
                 for i:=0 to (IntList4.Items.Count-1) do
                   List1.Items.Add(IntList4.Items.Strings[i]);
               end;
            2: begin
                 for i:=0 to (DecList4.Items.Count-1) do
                   List1.Items.Add(DecList4.Items.Strings[i]);
               end;
            3: begin
                 for i:=0 to (StrList4.Items.Count-1) do
                   List1.Items.Add(StrList4.Items.Strings[i]);
               end;
            4,5: begin
                   for i:=0 to (SpanList.Items.Count-1) do
                   begin
                     S:=SpanList.Items.Strings[i];
                     if S[1]='4' then
                       List1.Items.Add(Copy(S,2,Length(S)-1));
                   end;
                 end;
          end;
        end
      else
      begin
        ForceValue.Checked:=False;
        List1.Items.Clear;
      end;
   end;
    5: begin
        if Sheet5.TypeIndex>=3 then
        begin
          UseFilter.Checked:=False;
          UseFilter.Enabled:=False;
        end
        else
        begin
          UseFilter.Enabled:=True;
          if Sheet5.Filter then
            UseFilter.Checked:=True
          else
            UseFilter.Checked:=False;
          FXLabel.Caption:=StrTemp.Items.Strings[Sheet5.FilterIndex];
        end;
        StrLimitsTable.Enabled:=False;
        case Sheet5.TypeIndex of
          1:;
          2:;
          3: begin
               StrLimitsTable.Enabled:=True;
               StrLimitsTable.Cells[1,1]:=IntToStr(Sheet5.MaxStringLength);
               case Sheet5.CharStyle of
                 csUppercase:StrLimitsTable.Cells[1,2]:='Uppercase';
                 csLowercase:StrLimitsTable.Cells[1,2]:='Lowercase';
                 csNone:StrLimitsTable.Cells[1,2]:='None';
               end;
             end;
          4:;
        end;
        if Sheet5.ForceValue then
        begin
          ForceValue.Checked:=True;
          List1.Items.Clear;
          case Sheet5.TypeIndex of
            1: begin
                 for i:=0 to (IntList5.Items.Count-1) do
                   List1.Items.Add(IntList5.Items.Strings[i]);
               end;
            2: begin
                 for i:=0 to (DecList5.Items.Count-1) do
                   List1.Items.Add(DecList5.Items.Strings[i]);
               end;
            3: begin
                 for i:=0 to (StrList5.Items.Count-1) do
                   List1.Items.Add(StrList5.Items.Strings[i]);
               end;
            4,5: begin
                   for i:=0 to (SpanList.Items.Count-1) do
                   begin
                     S:=SpanList.Items.Strings[i];
                     if S[1]='5' then
                       List1.Items.Add(Copy(S,2,Length(S)-1));
                   end;
                 end;
          end;
        end
      else
      begin
        ForceValue.Checked:=False;
        List1.Items.Clear;
      end;
   end;
  end; {case}
  if CharStyleCombo.Enabled then
    CharStyleCombo.ItemIndex:=CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]);
  CharStyleCombo.Enabled:=StrLimitsTable.Enabled;
end;

procedure TMainForm.UpdateSheet(FieldID:Byte);
var UR:Byte;
    LI,UI:Integer;
begin
  UR:=FieldID;
  case FieldID of
    1: begin
         Sheet1.FieldName:=SettingGrid.Cells[1,UR];
         StrTemp.Items.Strings[Sheet1.DesIndex]:=SettingGrid.Cells[4,UR];
         Sheet1.TypeIndex:=TypeCombo.Items.IndexOf(SettingGrid.Cells[2,UR])+1;
         case Sheet1.TypeIndex of
           1: Sheet1.DValue1:=StrToInt(SettingGrid.Cells[3,UR]);
           2: Sheet1.DValue2:=StrToExt(SettingGrid.Cells[3,UR]);
           3: begin
                Sheet1.DValue3:=SettingGrid.Cells[3,UR];
                Sheet1.MaxStringLength:=StrToInt(StrLimitsTable.Cells[1,1]);
                case CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]) of
                  0: Sheet1.CharStyle:=csUppercase;
                  1: Sheet1.CharStyle:=csLowercase;
                  2: Sheet1.CharStyle:=csNone;
                end;
              end;
           4: begin
                ExtractSpin(LI,Ui,SettingGrid.Cells[3,UR]);
                Sheet1.LBoundI:=LI;
                Sheet1.UBoundI:=UI;
              end;
         end;
         if UseFilter.Enabled=True then
         begin
           Sheet1.Filter:=UseFilter.Checked;
           if UseFilter.Checked  then
             StrTemp.Items.Strings[Sheet1.FilterIndex]:=FXLabel.Caption;
         end
         else
           Sheet1.Filter:=False;
         if not(ForceValue.Checked) then
           EmptyLists(1);
       end;
    2: begin
         Sheet2.FieldName:=SettingGrid.Cells[1,UR];
         StrTemp.Items.Strings[Sheet2.DesIndex]:=SettingGrid.Cells[4,UR];
         Sheet2.TypeIndex:=TypeCombo.Items.IndexOf(SettingGrid.Cells[2,UR])+1;
         case Sheet2.TypeIndex of
           1: Sheet2.DValue1:=StrToInt(SettingGrid.Cells[3,UR]);
           2: Sheet2.DValue2:=StrToExt(SettingGrid.Cells[3,UR]);
           3: begin
                Sheet2.DValue3:=SettingGrid.Cells[3,UR];
                Sheet2.MaxStringLength:=StrToInt(StrLimitsTable.Cells[1,1]);
                case CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]) of
                  0: Sheet2.CharStyle:=csUppercase;
                  1: Sheet2.CharStyle:=csLowercase;
                  2: Sheet2.CharStyle:=csNone;
                end;
              end;
           4: begin
                ExtractSpin(LI,Ui,SettingGrid.Cells[3,UR]);
                Sheet2.LBoundI:=LI;
                Sheet2.UBoundI:=UI;
              end;
{            5: begin
                 ExtractSpin(LE,UE,SettingGrid.Cells[3,UR]);
                 Sheet2.LBoundE:=LE;
                 Sheet2.UBoundE:=UE;
               end;}
         end;
         if UseFilter.Enabled=True then
         begin
           Sheet2.Filter:=UseFilter.Checked;
           if UseFilter.Checked  then
             StrTemp.Items.Strings[Sheet2.FilterIndex]:=FXLabel.Caption;
         end
         else
           Sheet2.Filter:=False;
         if not(ForceValue.Checked) then
           EmptyLists(2);
       end;
    3: begin
         Sheet3.FieldName:=SettingGrid.Cells[1,UR];
         StrTemp.Items.Strings[Sheet3.DesIndex]:=SettingGrid.Cells[4,UR];
         Sheet3.TypeIndex:=TypeCombo.Items.IndexOf(SettingGrid.Cells[2,UR])+1;
         case Sheet3.TypeIndex of
           1: Sheet3.DValue1:=StrToInt(SettingGrid.Cells[3,UR]);
           2: Sheet3.DValue2:=StrToExt(SettingGrid.Cells[3,UR]);
           3: begin
                Sheet3.DValue3:=SettingGrid.Cells[3,UR];
                Sheet3.MaxStringLength:=StrToInt(StrLimitsTable.Cells[1,1]);
                case CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]) of
                  0: Sheet3.CharStyle:=csUppercase;
                  1: Sheet3.CharStyle:=csLowercase;
                  2: Sheet3.CharStyle:=csNone;
                end;
              end;
           4: begin
                ExtractSpin(LI,Ui,SettingGrid.Cells[3,UR]);
                Sheet3.LBoundI:=LI;
                Sheet3.UBoundI:=UI;
              end;
   {         5: begin
                 ExtractSpin(LE,UE,SettingGrid.Cells[3,UR]);
                 Sheet3.LBoundE:=LE;
                 Sheet3.UBoundE:=UE;
               end;}
         end;
         if UseFilter.Enabled=True then
         begin
           Sheet3.Filter:=UseFilter.Checked;
           if UseFilter.Checked  then
             StrTemp.Items.Strings[Sheet3.FilterIndex]:=FXLabel.Caption;
         end
         else
           Sheet3.Filter:=False;
         if not(ForceValue.Checked) then
           EmptyLists(3);
       end;
    4: begin
         Sheet4.FieldName:=SettingGrid.Cells[1,UR];
         StrTemp.Items.Strings[Sheet4.DesIndex]:=SettingGrid.Cells[4,UR];
         Sheet4.TypeIndex:=TypeCombo.Items.IndexOf(SettingGrid.Cells[2,UR])+1;
         case Sheet4.TypeIndex of
           1: Sheet4.DValue1:=StrToInt(SettingGrid.Cells[3,UR]);
           2: Sheet4.DValue2:=StrToExt(SettingGrid.Cells[3,UR]);
           3: begin
                Sheet4.DValue3:=SettingGrid.Cells[3,UR];
                Sheet4.MaxStringLength:=StrToInt(StrLimitsTable.Cells[1,1]);
                case CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]) of
                  0: Sheet4.CharStyle:=csUppercase;
                  1: Sheet4.CharStyle:=csLowercase;
                  2: Sheet4.CharStyle:=csNone;
                end;
              end;
           4: begin
                ExtractSpin(LI,Ui,SettingGrid.Cells[3,UR]);
                Sheet4.LBoundI:=LI;
                Sheet4.UBoundI:=UI;
              end;
{            5: begin
                 ExtractSpin(LE,UE,SettingGrid.Cells[3,UR]);
                 Sheet4.LBoundE:=LE;
                 Sheet4.UBoundE:=UE;
               end;}
         end;
         if UseFilter.Enabled=True then
         begin
           Sheet4.Filter:=UseFilter.Checked;
           if UseFilter.Checked  then
             StrTemp.Items.Strings[Sheet4.FilterIndex]:=FXLabel.Caption;
         end
         else
           Sheet4.Filter:=False;
         if not(ForceValue.Checked) then
           EmptyLists(4);
       end;
    5: begin
         Sheet5.FieldName:=SettingGrid.Cells[1,UR];
         StrTemp.Items.Strings[Sheet5.DesIndex]:=SettingGrid.Cells[4,UR];
         Sheet5.TypeIndex:=TypeCombo.Items.IndexOf(SettingGrid.Cells[2,UR])+1;
         case Sheet5.TypeIndex of
           1: Sheet5.DValue1:=StrToInt(SettingGrid.Cells[3,UR]);
           2: Sheet5.DValue2:=StrToExt(SettingGrid.Cells[3,UR]);
           3: begin
                Sheet5.DValue3:=SettingGrid.Cells[3,UR];
                Sheet5.MaxStringLength:=StrToInt(StrLimitsTable.Cells[1,1]);
                case CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]) of
                  0: Sheet5.CharStyle:=csUppercase;
                  1: Sheet5.CharStyle:=csLowercase;
                  2: Sheet5.CharStyle:=csNone;
                end;
              end;
           4: begin
                ExtractSpin(LI,Ui,SettingGrid.Cells[3,UR]);
                Sheet5.LBoundI:=LI;
                Sheet5.UBoundI:=UI;
              end;
{            5: begin
                 ExtractSpin(LE,UE,SettingGrid.Cells[3,UR]);
                 Sheet5.LBoundE:=LE;
                 Sheet5.UBoundE:=UE;
               end;}
         end;
         if UseFilter.Enabled=True then
         begin
           Sheet5.Filter:=UseFilter.Checked;
           if UseFilter.Checked  then
             StrTemp.Items.Strings[Sheet5.FilterIndex]:=FXLabel.Caption;
         end
         else
           Sheet5.Filter:=False;
         if not(ForceValue.Checked) then
           EmptyLists(5);
       end;
  end;
end;

function TMainForm.ValidateChecks:Boolean;
begin
  Result:=False;
  if not(S1Check.Checked) and not(S2Check.Checked) and not(S3Check.Checked) and not(S4Check.Checked) and not(S5Check.Checked) then
  begin
    ShowMessage('You should use at least one data sheet for your statistic project.');
    Result:=True;
  end;
end;

procedure TMainForm.ResetSheet(var Sheet:SheetSettings;LIndex:Integer);
begin
  with Sheet do
  begin
    FieldName:='';
    StrTemp.Items.Strings[DesIndex]:='';
    ForceValue:=False;
    Filter:=False;
    TypeIndex:=1;
    DValue1:=0;
  end;
  StrTemp.Items.Strings[Sheet.FilterIndex]:='F(X)=X';
  EmptyLists(LIndex);
end;

procedure TMainForm.ResetGridField(FieldID:Integer;Row:Integer;Fill:Boolean);
var i:Byte;
begin
  if Fill then
  begin
    with SettingGrid do
    begin
      Cells[0,FieldID]:='      Data Sheet'+IntToStr(FieldID);
      Cells[1,FieldID]:='Field '+IntToStr(FieldID);
      Cells[2,FieldID]:='Integer';
      Cells[3,FieldID]:='0';
      Cells[4,FieldID]:='My field '+IntToStr(FieldID);
    end;
  end
  else
  begin
    with SettingGrid do
    begin
      Cells[0,FieldID]:='      [Not Used]';
      for i:=1 to 4 do
        Cells[i,FieldID]:='';
    end;
  end;
  Checking:=True;
  case FieldID of
    1: S1Check.Checked:=Fill;
    2: S2Check.Checked:=Fill;
    3: S3Check.Checked:=Fill;
    4: S4Check.Checked:=Fill;
    5: S5Check.Checked:=Fill;
  end;
  Checking:=False;
end;

procedure TMainForm.EmptyLists(FieldID:Byte);
var i,count:Integer; Item:String;
    DelList:array[1..300] of Integer;
begin
  case FieldID of
    1: begin
         IntList1.Items.Clear;
         DecList1.Items.Clear;
         StrList1.Items.Clear;
         count:=0;
         for i:=0 to (sPANlIST.Items.Count-1) do
         begin
           Item:=SpanList.Items.Strings[i];
           if Item[1]='1' then
           begin
             Inc(count);
             DelList[count]:=i;
           end;
         end;
         for i:=1 to count do
           SpanList.Items.Delete(DelList[i]);
         Sheet1.ForceValue:=False;
       end;
    2: begin
         IntList2.Items.Clear;
         DecList2.Items.Clear;
         StrList2.Items.Clear;
         count:=0;
         for i:=0 to (sPANlIST.Items.Count-1) do
         begin
           Item:=SpanList.Items.Strings[i];
           if Item[1]='2' then
           begin
             Inc(count);
             DelList[count]:=i;
           end;
         end;
         for i:=1 to count do
           SpanList.Items.Delete(DelList[i]);
         Sheet2.ForceValue:=False;
       end;
    3: begin
         IntList3.Items.Clear;
         DecList3.Items.Clear;
         StrList3.Items.Clear;
         count:=0;
         for i:=0 to (sPANlIST.Items.Count-1) do
         begin
           Item:=SpanList.Items.Strings[i];
           if Item[1]='3' then
           begin
             Inc(count);
             DelList[count]:=i;
           end;
         end;
         for i:=1 to count do
           SpanList.Items.Delete(DelList[i]);
         Sheet3.ForceValue:=False;
       end;
    4: begin
         IntList4.Items.Clear;
         DecList4.Items.Clear;
         StrList4.Items.Clear;
         count:=0;
         for i:=0 to (sPANlIST.Items.Count-1) do
         begin
           Item:=SpanList.Items.Strings[i];
           if Item[1]='4' then
           begin
             Inc(count);
             DelList[count]:=i;
           end;
         end;
         for i:=1 to count do
           SpanList.Items.Delete(DelList[i]);
         Sheet4.ForceValue:=False;
       end;
    5: begin
         IntList5.Items.Clear;
         DecList5.Items.Clear;
         StrList5.Items.Clear;
         count:=0;
         for i:=0 to (sPANlIST.Items.Count-1) do
         begin
           Item:=SpanList.Items.Strings[i];
           if Item[1]='5' then
           begin
             Inc(count);
             DelList[count]:=i;
           end;
         end;
         for i:=1 to count do
           SpanList.Items.Delete(DelList[i]);
         Sheet5.ForceValue:=False;
       end;
  end; {case}
end;

function TMainForm.IsValidInt(S:String):Boolean;
var n,Code:Integer;
begin
  Result:=True;
  if pos('.',S)>0 then
  begin
    Result:=False;
    Exit;
  end;
  val(S,n,Code);
  if Code <>0 then
    Result:=False;
end;

function TMainForm.IsValidDec(S:String):Boolean;
var n:Extended; Code:Integer;
begin
  Result:=True;
  val(S,n,Code);
  if Code<>0 then
    Result:=False;
end;

procedure TMainForm.MoveCombo(SG:TStringGrid;var Combo:TComboBox;ACol,ARow:Integer);
var i:Integer;
begin
  Combo.Left:=SG.Left;
  Combo.Left:=Combo.Left+SG.ColWidths[0];
  if ACol>1 then
    for i:=SG.LeftCol to (ACol-1) do
      Combo.Left:=Combo.Left+SG.ColWidths[i]+SG.GridLineWidth;
  Combo.Top:=SG.Top;
  Combo.Top:=Combo.Top+SG.RowHeights[0];
  for i:=SG.TopRow to (ARow-1) do
    Combo.Top:=Combo.Top+SG.RowHeights[i]+SG.GridLineWidth;
  Combo.Width:=SG.ColWidths[ACol]+1;
  Combo.Height:=SG.RowHeights[ARow]+1;
  Combo.Visible:=True;
  Combo.BringToFront;
end;

procedure TMainForm.AlertTimerTimer(Sender: TObject);
begin
  if NewAlert then
  begin
    AlertIndex:=0;
    Newalert:=False;
  end;
  Inc(AlertIndex);
  if AlertIndex=AlertCount then
  begin
    AlertTimer.Enabled:=False;
    ShowMessage(Alertstring);
    AlertImage.Visible:=False;
  end;
end;

procedure TMainForm.MemberButtonNormalMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if (x>2) and (y>2) and (x<MemberButtonNormal.Width-2) and (y<MemberButtonNormal.Height-2) then
  begin
    MemberButtonPressed.Visible:=True;
    MemberButtonNormal.Visible:=False;
  end;
end;

procedure TMainForm.MemberButtonPressedMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if (x<=2) or (y<=2) or (x>=MemberButtonPressed.Width-2) or (y>=MemberButtonPressed.Height-2) then
  begin
    MemberButtonNormal.Visible:=True;
    MemberButtonPressed.Visible:=False;
  end;
end;

procedure TMainForm.MemberButtonPressedMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var m:TMouse;
begin
  if (Button=mbLeft) and (x>2) and (y>2) and (x<MemberButtonPressed.Width-2) and (y<MemberButtonPressed.Height-2) then
  begin
    m:=TMouse.Create;
    ViewPopup.Popup(m.CursorPos.X,m.CursorPos.Y);
    m.Free;
    MemberbuttonNormal.Visible:=True;
    MemberbuttonPressed.Visible:=False;
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var i:Integer;
    Ch:File of Double;
begin
  OfficeImageList1.GetBitmap(8,FreqExportBtn.Bitmap);
  SetMenuOwnerDraw;
  Changed:=True;
  ChartHistoryCount:=0;
  for i:=1 to 5 do
    LastFilter[i].Str:='';
{  FastCalcToolbar.Parent:=StatusBar;
  FastCalcToolbar.Left:=StatusBar.Width-FastCalcToolbar.Width-105;
  FastCalcToolbar.Top:=3;}
  ReadConfiguration;
  NeedChartRefresh:=True;
  ArrangeCheckBoxes;
  TTIcon.Icon:=TrayIconImage.Picture.Icon;
  Application.HelpFile:='STATMAN.HLP';
  SetChartDefaults;
  AppendMenu(GetSystemMenu(Application.Handle,False),MF_SEPARATOR,0,'');
  AppendMenu(GetSystemMenu(Application.Handle,False),MF_STRING,SC_AboutItem,'About StatMan...');
  HHP:=Application.HintHidePause;
  HP:=Application.HintPause;
  StrTemp.Items.Clear;
  for i:=1 to 5 do
    StrTemp.Items.Append('');
  for i:=1 to 5 do
    StrTemp.Items.Append('F(X)=X');
  SettingGrid.Cells[0,1]:='      Data Sheet 1';
  for i:=2 to 5 do
    SettingGrid.Cells[0,i]:='      [Not Used]';
  SettingGrid.Cells[1,0]:='Field Name';
  SettingGrid.Cells[2,0]:='Type';
  SettingGrid.Cells[3,0]:='Default Value';
  SettingGrid.Cells[4,0]:='Description';
  SettingGrid.Cells[1,1]:='Field 1';
  SettingGrid.Cells[2,1]:='Integer';
  SettingGrid.Cells[3,1]:='0';
  SettingGrid.Cells[4,1]:='My field 1';
  SetSheetDefaults;
  with Sheet1 do
  begin
    Used:=True;
    FieldName:='Field 1';
    TypeIndex:=1;
    DValue1:=0;
    StrTemp.Items.Strings[DesIndex]:='My field 1';
  end;
  ActiveField:=1;
  SGrid1.ColWidths[0]:=22;
  for i:=1 to (SGrid1.RowCount-2) do
  begin
    SGrid1.Cells[0,i]:=IntToStr(i-1);
    if i<=10 then
      SGrid1.Cells[i,0]:=IntToStr(i-1);
  end;
  SGrid2.ColWidths[0]:=22;
  for i:=1 to (SGrid2.RowCount-2) do
  begin
    SGrid2.Cells[0,i]:=IntToStr(i-1);
    if i<=10 then
      SGrid2.Cells[i,0]:=IntToStr(i-1);
  end;
  SGrid3.ColWidths[0]:=22;
  for i:=1 to (SGrid3.RowCount-2) do
  begin
    SGrid3.Cells[0,i]:=IntToStr(i-1);
    if i<=10 then
      SGrid3.Cells[i,0]:=IntToStr(i-1);
  end;
  SGrid4.ColWidths[0]:=22;
  for i:=1 to (SGrid4.RowCount-2) do
  begin
    SGrid4.Cells[0,i]:=IntToStr(i-1);
    if i<=10 then
      SGrid4.Cells[i,0]:=IntToStr(i-1);
  end;
  SGrid5.ColWidths[0]:=22;
  for i:=1 to (SGrid5.RowCount-2) do
  begin
    SGrid5.Cells[0,i]:=IntToStr(i-1);
    if i<=10 then
      SGrid5.Cells[i,0]:=IntToStr(i-1);
  end;
  StrLimitsTable.Cells[0,0]:='Item';
  StrLimitsTable.Cells[1,0]:='Value';
  StrLimitsTable.Cells[0,1]:='Max String Length';
  StrLimitsTable.Cells[0,2]:='Character Style';
  StrLimitsTable.Cells[1,1]:='5';
  StrLimitsTable.Cells[1,2]:='None';
  StrLimitsTable.Col:=1;
  TableGrid.Cells[0,0]:='Limits';
  TableGrid.Cells[1,0]:='fi';
  TableGrid.Cells[2,0]:='fpi';
  TableGrid.Cells[3,0]:='Pi';
  TableGrid.Cells[4,0]:='Fi';
  TableGrid.Cells[5,0]:='Fpi';
  TableGrid.Cells[6,0]:='Pci';
  TableGrid.Cells[0,TableGrid.RowCount-1]:='Total';
  SGrid1.Tag:=10; //Draw for first
  AutoSizeGridRows(SGrid5);
  ExeDir:=Copy(Application.ExeName,1,Length(Application.ExeName)
            -Length(ExtractFileName(Application.ExeName)));
  ResizeForm;
end;

procedure TMainForm.EditFmlClick(Sender: TObject);
begin
  FormulaForm.FmlEdit.Text:=FXLabel.Caption;
  FormulaForm.OKbtn.Enabled:=True;
  FormulaForm.Cancelbtn.Visible:=True;
  if (FormulaForm.ShowModal=mrOK) then
  begin
    FXLabel.Caption:=FormulaForm.FmlEdit.Text;
    FXLabel.Hint:=FXLabel.Caption+'  {Double click to edit}';
  end;
end;

procedure TMainForm.AlertImageClick(Sender: TObject);
begin
  AlertTimer.Enabled:=False;
  HelpLabel.Font.Color:=clRed;
  HelpLabel.Caption:=AlertString;
  alertImage.Visible:=False;
end;

procedure TMainForm.HelpLabelMouseLeave(Sender: TObject);
begin
  if HelpLabel.Font.Color=clRed then
  begin
    HelpLabel.Font.Color:=clBlack;
    HelpLabel.Caption:='';
  end;
end;

procedure TMainForm.PaneMenuBtnMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var P1,P2:TPoint;
begin
  if Button=mbLeft then
  begin
    P1.X:=PaneMenuBtn.Left-6;
    P1.Y:=PaneMenuBtn.Top-15;
    P2:=PaneMenuBtn.ClientToScreen(P1);
    MemberMenu.Popup(P2.X,P2.Y);
  end;
end;

procedure TMainForm.MemberMenuPopup(Sender: TObject);
begin
  MSheet1Section.Visible:=S1Check.Checked;
  MSheet2Section.Visible:=S2Check.Checked;
  MSheet3Section.Visible:=S3Check.Checked;
  MSheet4Section.Visible:=S4Check.Checked;
  MSheet5Section.Visible:=S5Check.Checked;
  if (SettingSection.Down or ToolSettings.Down or IconSettings.Down) then MSettingSection.Checked:=True
  else  if (Sheet1Section.Down or ToolSheet1.Down or IconSheet1.Down) then MSheet1Section.Checked:=True
  else  if (Sheet2Section.Down or ToolSheet2.Down or IconSheet2.Down) then MSheet2Section.Checked:=True
  else  if (Sheet3Section.Down or ToolSheet3.Down or IconSheet3.Down) then MSheet3Section.Checked:=True
  else  if (Sheet4Section.Down or ToolSheet4.Down or IconSheet4.Down) then MSheet4Section.Checked:=True
  else  if (Sheet5Section.Down or ToolSheet5.Down or IconSheet5.Down) then MSheet5Section.Checked:=True
  else if (TableSection.Down or ToolTable.Down or IconTable.Down) then MTableSection.Checked:=True
  else if (ChartSection.Down or ToolChart.Down or IconChart.Down) then MChartSection.Checked:=True
  else MAnalyzeSection.Checked:=True;
end;

procedure TMainForm.SettingGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var S:String; i:Integer;
begin
  if ACol=0 then Exit;
  S:='';
//  ShowMessage('Selected a cell');
  if (S1Check.Checked) and (Length(SettingGrid.Cells[3,1])=0) then
    SettingGrid.Cells[3,1]:=GetDefaultValue(1);
  if (S2Check.Checked) and (Length(SettingGrid.Cells[3,2])=0) then
    SettingGrid.Cells[3,2]:=GetDefaultValue(2);
  if (S3Check.Checked) and (Length(SettingGrid.Cells[3,3])=0) then
    SettingGrid.Cells[3,3]:=GetDefaultValue(3);
  if (S4Check.Checked) and (Length(SettingGrid.Cells[3,4])=0) then
    SettingGrid.Cells[3,4]:=GetDefaultValue(4);
  if (S5Check.Checked) and (Length(SettingGrid.Cells[3,5])=0) then
    SettingGrid.Cells[3,5]:=GetDefaultValue(5);

  for i:=1 to (SettingGrid.RowCount-1) do
  begin
    if (Length(SettingGrid.Cells[2,i])>0) and (Length(SettingGrid.Cells[1,i])=0) then
    begin
      S:='You should define a name for the field.';
      SettingGrid.Cells[1,i]:='Field Name';
    end;
    if (SettingGrid.Cells[2,i]='Integer') then
    begin
      if not(IsValidInt(SettingGrid.Cells[3,i])) then
      begin
        S:='"'+SettingGrid.Cells[3,i]+'" is not a valid integer value.';
        SettingGrid.Cells[3,i]:=GetDefaultValue(i);
      end
      else
        SettingGrid.Cells[3,i]:=IntToStr(StrToInt(SettingGrid.Cells[3,i]));
    end;
    if (SettingGrid.Cells[2,i]='Span') and not(IsValidSpan(SettingGrid.Cells[3,i])) then
    begin
      S:='"'+SettingGrid.Cells[3,i]+'" is not a valid span.';
      SettingGrid.Cells[3,i]:=GetDefaultValue(i);
    end;
    if SettingGrid.Cells[2,i]='Decimal' then
    begin
      if not(IsValidDec(SettingGrid.Cells[3,i])) then
      begin
        S:='"'+SettingGrid.Cells[3,i]+'" is not a valid decimal value.';
        SettingGrid.Cells[3,i]:=GetDefaultValue(i);
      end
      else
      begin
        S:=FloatToStr(StrToFloat(SettingGrid.Cells[3,i]));
        if S='0' then S:='0.0';
        if not(Pos('.',S)>0) then
          S:=S+'.0';
        SettingGrid.Cells[3,i]:=S;
        S:='';
      end;
    end;
  end;
  if S<>'' then
  begin
    ShowMessage(S);
    CanSelect:=False;
    Exit;
  end;
  CanSelect:=True;
  case ARow of
    1: if not(S1Check.Checked) then CanSelect:=False;
    2: if not(S2Check.Checked) then CanSelect:=False;
    3: if not(S3Check.Checked) then CanSelect:=False;
    4: if not(S4Check.Checked) then CanSelect:=False;
    5: if not(S5Check.Checked) then CanSelect:=False;
  end;
  if not(CanSelect) then
    Exit;
  if Acol=2 then
  begin
    SettingGrid.Options:=SettingGrid.Options-[goEditing];
    TypeCombo.Visible:=False;
    TypeCombo.ItemIndex:=TypeCombo.Items.IndexOf(SettingGrid.Cells[2,Arow]);
    MoveCombo(SettingGrid,TypeCombo,ACol,ARow);
    TypeCombo.Visible:=True;
  end
  else
  begin
    SettingGrid.Options:=SettingGrid.Options+[goEditing];
    TypeCombo.Visible:=False;
  end;
//  S:=SettingGrid.Cells[3,ARow];
//  if S='Str1' then ShowMessage('Error');
  if (ActiveField=1) and S1Check.Checked and (Length(SettingGrid.Cells[1,1])>0) then UpdateSheet(1);
  if (ActiveField=2) and S2Check.Checked and (Length(SettingGrid.Cells[1,2])>0) then UpdateSheet(2);
  if (ActiveField=3) and S3Check.Checked and (Length(SettingGrid.Cells[1,3])>0) then UpdateSheet(3);
  if (ActiveField=4) and S4Check.Checked and (Length(SettingGrid.Cells[1,4])>0) then UpdateSheet(4);
  if (ActiveField=5) and S5Check.Checked and (Length(SettingGrid.Cells[1,5])>0) then UpdateSheet(5);
//  if ARow=ActiveField then
//    Exit;
//  UpdateSheet(ActiveField);
  ActiveField:=ARow;
  case ActiveField of
    1: CurrentType:=Sheet1.TypeIndex;
    2: CurrentType:=Sheet2.TypeIndex;
    3: CurrentType:=Sheet3.TypeIndex;
    4: CurrentType:=Sheet4.TypeIndex;
    5: CurrentType:=Sheet5.TypeIndex;
  end;
//  ShowMessage('Before loadsheet CurrentType = '+IntToStr(CurrentType));
  LoadSheet(ActiveField);
//  ShowMessage('CurrentType = '+IntToStr(CurrentType));
end;

procedure TMainForm.ForceValueClick(Sender: TObject);
begin
  VEdit1.Enabled:=ForceValue.Checked;
  if VEdit1.Text<>'' then
    Add1.Enabled:=ForceValue.Checked
  else
    Add1.Enabled:=False;
  Add1.Repaint;  
  if List1.ItemIndex >=0 then
    Del1.Enabled:=ForceValue.Checked
  else
    Del1.Enabled:=False;
  Del1.Repaint;  
  List1.Enabled:=ForceValue.Checked;
  case ActiveField of
    1: Sheet1.ForceValue:=ForceValue.Checked;
    2: Sheet2.ForceValue:=ForceValue.Checked;
    3: Sheet3.ForceValue:=ForceValue.Checked;
    4: Sheet4.ForceValue:=ForceValue.Checked;
    5: Sheet5.ForceValue:=ForceValue.Checked;
  end;
  VEdit1.Text:='';
end;

procedure TMainForm.S1CheckClick(Sender: TObject);
begin
  RebuildFreqTable:=True;
  if ValidateChecks=True then
  begin
    Checking:=True;
    S1Check.OnClick:=Nil;
    S1Check.Checked:=True;
    S1Check.OnClick:=S1CheckClick;
    Checking:=False;
    Exit;
  end;
  if not(Checking) then
  begin
    if S1Check.Checked then
    begin
      ResetGridField(1,1,True);
      ResetSheet(Sheet1,1);
      CurrentType:=1;
      Sheet1.TypeIndex:=1;
      SettingGrid.Row:=1;
      SettingGrid.Col:=1;
    end
    else
    begin
      DeSelectRow;
      ResetGridField(1,1,False);
    end;
    Sheet1.Used:=S1Check.Checked;
  end;
  Sheet1Section.Enabled:=S1Check.Checked;
  ToolSheet1.Enabled:=S1Check.Checked;
  IconSheet1.Enabled:=S1Check.Checked;
end;

procedure TMainForm.VEdit1Change(Sender: TObject);
var b:Boolean;
begin
  if (Length(VEdit1.Text)=0) or not(ForceValue.Checked) then
    b:=False
  else
    b:=True;
  Add1.Enabled:=b;
  Add1.Repaint;
end;

procedure TMainForm.List1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  if (List1.ItemIndex>=0) and not(Del1.Enabled) then
  begin
    Del1.Enabled:=True;
    Del1.Repaint;
  end
  else if (List1.ItemIndex<0) and Del1.Enabled then
  begin
    Del1.Enabled:=False;
    Del1.Repaint;
  end;
  HelpLabel.Caption:='Value List:'+Chr(13)+Chr(13)+'Displays the list of the force values of the selected data sheet';
end;

procedure TMainForm.SettingGridSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: String);
var Alert:Boolean; S:String;
begin
  if not(Changed) then
    Changed:=True;
  S:=Value;
  Alert:=False;
  if (ACol=1) and (Length(Value)>MaxFieldLen) then
  begin
    Alert:=True;
    AlertString:='The maximum length of field name is '+IntToStr(MaxFieldLen)+' characters.';
    SettingGrid.Cells[ACol,ARow]:=Copy(Value,1,MaxFieldLen);
  end;
  if (ACol=5) and (Length(Value)>MaxDesLen) then
  begin
    Alert:=True;
    AlertString:='The maximum length of description for a field is '+IntToStr(MaxDesLen)+' characters.';
    SettingGrid.Cells[ACol,ARow]:=Copy(Value,1,MaxDesLen);
  end;
  if (Length(Value)>MaxStringLen) and (SettingGrid.Cells[2,ARow]='String') then
  begin
    Alert:=True;
    AlertString:='The maximum length of string type is '+IntToStr(MaxStringLen)+' characters.';
    SettingGrid.Cells[ACol,ARow]:=Copy(Value,1,MaxStringLen);
  end;
  if Alert and (AlertTimer.Enabled=False) then
  begin
    AlertCount:=8;
    NewAlert:=True;
    AlertImage.Visible:=True;
    AlertTimer.Enabled:=True;
  end;
end;

procedure TMainForm.S2CheckClick(Sender: TObject);
begin
  RebuildFreqTable:=True;
  if ValidateChecks=True then
  begin
    Checking:=True;
    S2Check.OnClick:=Nil;
    S2Check.Checked:=True;
    S2Check.OnClick:=S2CheckClick;
    Checking:=False;
    Exit;
  end;
  if not(Checking) then
  begin
    if S2Check.Checked then
    begin
      ResetGridField(2,2,True);
      ResetSheet(Sheet2,2);
      CurrentType:=1;
      Sheet2.TypeIndex:=1;
      SettingGrid.Row:=2;
      SettingGrid.Col:=1;
    end
    else
    begin
      DeSelectRow;
      ResetGridField(2,2,False);
    end;
    Sheet2.Used:=S2Check.Checked;
  end;
  Sheet2Section.Enabled:=S2Check.Checked;
  ToolSheet2.Enabled:=S2Check.Checked;
  IconSheet2.Enabled:=S2Check.Checked;
end;

procedure TMainForm.S3CheckClick(Sender: TObject);
begin
  RebuildFreqTable:=True;
  if ValidateChecks=True then
  begin
    Checking:=True;
    S3Check.OnClick:=Nil;
    S3Check.Checked:=True;
    S3Check.OnClick:=S3CheckClick;
    Checking:=False;
    Exit;
  end;
  if not(Checking) then
  begin
    if S3Check.Checked then
    begin
      ResetGridField(3,3,True);
      ResetSheet(Sheet3,3);
      CurrentType:=1;
      Sheet3.TypeIndex:=1;
      SettingGrid.Row:=3;
      SettingGrid.Col:=1;
    end
    else
    begin
      DeSelectRow;
      ResetGridField(3,3,False);
    end;
    Sheet3.Used:=S3Check.Checked;
  end;
  Sheet3Section.Enabled:=S3Check.Checked;
  ToolSheet3.Enabled:=S3Check.Checked;
  IconSheet3.Enabled:=S3Check.Checked;
end;

procedure TMainForm.S4CheckClick(Sender: TObject);
begin
  RebuildFreqTable:=True;
  if ValidateChecks=True then
  begin
    Checking:=True;
    S4Check.OnClick:=Nil;
    S4Check.Checked:=True;
    S4Check.OnClick:=S4CheckClick;
    Checking:=False;
    Exit;
  end;
  if not(Checking) then
  begin
    if S4Check.Checked then
    begin
      ResetGridField(4,4,True);
      ResetSheet(Sheet4,4);
      CurrentType:=1;
      Sheet4.TypeIndex:=1;
      SettingGrid.Row:=4;
      SettingGrid.Col:=1;
    end
    else
    begin
      DeSelectRow;
      ResetGridField(4,4,False);
    end;
    Sheet4.Used:=S4Check.Checked;
  end;
  Sheet4Section.Enabled:=S4Check.Checked;
  ToolSheet4.Enabled:=S4Check.Checked;
  IconSheet4.Enabled:=S4Check.Checked;
end;

procedure TMainForm.S5CheckClick(Sender: TObject);
begin
  RebuildFreqTable:=True;
  if ValidateChecks=True then
  begin
    Checking:=True;
    S5Check.OnClick:=Nil;
    S5Check.Checked:=True;
    S5Check.OnClick:=S5CheckClick;
    Checking:=False;
    Exit;
  end;
  if not(Checking) then
  begin
    if S5Check.Checked then
    begin
      ResetGridField(5,5,True);
      ResetSheet(Sheet5,5);
      CurrentType:=1;
      Sheet5.TypeIndex:=1;
      SettingGrid.Row:=5;
      SettingGrid.Col:=1;
    end
    else
    begin
      DeSelectRow;
      ResetGridField(5,5,False);
    end;
    Sheet5.Used:=S5Check.Checked;
  end;
  Sheet5Section.Enabled:=S5Check.Checked;
  ToolSheet5.Enabled:=S5Check.Checked;
  IconSheet5.Enabled:=S5Check.Checked;
end;

procedure TMainForm.HelpLabelMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (AlertTimer.Enabled) and (x>=117) and (y>=133) then
  begin
    AlertTimer.Enabled:=False;
    HelpLabel.Font.Color:=clRed;
    HelpLabel.Caption:=AlertString;
    AlertImage.Visible:=True;
  end;

end;

procedure TMainForm.UseFilterClick(Sender: TObject);
begin
  if UseFilter.Checked then
  begin
    case ActiveField of
      1: begin
           Sheet1.Filter:=UseFilter.Checked;
           FXLabel.Caption:=StrTemp.Items.Strings[sheet1.FilterIndex];
         end;
      2: begin
           Sheet2.Filter:=UseFilter.Checked;
           FXLabel.Caption:=StrTemp.Items.Strings[sheet2.FilterIndex];
         end;
      3: begin
           Sheet3.Filter:=UseFilter.Checked;
           FXLabel.Caption:=StrTemp.Items.Strings[sheet3.FilterIndex];
         end;
      4: begin
           Sheet4.Filter:=UseFilter.Checked;
           FXLabel.Caption:=StrTemp.Items.Strings[sheet4.FilterIndex];
         end;
      5: begin
           Sheet5.Filter:=UseFilter.Checked;
           FXLabel.Caption:=StrTemp.Items.Strings[sheet5.FilterIndex];
         end;
    end;
    FXLabel.Font.Color:=clBlack;
    FXLabel.ShowHint:=UseFilter.Checked;
    FXLabel.Hint:=FXLabel.Caption+'  {Double click to edit}';
  end
  else
  begin
    FXLabel.ShowHint:=False;
    FXLabel.Font.Color:=clGray;
  end;
  EditFml.Enabled:=UseFilter.Checked;
  EditFml.Repaint;
end;

procedure TMainForm.FxLabelDblClick(Sender: TObject);
begin
  if UseFilter.Checked then
    EditFml.OnClick(EditFml);
end;

procedure TMainForm.Add1Click(Sender: TObject);
var Alert:Boolean;
    Code:Byte;
begin
  Alert:=False;
  case CurrentType of
    1:begin
        if not(IsValidInt(VEdit1.Text)) then
        begin
          Alert:=True;
          AlertString:='You shoud enter an integer for the values of this field.';
        end
        else
        begin
          VEdit1.Text:=IntToStr(StrToInt(VEdit1.Text));
          List1.Items.Add(VEdit1.Text);
          case Activefield of
            1:IntList1.Items.Add(VEdit1.Text);
            2:IntList2.Items.Add(VEdit1.Text);
            3:IntList3.Items.Add(VEdit1.Text);
            4:IntList4.Items.Add(VEdit1.Text);
            5:IntList5.Items.Add(VEdit1.Text);
          end;
          VEdit1.Text:='';
        end;
      end;
    2:begin
        if not(IsValidDec(VEdit1.Text)) or (VEdit1.Text='.') then
        begin
          Alert:=True;
          Alertstring:='You shoud enter a decimal number for the values of this field.';
        end
        else
        begin
          VEdit1.Text:=FloatToStr(StrToFloat(VEdit1.Text));
          if not(Pos('.',VEdit1.Text)>0) then
            VEdit1.Text:=VEdit1.Text+'.0';
          List1.Items.Add(VEdit1.Text);
          case Activefield of
            1:DecList1.Items.Add(VEdit1.Text);
            2:DecList2.Items.Add(VEdit1.Text);
            3:DecList3.Items.Add(VEdit1.Text);
            4:DecList4.Items.Add(VEdit1.Text);
            5:DecList5.Items.Add(VEdit1.Text);
          end;
          VEdit1.Text:='';
        end;
      end;
    3:begin
        List1.Items.Add(VEdit1.Text);
        case Activefield of
          1:StrList1.Items.Add(VEdit1.Text);
          2:StrList2.Items.Add(VEdit1.Text);
          3:StrList3.Items.Add(VEdit1.Text);
          4:StrList4.Items.Add(VEdit1.Text);
          5:StrList5.Items.Add(VEdit1.Text);
        end;
        VEdit1.Text:='';
      end;
    4:begin
        if IsValidSpan(VEdit1.Text) then
        begin
          List1.Items.Add(VEdit1.Text);
          case Activefield of
            1:Code:=Sheet1.VListIndex;
            2:Code:=Sheet2.VListIndex;
            3:Code:=Sheet3.VListIndex;
            4:Code:=Sheet4.VListIndex;
            5:Code:=Sheet5.VListIndex;
          end;
          if StrToInt(Copy(VEdit1.Text,2,pos(',',VEdit1.Text)-2))>=StrToInt(Copy(VEdit1.Text,pos(',',VEdit1.Text)+1,Length(VEdit1.Text)-pos(',',VEdit1.Text)-1)) then
          begin
            ShowMessage('The span lower bound value should be lower than the span upper bound value.');
            VEdit1.SelectAll;
            VEdit1.SetFocus;
            Exit;
          end;
          SpanList.Items.Add(IntToStr(Code)+VEdit1.Text);
          VEdit1.Text:='';
        end
        else
        begin
          Alert:=True;
          AlertString:='The value you typed is not a valid span.';
        end;
      end;
    end;
  if Alert and (AlertTimer.Enabled=False) then
  begin
    AlertCount:=8;
    NewAlert:=True;
    AlertImage.Visible:=True;
    AlertTimer.Enabled:=True;
  end;
  VEdit1.SelectAll;
  VEdit1.SetFocus;
end;

procedure TMainForm.Del1Click(Sender: TObject);
var Code:Byte;
begin
  case CurrentType of
    1:begin
        case ActiveField of
          1:IntList1.Items.Delete(IntList1.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          2:IntList2.Items.Delete(IntList2.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          3:IntList3.Items.Delete(IntList3.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          4:IntList4.Items.Delete(IntList4.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          5:IntList5.Items.Delete(IntList5.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
        end;
        VEdit1.Text:=List1.Items.Strings[List1.ItemIndex];
        List1.Items.Delete(List1.ItemIndex);
      end;
    2:begin
        case ActiveField of
          1:DecList1.Items.Delete(DecList1.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          2:DecList2.Items.Delete(DecList2.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          3:DecList3.Items.Delete(DecList3.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          4:DecList4.Items.Delete(DecList4.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          5:DecList5.Items.Delete(DecList5.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
        end;
        VEdit1.Text:=List1.Items.Strings[List1.ItemIndex];
        List1.Items.Delete(List1.ItemIndex);
      end;
    3:begin
        case ActiveField of
          1:StrList1.Items.Delete(StrList1.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          2:StrList2.Items.Delete(StrList2.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          3:StrList3.Items.Delete(StrList3.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          4:StrList4.Items.Delete(StrList4.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
          5:StrList5.Items.Delete(StrList5.Items.IndexOf(List1.Items.Strings[List1.ItemIndex]));
        end;
        VEdit1.Text:=List1.Items.Strings[List1.ItemIndex];
        List1.Items.Delete(List1.ItemIndex);
      end;
    4,5:begin
          case ActiveField of
            1:Code:=Sheet1.VListIndex;
            2:Code:=Sheet2.VListIndex;
            3:Code:=Sheet3.VListIndex;
            4:Code:=Sheet4.VListIndex;
            5:Code:=Sheet5.VListIndex;
          end;
          SpanList.Items.Delete(SpanList.Items.IndexOf(IntToStr(Code)+List1.Items.Strings[List1.ItemIndex]));
          VEdit1.Text:=List1.Items.Strings[List1.ItemIndex];
          List1.Items.Delete(List1.ItemIndex);
        end;
  end;
  if List1.ItemIndex<0 then
  begin
    Del1.Enabled:=False;
    Del1.Repaint;
  end;  
end;

procedure TMainForm.VEdit1KeyPress(Sender: TObject; var Key: Char);
begin
  if key=#13 then
    if Add1.Enabled then
    begin
      Key:=#0;
      Add1.OnClick(Add1);
    end;
end;

procedure TMainForm.TypeComboChange(Sender: TObject);
begin
  {OnChange}
  if not(Changed) then
    Changed:=True;
  RebuildFreqTable:=True;
  case TypeCombo.ItemIndex of
    0:begin
        SettingGrid.Cells[SettingGrid.Col,SettingGrid.Row]:='Integer';
        if CurrentType>=2 then
        begin
          SettingGrid.Cells[3,SettingGrid.Row]:='0';
          if ForceValue.Checked then
          begin
            EmptyLists(ActiveField);
            ForceValue.Checked:=False;
            ForceValue.Checked:=True;
          end;
        end;
      end;
    1:begin
        SettingGrid.Cells[SettingGrid.Col,SettingGrid.Row]:='Decimal';
        if (CurrentType>=3) or (CurrentType=1) then
        begin
          SettingGrid.Cells[3,SettingGrid.Row]:='0.0';
          if ForceValue.Checked then
          begin
            EmptyLists(ActiveField);
            ForceValue.Checked:=False;
            ForceValue.Checked:=True;
          end;
        end;
      end;
     2:begin
         SettingGrid.Cells[SettingGrid.Col,SettingGrid.Row]:='String';
         if (CurrentType>=4) or (CurrentType<=2) then
         begin
           SettingGrid.Cells[3,SettingGrid.Row]:='Str1';
           if ForceValue.Checked then
           begin
             EmptyLists(ActiveField);
             ForceValue.Checked:=False;
             ForceValue.Checked:=True;
           end;
         end;
       end;
     3:begin
         SettingGrid.Cells[SettingGrid.Col,SettingGrid.Row]:='Span';
         SettingGrid.Cells[3,SettingGrid.Row]:='[0,1]';
         if ForceValue.Checked then
         begin
           EmptyLists(ActiveField);
           ForceValue.Checked:=False;
           ForceValue.Checked:=True;
         end;
       end;
   end; {case}
//   S:=SettingGrid.Cells[2,SettingGrid.Row
   CurrentType:=TypeCombo.ItemIndex+1;
   case SettingGrid.Row of
     1: begin
         Sheet1.TypeIndex:=CurrentType;
         Sheet1.MaxStringLength:=5;
         Sheet1.CharStyle:=csNone;
       end;
     2: begin
         Sheet2.TypeIndex:=CurrentType;
         Sheet2.MaxStringLength:=5;
         Sheet2.CharStyle:=csNone;
       end;
     3: begin
         Sheet3.TypeIndex:=CurrentType;
         Sheet3.MaxStringLength:=5;
         Sheet3.CharStyle:=csNone;
       end;
     4: begin
         Sheet4.TypeIndex:=CurrentType;
         Sheet4.MaxStringLength:=5;
         Sheet4.CharStyle:=csNone;
       end;
     5: begin
         Sheet5.TypeIndex:=CurrentType;
         Sheet5.MaxStringLength:=5;
         Sheet5.CharStyle:=csNone;
       end;
   end;
//   SettingGrid.Col:=2;
//   SettingGrid.Row:=SettingGrid.Row;
   LoadSheet(ActiveField);
//   ShowMessage('CurrentType = '+IntToStr(CurrentType));
end;

procedure TMainForm.SGrid1SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var S:String;
begin
  if SGrid1.Tag=5 then
  begin
    CanSelect:=True;
    Exit;
  end;
  GridSelectCell(Sheet1,1,SGrid1,S1Combo,S1NLabel,ACol,ARow);
  if CheckGridCell(SGrid1,S1TypeLabel)=False then
  begin
    CanSelect:=False;
    Exit;
  end;
  case Sheet1.TypeIndex of
    1:S:=IntToStr(Sheet1.DValue1);
    2:begin
        S:=FloatToStr(Sheet1.DValue2);
        if (Pos('E',S)>0) or (Pos('e',S)>0) then S:='0.00';
      end;
    3:S:=Sheet1.DValue3;
    4:S:='['+IntToStr(Sheet1.LBoundI)+','+IntToStr(Sheet1.UBoundI)+']';
  end;
  if S1Combo.Visible then S1Combo.Visible:=False;
  if {not(goEditing in SGrid1.Options)}Sheet1.ForceValue and CanEditCell(ACol,ARow,ActiveSheet) then
  begin
    if not(DVCheck.Checked) then
    begin
      if S1Combo.Items.IndexOf(S)>0 then
        S1Combo.Items.Delete(S1Combo.Items.IndexOf(S));
      if S1Combo.Items.Count>0 then
        S1Combo.ItemIndex:=0;
    end;
    if DVCheck.Checked then
    begin
      if not(S1Combo.Items.IndexOf(S)>=0) then
        S1Combo.Items.Add(S);
      if (Length(SGrid1.Cells[ACol,ARow])>0) and (S1Combo.Items.IndexOf(SGrid1.Cells[ACol,ARow])>=0) then
        S1Combo.ItemIndex:=S1Combo.Items.IndexOf(SGrid1.Cells[ACol,ARow])
      else
        S1Combo.ItemIndex:=S1Combo.Items.IndexOf(S);
    end;
    S1Combo.Visible:=False;
    MoveCombo(SGrid1,S1Combo,ACol,ARow);
    S1Combo.Visible:=True;
    S1Combo.SetFocus;
    S1Combo.SelectAll;
  end;
  if DVCheck.Checked and (Length(SGrid1.Cells[ACol,ARow])=0) then
    SGrid1.Cells[ACol,ARow]:=S;
  if Length(SGrid1.Cells[ACol,ARow])>0 then
    S1Combo.Hint:='Current value: '+SGrid1.Cells[ACol,ARow]
  else
    S1Combo.Hint:='No value is set';
end;

procedure TMainForm.S1ComboKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE	 then
  begin
    SGrid1.Cells[SGrid1.Col,SGrid1.Row]:='';
    S1Combo.Visible:=False;
  end;
end;

procedure TMainForm.S1ComboChange(Sender: TObject);
begin
  if CanEditCell(SGrid1.Col,SGrid1.Row,1) then
    with SGrid1 do
      Cells[Col,Row]:=S1Combo.Items.Strings[S1Combo.ItemIndex];
  if Length(SGrid1.Cells[SGrid1.Col,SGrid1.Row])>0 then
    S1Combo.Hint:='Current value: '+SGrid1.Cells[SGrid1.Col,SGrid1.Row]
  else
    S1Combo.Hint:='No value is set';
end;

procedure TMainForm.Sheet1SectionClick(Sender: TObject);
var i:Integer; S:String;
    B:Boolean;
begin
  PrintToolbtn.Enabled:=True;
  CutToolbtn.Enabled:=True;
  CopyToolbtn.Enabled:=True;
  PasteToolbtn.Enabled:=True;
  FindToolbtn.Enabled:=True;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  ActiveSheet:=1;
  if SettingGrid.Col=4 then
    SettingGrid.Col:=1
  else
    SettingGrid.Col:=4;
  PrepareSheetGrid(Sheet1,SGrid1,1,1,S1FieldLabel,S1TypeLabel,S1DesLabel,S1FilterLabel,S1NLabel,S1Combo,S1Filterbtn,S1PSpin);
  if Sheet1.ForceValue then
  begin
    S1Combo.Items.Clear;
    case Sheet1.TypeIndex of
      1:S1Combo.Items:=IntList1.Items;
      2:S1Combo.Items:=DecList1.Items;
      3:S1Combo.Items:=StrList1.Items;
      4:for i:=0 to (SpanList.Items.Count-1) do
        begin
          S:=SpanList.Items.Strings[i];
          if S[1]=IntToStr(Sheet1.VListIndex) then
            S1Combo.Items.Append(Copy(S,2,Length(s)-1));
        end;
    end;
    SGrid1.Col:=1;
    SGrid1.Row:=1;
    MoveCombo(SGrid1,S1Combo,1,1);
    S1Combo.Visible:=True;
  end
  else
    S1Combo.Visible:=False;
  S:=SGrid1.Cells[2,1];
  SGrid1.Col:=2;
  SGrid1.Col:=1;
  SGrid1.Cells[2,1]:=S;
  Sheet1Panel.Visible:=True;
  Sheet1Panel.BringToFront;
  ToolSheet1.Down:=True;
  IconSheet1.Down:=True;
  MSheet1Section.Checked:=True;
  ForceSheetLimitations(Sheet1,SGrid1);
  GridFindForm.Grid:=SGrid1;
  GridReplaceForm.Grid:=SGrid1;
end;

procedure TMainForm.SGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i,j:Integer; R:TRect; S:String;
begin
  for i:=0 to (SGrid1.ColCount-1) do
  begin
    R:=SGrid1.CellRect(i,0);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S1IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=0 to (SGrid1.RowCount-2) do
  begin
    R:=Sgrid1.CellRect(0,i);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S1IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=SGrid1.TopRow to (SGrid1.RowCount-2) do
    for j:=1 to (SGrid1.ColCount-1) do
    begin
      R:=SGrid1.CellRect(j,i);
      if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
      begin
        S:=Sgrid1.Cells[0,i]+SGrid1.Cells[j,0];
        if S[1]='0' then
          S:=Copy(S,2,Length(S)-1);
        S1IDEdit.Caption:=S;
        Exit;
      end;
    end;
  S1IDEdit.Caption:='000';
end;

procedure TMainForm.S1ClrbtnClick(Sender: TObject);
begin
  if Application.MessageBox('Are you sure you want to clear all values in data sheet 1?','Delete',MB_YesNo)=IDYes then
  begin
    SGrid1.Col:=1; SGrid1.Row:=1;
    ClearGrid(SGrid1);
    ValueCount[1]:=0;
    S1NLabel.Caption:='0';
  end;
end;

procedure TMainForm.SGrid1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_Delete) and not(goEditing in SGrid1.Options) then
    SGrid1.Cells[SGrid1.Col,SGrid1.Row]:='';
end;

procedure TMainForm.SettingSectionClick(Sender: TObject);
begin
  SettingPanel.Visible:=True;
  SettingPanel.BringToFront;
  ActiveSheet:=0;
  ToolSettings.Down:=True;
  IconSettings.Down:=True;
  PrintToolbtn.Enabled:=False;
  CutToolbtn.Enabled:=False;
  CopyToolbtn.Enabled:=False;
  PasteToolbtn.Enabled:=False;
  FindToolbtn.Enabled:=False;
end;

procedure TMainForm.Sheet2SectionClick(Sender: TObject);
var i:Integer; S:String; B:Boolean;
begin
  PrintToolbtn.Enabled:=True;
  CutToolbtn.Enabled:=True;
  CopyToolbtn.Enabled:=True;
  PasteToolbtn.Enabled:=True;
  FindToolbtn.Enabled:=True;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  ActiveSheet:=2;
  if SettingGrid.Col=4 then
    SettingGrid.Col:=1
  else
    SettingGrid.Col:=4;
//  Updatesheet(2);
  PrepareSheetGrid(Sheet2,SGrid2,2,2,S2FieldLabel,S2TypeLabel,S2DesLabel,S2FilterLabel,S2NLabel,S2Combo,S2Filterbtn,S2PSpin);
  if Sheet2.ForceValue then
  begin
    S2Combo.Items.Clear;
    case Sheet2.TypeIndex of
      1:S2Combo.Items:=IntList2.Items;
      2:S2Combo.Items:=DecList2.Items;
      3:S2Combo.Items:=StrList2.Items;
      4:for i:=0 to (SpanList.Items.Count-1) do
        begin
          S:=SpanList.Items.Strings[i];
          if S[1]=IntToStr(Sheet2.VListIndex) then
            S2Combo.Items.Append(Copy(S,2,Length(s)-1));
        end;
    end;
    SGrid2.Col:=1;
    SGrid2.Row:=1;
    MoveCombo(SGrid2,S2Combo,1,1);
    S2Combo.Visible:=True;
  end
  else
    S2Combo.Visible:=False;
  S:=SGrid2.Cells[2,1];
  SGrid2.Col:=2;
  SGrid2.Col:=1;
  SGrid2.Cells[2,1]:=S;
  Sheet2Panel.Visible:=True;
  Sheet2Panel.BringToFront;
  ToolSheet2.Down:=True;
  IconSheet2.Down:=True;
  MSheet2Section.Checked:=True;
  ForceSheetLimitations(Sheet2,SGrid2);
  GridFindForm.Grid:=SGrid2;
  GridReplaceForm.Grid:=SGrid2;
end;

procedure TMainForm.Sheet3SectionClick(Sender: TObject);
var i:Integer; S:String; B:Boolean;
begin
  PrintToolbtn.Enabled:=True;
  CutToolbtn.Enabled:=True;
  CopyToolbtn.Enabled:=True;
  PasteToolbtn.Enabled:=True;
  FindToolbtn.Enabled:=True;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  ActiveSheet:=3;
  if SettingGrid.Col=4 then
    SettingGrid.Col:=1
  else
    SettingGrid.Col:=4;
//  Updatesheet(3);
  PrepareSheetGrid(Sheet3,SGrid3,3,3,S3FieldLabel,S3TypeLabel,S3DesLabel,S3FilterLabel,S3NLabel,S3Combo,S3Filterbtn,S3PSpin);
  if Sheet3.ForceValue then
  begin
    S3Combo.Items.Clear;
    case Sheet3.TypeIndex of
      1:S3Combo.Items:=IntList3.Items;
      2:S3Combo.Items:=DecList3.Items;
      3:S3Combo.Items:=StrList3.Items;
      4:for i:=0 to (SpanList.Items.Count-1) do
        begin
          S:=SpanList.Items.Strings[i];
          if S[1]=IntToStr(Sheet3.VListIndex) then
            S3Combo.Items.Append(Copy(S,2,Length(s)-1));
        end;
    end;
    SGrid3.Col:=1;
    SGrid3.Row:=1;
    MoveCombo(SGrid3,S3Combo,1,1);
    S3Combo.Visible:=True;
  end
  else
    S3Combo.Visible:=False;
  S:=SGrid3.Cells[2,1];
  SGrid3.Col:=2;
  SGrid3.Col:=1;
  SGrid3.Cells[2,1]:=S;
  Sheet3Panel.Visible:=True;
  Sheet3Panel.BringToFront;
  ToolSheet3.Down:=True;
  IconSheet3.Down:=True;
  MSheet3Section.Checked:=True;
  ForceSheetLimitations(Sheet3,SGrid3);
  GridFindForm.Grid:=SGrid3;
  GridReplaceForm.Grid:=SGrid3;
end;

procedure TMainForm.Sheet4SectionClick(Sender: TObject);
var i:Integer; S:String; B:Boolean;
begin
  PrintToolbtn.Enabled:=True;
  CutToolbtn.Enabled:=True;
  CopyToolbtn.Enabled:=True;
  PasteToolbtn.Enabled:=True;
  FindToolbtn.Enabled:=True;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  ActiveSheet:=4;
  if SettingGrid.Col=4 then
    SettingGrid.Col:=1
  else
    SettingGrid.Col:=4;
//  Updatesheet(4);
  PrepareSheetGrid(Sheet4,SGrid4,4,4,S4FieldLabel,S4TypeLabel,S4DesLabel,S4FilterLabel,S4NLabel,S4Combo,S4Filterbtn,S4PSpin);
  if Sheet4.ForceValue then
  begin
    S4Combo.Items.Clear;
    case Sheet4.TypeIndex of
      1:S4Combo.Items:=IntList4.Items;
      2:S4Combo.Items:=DecList4.Items;
      3:S4Combo.Items:=StrList4.Items;
      4:for i:=0 to (SpanList.Items.Count-1) do
        begin
          S:=SpanList.Items.Strings[i];
          if S[1]=IntToStr(Sheet4.VListIndex) then
            S4Combo.Items.Append(Copy(S,2,Length(s)-1));
        end;
    end;
    S4Combo.Visible:=False;
    MoveCombo(SGrid4,S4Combo,1,1);
    S4Combo.Visible:=True;
  end
  else
    S4Combo.Visible:=False;
  S:=SGrid4.Cells[2,1];
{  SGrid4.Col:=2;
  SGrid4.Col:=1;}
  SGrid4.Cells[2,1]:=S;
  Sheet4Panel.Visible:=True;
  Sheet4Panel.BringToFront;
  ToolSheet4.Down:=True;
  IconSheet4.Down:=True;
  MSheet4Section.Checked:=True;
  ForceSheetLimitations(Sheet4,SGrid4);
  GridFindForm.Grid:=SGrid4;
  GridReplaceForm.Grid:=SGrid4;
end;

procedure TMainForm.Sheet5SectionClick(Sender: TObject);
var i:Integer; S:String; B:Boolean;
begin
  PrintToolbtn.Enabled:=True;
  CutToolbtn.Enabled:=True;
  CopyToolbtn.Enabled:=True;
  PasteToolbtn.Enabled:=True;
  FindToolbtn.Enabled:=True;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  ActiveSheet:=5;
  if SettingGrid.Col=4 then
    SettingGrid.Col:=1
  else
    SettingGrid.Col:=4;
//  Updatesheet(5);
  PrepareSheetGrid(Sheet5,SGrid5,5,5,S5FieldLabel,S5TypeLabel,S5DesLabel,S5FilterLabel,S5NLabel,S5Combo,S5Filterbtn,S5PSpin);
  if Sheet5.ForceValue then
  begin
    S5Combo.Items.Clear;
    case Sheet5.TypeIndex of
      1:S5Combo.Items:=IntList5.Items;
      2:S5Combo.Items:=DecList5.Items;
      3:S5Combo.Items:=StrList5.Items;
      4:for i:=0 to (SpanList.Items.Count-1) do
        begin
          S:=SpanList.Items.Strings[i];
          if S[1]=IntToStr(Sheet5.VListIndex) then
            S5Combo.Items.Append(Copy(S,2,Length(s)-1));
        end;
    end;
    SGrid5.Col:=1;
    SGrid5.Row:=1;
    MoveCombo(SGrid5,S5Combo,1,1);
    S5Combo.Visible:=True;
  end
  else
    S5Combo.Visible:=False;
  S:=SGrid5.Cells[2,1];
  SGrid5.Col:=2;
  SGrid5.Col:=1;
  SGrid5.Cells[2,1]:=S;
  Sheet5Panel.Visible:=True;
  Sheet5Panel.BringToFront;
  ToolSheet5.Down:=True;
  IconSheet5.Down:=True;
  MSheet5Section.Checked:=True;
  ForceSheetLimitations(Sheet5,SGrid5);
  GridFindForm.Grid:=SGrid5;
  GridReplaceForm.Grid:=SGrid5;
end;

procedure TMainForm.MSettingSectionClick(Sender: TObject);
begin
  if SettingPanel.Visible then
    MSettingSection.Checked:=True;
  SettingSection.Down:=True;
  SettingSection.Click;
end;

procedure TMainForm.MSheet1SectionClick(Sender: TObject);
begin
  if Sheet1Panel.Visible then
    MSheet1Section.Checked:=True;
  Sheet1Section.Down:=True;
  Sheet1Section.Click;
end;

procedure TMainForm.MSheet2SectionClick(Sender: TObject);
begin
  if Sheet2Panel.Visible then
    MSheet2Section.Checked:=True;
  Sheet2Section.Down:=True;
  Sheet2Section.Click;
end;

procedure TMainForm.MSheet3SectionClick(Sender: TObject);
begin
  if Sheet3Panel.Visible then
    MSheet3Section.Checked:=True;
  Sheet3Section.Down:=True;
  Sheet3Section.Click;
end;

procedure TMainForm.MSheet4SectionClick(Sender: TObject);
begin
  if Sheet4Panel.Visible then
    MSheet4Section.Checked:=True;
  Sheet4Section.Down:=True;
  Sheet4Section.Click;
end;

procedure TMainForm.MSheet5SectionClick(Sender: TObject);
begin
  if Sheet5Panel.Visible then
    MSheet5Section.Checked:=True;
  Sheet5Section.Down:=True;
  Sheet5Section.Click;
end;

procedure TMainForm.MTableSectionClick(Sender: TObject);
begin
  if TablePanel.Visible then
    MTableSection.Checked:=True;
  TableSection.Down:=True;
  TableSection.Click;
end;

procedure TMainForm.MChartSectionClick(Sender: TObject);
begin
  ChartSection.Down:=True;
  ChartSection.Click;
end;

procedure TMainForm.MAnalyzeSectionClick(Sender: TObject);
begin
  AnalyzeSection.Down:=True;
  AnalyzeSection.Click;
end;

procedure TMainForm.S5ComboChange(Sender: TObject);
begin
  if CanEditCell(SGrid5.Col,SGrid5.Row,1) then
    with SGrid5 do
      Cells[Col,Row]:=S5Combo.Items.Strings[S5Combo.ItemIndex];
  if Length(SGrid5.Cells[SGrid5.Col,SGrid5.Row])>0 then
    S5Combo.Hint:='Current value: '+SGrid5.Cells[SGrid5.Col,SGrid5.Row]
  else
    S5Combo.Hint:='No value is set';
end;

procedure TMainForm.S5ComboKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE	 then
  begin
    SGrid1.Cells[SGrid1.Col,SGrid1.Row]:='';
    S1Combo.Visible:=False;
  end;
end;

procedure TMainForm.SGrid52KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_Delete) and not(goEditing in SGrid5.Options) then
    SGrid5.Cells[SGrid5.Col,SGrid5.Row]:='';
end;

procedure TMainForm.SGrid52MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i,j:Integer; R:TRect; S:String;
begin
  for i:=0 to (SGrid5.ColCount-1) do
  begin
    R:=SGrid5.CellRect(i,0);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S5IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=0 to (SGrid5.RowCount-2) do
  begin
    R:=Sgrid5.CellRect(0,i);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S5IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=SGrid5.TopRow to (SGrid5.TopRow+10) do
    for j:=1 to (SGrid5.ColCount-1) do
    begin
      R:=SGrid5.CellRect(j,i);
      if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
      begin
        S:=Sgrid5.Cells[0,i]+SGrid5.Cells[j,0];
        if S[1]='0' then
          S:=Copy(S,2,Length(S)-1);
        S5IDEdit.Caption:=S;
        Exit;
      end;
    end;
  S5IDEdit.Caption:='000';
end;

procedure TMainForm.SGrid52SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  GridSelectCell(Sheet5,5,SGrid5,S5Combo,S5NLabel,ACol,ARow);
end;

procedure TMainForm.SGrid52SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
var DV:DefaultValueType;
begin
  DV.TypeIndex:=Sheet5.TypeIndex;
  DV.DValue1:=Sheet5.DValue1;
  DV.DValue2:=Sheet5.DValue2;
  DV.DValue3:=Sheet5.DValue3;
  DV.LBoundI:=Sheet5.LBoundI;
  DV.UBoundI:=Sheet5.UBoundI;
  DV.LBoundE:=Sheet5.LBoundE;
  DV.UBoundE:=Sheet5.UBoundE;
end;

procedure TMainForm.S5ClrbtnClick(Sender: TObject);
begin
  if Application.MessageBox('Are you sure you want to clear all values in data sheet 5?','Delete',MB_YesNo)=IDYes then
  begin
    SGrid5.Col:=1; SGrid5.Row:=1;
    ClearGrid(SGrid5);
    ValueCount[5]:=0;
    S5NLabel.Caption:='0';
  end;
end;

procedure TMainForm.S4ComboChange(Sender: TObject);
begin
  if CanEditCell(SGrid4.Col,SGrid4.Row,1) then
    with SGrid4 do
      Cells[Col,Row]:=S4Combo.Items.Strings[S4Combo.ItemIndex];
  if Length(SGrid4.Cells[SGrid4.Col,SGrid4.Row])>0 then
    S4Combo.Hint:='Current value: '+SGrid4.Cells[SGrid4.Col,SGrid4.Row]
  else
    S4Combo.Hint:='No value is set';
end;

procedure TMainForm.S3ComboChange(Sender: TObject);
begin
  if CanEditCell(SGrid3.Col,SGrid3.Row,1) then
    with SGrid3 do
      Cells[Col,Row]:=S3Combo.Items.Strings[S3Combo.ItemIndex];
  if Length(SGrid3.Cells[SGrid3.Col,SGrid3.Row])>0 then
    S3Combo.Hint:='Current value: '+SGrid3.Cells[SGrid3.Col,SGrid3.Row]
  else
    S3Combo.Hint:='No value is set';
end;

procedure TMainForm.S2ComboChange(Sender: TObject);
begin
  if CanEditCell(SGrid2.Col,SGrid2.Row,1) then
    with SGrid2 do
      Cells[Col,Row]:=S2Combo.Items.Strings[S2Combo.ItemIndex];
  if Length(SGrid2.Cells[SGrid2.Col,SGrid2.Row])>0 then
    S2Combo.Hint:='Current value: '+SGrid2.Cells[SGrid2.Col,SGrid2.Row]
  else
    S2Combo.Hint:='No value is set';
end;

function TMainForm.CheckGridCell(Grid:TStringGrid;TypeLabel:TStaticText):Boolean;
var S,S1:String; i,j:Integer;
begin
  Result:=True;
  S:='';
  for i:=1 to (Grid.ColCount-1) do
    for j:=1 to (Grid.RowCount-2) do
    if Length(Grid.Cells[i,j])>0 then
    begin
      S1:=Grid.Cells[i,j];
      if TypeLabel.Caption='Integer' then
      begin
        if not(IsValidInt(S1)) then
        begin
          S:='"'+S1+'" is not a valid integer value.';
          S1:='';
        end
        else
          S1:=IntToStr(StrToInt(S1));
      end;
      if (TypeLabel.Caption='Span') and not(IsValidSpan(S1)) then
      begin
        S:='"'+S1+'" is not a valid span.';
        S1:='';
      end;
      if TypeLabel.Caption='Decimal' then
      begin
        if not(IsValidDec(S1)) then
        begin
          S:='"'+S1+'" is not a valid decimal value.';
          S1:='';
        end
        else
        begin
          S1:=FloatToStr(StrToFloat(S1));
          if S1='0' then S1:='0.0';
          if not(Pos('.',S1)>0) then
            S1:=S1+'.0';
        end;
      end;
      if S<>'' then
      begin
        Result:=False;
        ShowMessage(S); S:='';
      end;
      if S1<>Grid.Cells[i,j] then
        Grid.Cells[i,j]:=S1;
    end;
end;

procedure TMainForm.SGrid4SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var S:String;
begin
{  Debugging Expressions:
  if CanSelect then
    ShowMessage('CanSelect=True'+Chr(13)+'Row='+IntToStr(ARow)+Chr(13)+'Col='+IntToStr(ACol))
  else
    ShowMessage('CanSelect=False'+Chr(13)+'Row='+IntToStr(ARow)+Chr(13)+'Col='+IntToStr(ACol));
  Exit;}
  if SGrid4.Tag=5 then
  begin
    CanSelect:=True;
    Exit;
  end;
  GridSelectCell(Sheet4,4,SGrid4,S4Combo,S4NLabel,ACol,ARow);
  if CheckGridCell(SGrid4,S4TypeLabel)=False then
  begin
    CanSelect:=False;
    Exit;
  end;
  case Sheet4.TypeIndex of
    1:S:=IntToStr(Sheet4.DValue1);
    2:begin
        S:=FloatToStr(Sheet4.DValue2);
        if (Pos('E',S)>0) or (Pos('e',S)>0) then S:='0.00'
        else if not(Pos('.',S)>0) then S:=S+'.0';
      end;
    3:S:=Sheet4.DValue3;
    4:S:='['+IntToStr(Sheet4.LBoundI)+','+IntToStr(Sheet4.UBoundI)+']';
  end;
  if S4Combo.Visible then S4Combo.Visible:=False;
  if {not(goEditing in SGrid4.Options)}Sheet4.ForceValue and CanEditCell(ACol,ARow,ActiveSheet) then
  begin
    if not(DVCheck.Checked) then
    begin
      if S4Combo.Items.IndexOf(S)>0 then
        S4Combo.Items.Delete(S4Combo.Items.IndexOf(S));
      if S4Combo.Items.Count>0 then
        S4Combo.ItemIndex:=0;
    end;
    if DVCheck.Checked then
    begin
      if not(S4Combo.Items.IndexOf(S)>=0) then
        S4Combo.Items.Add(S);
      if (Length(SGrid4.Cells[ACol,ARow])>0) and (S4Combo.Items.IndexOf(SGrid4.Cells[ACol,ARow])>=0) then
        S4Combo.ItemIndex:=S4Combo.Items.IndexOf(SGrid4.Cells[ACol,ARow])
      else
        S4Combo.ItemIndex:=S4Combo.Items.IndexOf(S);
    end;
    S4Combo.Visible:=False;
    MoveCombo(SGrid4,S4Combo,ACol,ARow);
    S4Combo.Visible:=True;
    S4Combo.SetFocus;
    S4Combo.SelectAll;
  end;
  if DVCheck.Checked and (Length(SGrid4.Cells[ACol,ARow])=0) then
    SGrid4.Cells[ACol,ARow]:=S;
  if Length(SGrid4.Cells[ACol,ARow])>0 then
    S4Combo.Hint:='Current value: '+SGrid4.Cells[ACol,ARow]
  else
    S4Combo.Hint:='No value is set';
end;

procedure TMainForm.SGrid3SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var S:String;
begin
  if SGrid3.Tag=5 then
  begin
    CanSelect:=True;
    Exit;
  end;
  GridSelectCell(Sheet3,3,SGrid3,S3Combo,S3NLabel,ACol,ARow);
  if CheckGridCell(SGrid3,S3TypeLabel)=False then
  begin
    CanSelect:=False;
    Exit;
  end;
  case Sheet3.TypeIndex of
    1:S:=IntToStr(Sheet3.DValue1);
    2:begin
        S:=FloatToStr(Sheet3.DValue2);
        if (Pos('E',S)>0) or (Pos('e',S)>0) then S:='0.00';
      end;
    3:S:=Sheet3.DValue3;
    4:S:='['+IntToStr(Sheet3.LBoundI)+','+IntToStr(Sheet3.UBoundI)+']';
  end;
  if S3Combo.Visible then S3Combo.Visible:=False;
  if {not(goEditing in SGrid3.Options)}Sheet3.ForceValue and CanEditCell(ACol,ARow,ActiveSheet) then
  begin
    if not(DVCheck.Checked) then
    begin
      if S3Combo.Items.IndexOf(S)>0 then
        S3Combo.Items.Delete(S3Combo.Items.IndexOf(S));
      if S3Combo.Items.Count>0 then
        S3Combo.ItemIndex:=0;
    end;
    if DVCheck.Checked then
    begin
      if not(S3Combo.Items.IndexOf(S)>=0) then
        S3Combo.Items.Add(S);
      if (Length(SGrid3.Cells[ACol,ARow])>0) and (S3Combo.Items.IndexOf(SGrid3.Cells[ACol,ARow])>=0) then
        S3Combo.ItemIndex:=S3Combo.Items.IndexOf(SGrid3.Cells[ACol,ARow])
      else
        S3Combo.ItemIndex:=S3Combo.Items.IndexOf(S);
    end;
    S3Combo.Visible:=False;
    MoveCombo(SGrid3,S3Combo,ACol,ARow);
    S3Combo.Visible:=True;
    S3Combo.SetFocus;
    S3Combo.SelectAll;
  end;
  if DVCheck.Checked and (Length(SGrid3.Cells[ACol,ARow])=0) then
    SGrid3.Cells[ACol,ARow]:=S;
  if Length(SGrid3.Cells[ACol,ARow])>0 then
    S3Combo.Hint:='Current value: '+SGrid3.Cells[ACol,ARow]
  else
    S3Combo.Hint:='No value is set';
end;

procedure TMainForm.SGrid2SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var S:String;
begin
  if SGrid2.Tag=5 then
  begin
    CanSelect:=True;
    Exit;
  end;
  GridSelectCell(Sheet2,2,SGrid2,S2Combo,S2NLabel,ACol,ARow);
  if CheckGridCell(SGrid2,S2TypeLabel)=False then
  begin
    CanSelect:=False;
    Exit;
  end;
  case Sheet2.TypeIndex of
    1:S:=IntToStr(Sheet2.DValue1);
    2:begin
        S:=FloatToStr(Sheet2.DValue2);
        if (Pos('E',S)>0) or (Pos('e',S)>0) then S:='0.00';
      end;
    3:S:=Sheet2.DValue3;
    4:S:='['+IntToStr(Sheet2.LBoundI)+','+IntToStr(Sheet2.UBoundI)+']';
  end;
  if S2Combo.Visible then S2Combo.Visible:=False;
  if {not(goEditing in SGrid2.Options)}Sheet2.ForceValue and CanEditCell(ACol,ARow,ActiveSheet) then
  begin
    if not(DVCheck.Checked) then
    begin
      if S2Combo.Items.IndexOf(S)>0 then
        S2Combo.Items.Delete(S2Combo.Items.IndexOf(S));
      if S2Combo.Items.Count>0 then
        S2Combo.ItemIndex:=0;
    end;
    if DVCheck.Checked then
    begin
      if not(S2Combo.Items.IndexOf(S)>=0) then
        S2Combo.Items.Add(S);
      if (Length(SGrid2.Cells[ACol,ARow])>0) and (S2Combo.Items.IndexOf(SGrid2.Cells[ACol,ARow])>=0) then
        S2Combo.ItemIndex:=S2Combo.Items.IndexOf(SGrid2.Cells[ACol,ARow])
      else
        S2Combo.ItemIndex:=S2Combo.Items.IndexOf(S);
    end;
    S2Combo.Visible:=False;
    MoveCombo(SGrid2,S2Combo,ACol,ARow);
    S2Combo.Visible:=True;
    S2Combo.SetFocus;
    S2Combo.SelectAll;
  end;
  if DVCheck.Checked and (Length(SGrid2.Cells[ACol,ARow])=0) then
    SGrid2.Cells[ACol,ARow]:=S;
  if Length(SGrid2.Cells[ACol,ARow])>0 then
    S2Combo.Hint:='Current value: '+SGrid2.Cells[ACol,ARow]
  else
    S2Combo.Hint:='No value is set';
end;

procedure TMainForm.SGrid4MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i,j:Integer; R:TRect; S:String;
begin
  for i:=0 to (SGrid4.ColCount-1) do
  begin
    R:=SGrid4.CellRect(i,0);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S4IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=0 to (SGrid4.RowCount-2) do
  begin
    R:=Sgrid4.CellRect(0,i);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S4IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=SGrid4.TopRow to (SGrid4.RowCount-2) do
    for j:=1 to (SGrid4.ColCount-1) do
    begin
      R:=SGrid4.CellRect(j,i);
      if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
      begin
        S:=Sgrid4.Cells[0,i]+SGrid4.Cells[j,0];
        if S[1]='0' then
          S:=Copy(S,2,Length(S)-1);
        S4IDEdit.Caption:=S;
        Exit;
      end;
    end;
  S4IDEdit.Caption:='000';
end;

procedure TMainForm.SGrid3MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i,j:Integer; R:TRect; S:String;
begin
  for i:=0 to (SGrid3.ColCount-1) do
  begin
    R:=SGrid3.CellRect(i,0);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S3IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=0 to (SGrid3.RowCount-2) do
  begin
    R:=Sgrid3.CellRect(0,i);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S3IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=SGrid3.TopRow to (SGrid3.RowCount-2) do
    for j:=1 to (SGrid3.ColCount-1) do
    begin
      R:=SGrid3.CellRect(j,i);
      if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
      begin
        S:=Sgrid3.Cells[0,i]+SGrid3.Cells[j,0];
        if S[1]='0' then
          S:=Copy(S,2,Length(S)-1);
        S3IDEdit.Caption:=S;
        Exit;
      end;
    end;
  S3IDEdit.Caption:='000';
end;

procedure TMainForm.SGrid2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i,j:Integer; R:TRect; S:String;
begin
  for i:=0 to (SGrid2.ColCount-1) do
  begin
    R:=SGrid2.CellRect(i,0);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S2IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=0 to (SGrid2.RowCount-2) do
  begin
    R:=Sgrid2.CellRect(0,i);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S2IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=SGrid2.TopRow to (SGrid2.RowCount-2) do
    for j:=1 to (SGrid2.ColCount-1) do
    begin
      R:=SGrid2.CellRect(j,i);
      if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
      begin
        S:=Sgrid2.Cells[0,i]+SGrid2.Cells[j,0];
        if S[1]='0' then
          S:=Copy(S,2,Length(S)-1);
        S2IDEdit.Caption:=S;
        Exit;
      end;
    end;
  S2IDEdit.Caption:='000';
end;

procedure TMainForm.SGrid4KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_Delete) and not(goEditing in SGrid4.Options) then
    SGrid4.Cells[SGrid4.Col,SGrid4.Row]:='';
end;

procedure TMainForm.SGrid3KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_Delete) and not(goEditing in SGrid3.Options) then
    SGrid3.Cells[SGrid3.Col,SGrid3.Row]:='';
end;

procedure TMainForm.SGrid2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_Delete) and not(goEditing in SGrid2.Options) then
    SGrid2.Cells[SGrid2.Col,SGrid2.Row]:='';
end;

procedure TMainForm.S4ClrbtnClick(Sender: TObject);
begin
  if Application.MessageBox('Are you sure you want to clear all values in data sheet 4?','Delete',MB_YesNo)=IDYes then
  begin
    SGrid4.Col:=1; SGrid4.Row:=1;
    ClearGrid(SGrid4);
    ValueCount[4]:=0;
    S4NLabel.Caption:='0';
  end;
end;

procedure TMainForm.S3ClrbtnClick(Sender: TObject);
begin
  if Application.MessageBox('Are you sure you want to clear all values in data sheet 3?','Delete',MB_YesNo)=IDYes then
  begin
    SGrid3.Col:=1; SGrid3.Row:=1;
    ClearGrid(SGrid3);
    ValueCount[3]:=0;
    S3NLabel.Caption:='0';
  end;
end;

procedure TMainForm.S2ClrbtnClick(Sender: TObject);
begin
  if Application.MessageBox('Are you sure you want to clear all values in data sheet 2?','Delete',MB_YesNo)=IDYes then
  begin
    SGrid2.Col:=1; SGrid2.Row:=1;
    ClearGrid(SGrid2);
    ValueCount[2]:=0;
    S2NLabel.Caption:='0';
  end;
end;

procedure TMainForm.S4ComboKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE	 then
  begin
    SGrid4.Cells[SGrid4.Col,SGrid4.Row]:='';
    S4Combo.Visible:=False;
  end;
end;

procedure TMainForm.S3ComboKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE	 then
  begin
    SGrid3.Cells[SGrid3.Col,SGrid3.Row]:='';
    S3Combo.Visible:=False;
  end;
end;

procedure TMainForm.S2ComboKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_DELETE	 then
  begin
    SGrid2.Cells[SGrid2.Col,SGrid2.Row]:='';
    S2Combo.Visible:=False;
  end;
end;

procedure TMainForm.IconViewClick(Sender: TObject);
begin
  ListView.Click;
  IconView.Checked:=True;
  Shape1.Height:=363;
  Shape1.Top:=16;
  MemberButtonNormal.Top:=21;
  MemberButtonPressed.Top:=21;
  MemberPaneImage.Height:=320;
  MemberPaneImage.Top:=56;
  MainForm.Width:=803;
  SettingPanel.Left:=184;
  Sheet1Panel.Left:=184;
  Sheet2Panel.Left:=184;
  Sheet3Panel.Left:=184;
  Sheet4Panel.Left:=184;
  Sheet5Panel.Left:=184;
  TablePanel.Left:=184;
  ChartPanel.Left:=184;
  AnalyzePanel.Left:=184;
  MemberToolbar.Visible:=False;;
  MemberList.Visible:=False;
  MemberIcon.Visible:=True;
  MemberGroup.Width:=177;
  ResizeForm;
end;

procedure TMainForm.ToolViewClick(Sender: TObject);
begin
  if MemberToolbar.Visible then
    ToolView.Checked:=True;
  Shape1.Height:=363;
  Shape1.Top:=16;
  MemberButtonNormal.Top:=21;
  MemberButtonPressed.Top:=21;
  MemberPaneImage.Height:=320;
  MemberPaneImage.Top:=56;
  MainForm.Width:=720;
  SettingPanel.Left:=101;
  Sheet1Panel.Left:=101;
  Sheet2Panel.Left:=101;
  Sheet3Panel.Left:=101;
  Sheet4Panel.Left:=101;
  Sheet5Panel.Left:=101;
  TablePanel.Left:=101;
  ChartPanel.Left:=101;
  AnalyzePanel.Left:=101;
  MemberList.Visible:=False;
  MemberIcon.Visible:=False;
  MemberToolbar.Visible:=True;
  MemberGroup.Width:=88;
  ResizeForm;
end;

procedure TMainForm.ListViewClick(Sender: TObject);
begin
  if MemberList.Visible then
    ListView.Checked:=True;
  Shape1.Height:=363;
  Shape1.Top:=16;
  MemberButtonNormal.Top:=21;
  MemberButtonPressed.Top:=21;
  MemberPaneImage.Height:=320;
  MemberPaneImage.Top:=56;
  PaneMenuBtn.Visible:=False;
  MainForm.Width:=803;
  SettingPanel.Left:=184;
  Sheet1Panel.Left:=184;
  Sheet2Panel.Left:=184;
  Sheet3Panel.Left:=184;
  Sheet4Panel.Left:=184;
  Sheet5Panel.Left:=184;
  TablePanel.Left:=184;
  ChartPanel.Left:=184;
  AnalyzePanel.Left:=184;
  MemberToolbar.Visible:=False;;
  MemberIcon.Visible:=False;;
  MemberList.Visible:=True;
  MemberGroup.Width:=177;
  ResizeForm;
end;

procedure TMainForm.MenuViewClick(Sender: TObject);
begin
  if PaneMenuBtn.Visible then
    MenuView.Checked:=True;
  Shape1.Height:=331;
  Shape1.Top:=48;
  MemberButtonNormal.Top:=53;
  MemberButtonPressed.Top:=53;
  MemberGroup.Width:=44;
  MemberPaneImage.Top:=88;
  MemberPaneImage.Height:=288;
  SettingPanel.Left:=51;
  Sheet1Panel.Left:=51;
  Sheet2Panel.Left:=51;
  Sheet3Panel.Left:=51;
  Sheet4Panel.Left:=51;
  Sheet5Panel.Left:=51;
  TablePanel.Left:=51;
  ChartPanel.Left:=51;
  AnalyzePanel.Left:=51;
  MemberIcon.Visible:=False;
  MemberList.Visible:=False;
  MemberToolbar.Visible:=False;
  PaneMenuBtn.Visible:=True;
  MainForm.Width:=673;
  ResizeForm;
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
  IconView.Click;
    SGrid1.Repaint;
    AutoSizeGridRows(SGrid1);
end;

procedure TMainForm.ToolSettingsClick(Sender: TObject);
begin
  SettingSection.Down:=True;
  SettingSection.Click;
end;

procedure TMainForm.ToolSheet1Click(Sender: TObject);
begin
  Sheet1Section.Down:=True;
  Sheet1Section.Click;
  if SGrid1.Tag=10 then
  begin
    SGrid1.Repaint;
    AutoSizeGridRows(SGrid1);
    SGrid1.Tag:=0;
  end;
end;

procedure TMainForm.ToolSheet2Click(Sender: TObject);
begin
  Sheet2Section.Down:=True;
  Sheet2Section.Click;
  if SGrid2.Tag=10 then
  begin
    SGrid2.Repaint;
    AutoSizeGridRows(SGrid2);
    SGrid2.Tag:=0;
  end;
end;

procedure TMainForm.ToolSheet3Click(Sender: TObject);
begin
  Sheet3Section.Down:=True;
  Sheet3Section.Click;
  if SGrid3.Tag=10 then
  begin
    SGrid3.Repaint;
    AutoSizeGridRows(SGrid3);
    SGrid3.Tag:=0;
  end;
end;

procedure TMainForm.ToolSheet4Click(Sender: TObject);
begin
  Sheet4Section.Down:=True;
  Sheet4Section.Click;
  if SGrid4.Tag=10 then
  begin
    SGrid4.Repaint;
    AutoSizeGridRows(SGrid4);
    SGrid4.Tag:=0;
  end;
end;

procedure TMainForm.ToolSheet5Click(Sender: TObject);
begin
  Sheet5Section.Down:=True;
  Sheet5Section.Click;
  if SGrid5.Tag=10 then
  begin
    SGrid5.Repaint;
    AutoSizeGridRows(SGrid5);
    SGrid5.Tag:=0;
  end;
end;

procedure TMainForm.ToolTableClick(Sender: TObject);
begin
  TableSection.Down:=True;
  TableSection.Click;
end;

procedure TMainForm.ToolChartClick(Sender: TObject);
begin
  ChartSection.Down:=True;
  ChartSection.Click;
end;

procedure TMainForm.ToolAnalyzeClick(Sender: TObject);
begin
  AnalyzeSection.Down:=True;
  AnalyzeSection.Click;
end;

procedure TMainForm.Edit2Change(Sender: TObject);
begin
  HeadLabel.Caption:=(Sender as TEdit).Text;
end;

procedure TMainForm.Edit3Change(Sender: TObject);
begin
  FootLabel.Caption:=(Sender as TEdit).Text;
end;

procedure TMainForm.SpinEdit1Change(Sender: TObject);
begin
  TableGrid.GridLineWidth:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.CheckBox1Click(Sender: TObject);
begin
  if (Sender as TCheckBox).Checked then
    TableGrid.Options:=TableGrid.Options+[goVertLine]
  else
    TableGrid.Options:=TableGrid.Options-[goVertLine]
end;

procedure TMainForm.CheckBox2Click(Sender: TObject);
begin
  if (Sender as TCheckBox).Checked then
    TableGrid.Options:=TableGrid.Options+[goHorzLine]
  else
    TableGrid.Options:=TableGrid.Options-[goHorzLine]
end;

procedure TMainForm.CheckBox3Click(Sender: TObject);
begin
  if (Sender as TCheckBox).Checked then
    TableGrid.FixedRows:=1
  else
    TableGrid.FixedRows:=0;
end;

procedure TMainForm.CheckBox4Click(Sender: TObject);
begin
  if (Sender as TCheckBox).Checked then
    TableGrid.FixedCols:=1
  else
    TableGrid.FixedCols:=0;
end;

procedure TMainForm.CheckBox5Click(Sender: TObject);
begin
  if (Sender as TCheckBox).Checked then
    TableGrid.BorderStyle:=bsSingle
  else
    TableGrid.BorderStyle:=bsNone;  
end;

procedure TMainForm.TableSectionClick(Sender: TObject);
var B:Boolean;
begin
  PrintToolbtn.Enabled:=False;
  CutToolbtn.Enabled:=False;
  CopyToolbtn.Enabled:=False;
  PasteToolbtn.Enabled:=False;
  FindToolbtn.Enabled:=False;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  TablePanel.Visible:=True;
  TablePanel.BringToFront;
  IconTable.Down:=True;
  ToolTable.Down:=True;
  MTableSection.Checked:=True;
  ActiveSheet:=6;
end;

procedure TMainForm.IconSettingsMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  Application.HintPause:=0;
  Application.HintHidePause:=6000;
  IconSettings.Hint:='Used data sheets:0'+Chr(13)+'Type: Integer';
end;

procedure TMainForm.MemberGroupMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  Application.HintHidePause:=HHP;
  Application.HintPause:=HP;
end;

procedure TMainForm.FormMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  Application.HintHidePause:=HHP;
  Application.HintPause:=HP;
end;

procedure TMainForm.CreateTable2Click(Sender: TObject);
var S:String;
    grid:TAdvStringGrid;
begin
  TableForm.Visible:=not(TableForm.Visible);
{  with TableForm do
  begin
    BaseCombo1.Enabled:=True;
    BaseCombo2.Enabled:=True;
    AssignPanel.Enabled:=True;
    BaseCombo1.Items.Clear;
    BaseCombo2.Items.Clear;
    SpanCombo.Items.Clear;
    ItemList.Items.Clear;
    SpanList.Items.Clear;
    AutoCheck.Checked:=True;
    FillComboSpecific(BaseCombo1,1);
    FillComboSpecific(BaseCombo1,2);
    FillComboSpecific(BaseCombo2,3);
    FillComboSpecific(BaseCombo2,4);
    FillComboSpecific(SpanCombo,4);
    if SpanCombo.Items.Count>0 then
      SpanCombo.ItemIndex:=0;
  end;
  if (TableForm.BaseCombo1.Items.Count=1) and (TableForm.BaseCombo2.Items.Count=0) then
  begin
    if Sheet1.Used then FillValueList(SGrid1)
    else if Sheet2.Used then FillValueList(SGrid2)
    else if Sheet3.Used then FillValueList(SGrid3)
    else if Sheet4.Used then FillValueList(SGrid4)
    else if Sheet5.Used then FillValueList(SGrid5);
    if VListCount<2 then
    begin
      ShowMessage('You should enter at least two numbers in a sheet to create the frequency table.');
      Exit;
    end;
    FillRangeList;
    CreateTable;
    CanCreateAChart:=True;
    Exit;
  end;
  if TableForm.BaseCombo1.Items.Count=0 then
  begin
    ShowMessage('There is no data sheet with integer or decimal data type.To create a table, you should make a data sheet with integer or decimal data type to provide the table frequency values.');
    Exit;
  end;
  if TableForm.BaseCombo2.Items.Count=0 then
  begin
    TableForm.AutoCheck.Enabled:=False;
    TableForm.BaseCombo2.Enabled:=False;
    TableForm.BaseCombo1.ItemIndex:=0;
    if TableForm.ShowModal=mrOK then
    begin
      S:=TableForm.BaseCombo1.Items.Strings[TableForm.BaseCombo1.ItemIndex];
      if S='Data Sheet 1' then FillValueList(SGrid1)
      else if S='Data Sheet 2' then FillValueList(SGrid2)
      else if S='Data Sheet 3' then FillValueList(SGrid3)
      else if S='Data Sheet 4' then FillValueList(SGrid4)
      else if S='Data Sheet 5' then FillValueList(SGrid5);
      FillRangeList;
      CreateTable;
      CanCreateAChart:=True;
    end;
    Exit;
  end;
  TableForm.BaseCombo2.Enabled:=False;
  with TableForm do
  begin
    AutoCheck.Checked:=True;
    BaseCombo2.Enabled:=False;
    AssignPanel.Enabled:=False;
    BaseCombo2.Items.Clear;
    FillComboSpecific(BaseCombo2,3);
    FillComboSpecific(BaseCombo2,4);
    BaseCombo2.ItemIndex:=0;
    BaseCombo1.ItemIndex:=0;
    if TableForm.ShowModal=mrOK then
    begin
      if TableUnit.CanCreate then
      begin
        grid:=GetGrid(BaseCombo1.Items.Strings[BaseCombo1.ItemIndex]);
        FillValueList(grid);
        if AutoCheck.Checked then
          FillRangeList;
        CreateTable;
        CanCreateAChart:=True;
      end;
    end;
  end;}
end;

procedure TMainForm.ApplicationEvents1ShowHint(var HintStr: String;
  var CanShow: Boolean; var HintInfo: THintInfo);
var S:String; N:Byte;
begin
  if HintInfo.HintControl=LegendStyle then
    HintInfo.HideTimeout:=600000;
  if (sheet1.Used) and ((HintInfo.HintControl=Sheet1Section) or (HintInfo.HintControl=IconSheet1) or (HintInfo.HintControl=ToolSheet1)) then
  begin
    HintInfo.HideTimeout:=15000;
    S:='*Data Sheet 1'+Chr(13)+'Field Name: '+
        SettingGrid.Cells[1,1]+Chr(13)+'Type: '+
        SettingGrid.Cells[2,1]+Chr(13)+'Force To Choose Value: ';
    if Sheet1.ForceValue then S:=S+'Yes' else S:=S+'No';
    S:=S+Chr(13)+'Value Count(N)= '+S1NLabel.Caption;
    HintInfo.HintStr:=S;
  end
  else if (sheet2.Used) and ((HintInfo.HintControl=Sheet2Section) or (HintInfo.HintControl=IconSheet2) or (HintInfo.HintControl=ToolSheet2)) then
  begin
    HintInfo.HideTimeout:=15000;
    S:='*Data Sheet 2'+Chr(13)+'Field Name: '+
        SettingGrid.Cells[1,2]+Chr(13)+'Type: '+
        SettingGrid.Cells[2,2]+Chr(13)+'Force To Choose Value: ';
    if Sheet2.ForceValue then S:=S+'Yes' else S:=S+'No';
    S:=S+Chr(13)+'Value Count(N)= '+S2NLabel.Caption;
    HintInfo.HintStr:=S;
  end
  else if (sheet3.Used) and ((HintInfo.HintControl=Sheet3Section) or (HintInfo.HintControl=IconSheet3) or (HintInfo.HintControl=ToolSheet3)) then
  begin
    HintInfo.HideTimeout:=15000;
    S:='*Data Sheet 3'+Chr(13)+'Field Name: '+
        SettingGrid.Cells[1,3]+Chr(13)+'Type: '+
        SettingGrid.Cells[2,3]+Chr(13)+'Force To Choose Value: ';
    if Sheet3.ForceValue then S:=S+'Yes' else S:=S+'No';
    S:=S+Chr(13)+'Value Count(N)= '+S3NLabel.Caption;
    HintInfo.HintStr:=S;
  end
  else if (sheet4.Used) and ((HintInfo.HintControl=Sheet4Section) or (HintInfo.HintControl=IconSheet4) or (HintInfo.HintControl=ToolSheet4)) then
  begin
    HintInfo.HideTimeout:=15000;
    S:='*Data Sheet 4'+Chr(13)+'Field Name: '+
        SettingGrid.Cells[1,4]+Chr(13)+'Type: '+
        SettingGrid.Cells[2,4]+Chr(13)+'Force To Choose Value: ';
    if Sheet4.ForceValue then S:=S+'Yes' else S:=S+'No';
    S:=S+Chr(13)+'Value Count(N)= '+S4NLabel.Caption;
    HintInfo.HintStr:=S;
  end
  else if (sheet5.Used) and ((HintInfo.HintControl=Sheet5Section) or (HintInfo.HintControl=IconSheet5) or (HintInfo.HintControl=ToolSheet5)) then
  begin
    HintInfo.HideTimeout:=15000;
    S:='*Data Sheet 5'+Chr(13)+'Field Name: '+
        SettingGrid.Cells[1,5]+Chr(13)+'Type: '+
        SettingGrid.Cells[2,5]+Chr(13)+'Force To Choose Value: ';
    if Sheet5.ForceValue then S:=S+'Yes' else S:=S+'No';
    S:=S+Chr(13)+'Value Count(N)= '+S5NLabel.Caption;
    HintInfo.HintStr:=S;
  end
  else if (HintInfo.HintControl=SettingSection) or (HintInfo.HintControl=IconSettings) or (HintInfo.HintControl=ToolSettings) then
  begin
    HintInfo.HideTimeout:=11000;
    S:='Data Settings'+chr(13)+'Includes:';
    N:=0;
    if Sheet1.Used then
    begin
      Inc(N);
      S:=S+Chr(13)+'  '+IntTostr(N)+')Data Sheet 1: '+SettingGrid.Cells[1,1]+'('+SettingGrid.Cells[2,1]+')';
    end;
    if Sheet2.Used then
    begin
      Inc(N);
      S:=S+Chr(13)+'  '+IntTostr(N)+')Data Sheet 1: '+SettingGrid.Cells[1,2]+'('+SettingGrid.Cells[2,2]+')';
    end;
    if Sheet3.Used then
    begin
      Inc(N);
      S:=S+Chr(13)+'  '+IntTostr(N)+')Data Sheet 1: '+SettingGrid.Cells[1,3]+'('+SettingGrid.Cells[2,3]+')';
    end;
    if Sheet4.Used then
    begin
      Inc(N);
      S:=S+Chr(13)+'  '+IntTostr(N)+')Data Sheet 1: '+SettingGrid.Cells[1,4]+'('+SettingGrid.Cells[2,4]+')';
    end;
    if Sheet5.Used then
    begin
      Inc(N);
      S:=S+Chr(13)+'  '+IntTostr(N)+')Data Sheet 1: '+SettingGrid.Cells[1,5]+'('+SettingGrid.Cells[2,5]+')';
    end;
    HintInfo.HintStr:=S;
  end
  else if HintInfo.HintControl=FastResult then
  begin
    if FastResult.Hint='You can see the calculated result here.' then
      HintInfo.HideTimeout:=3000
    else
      HintInfo.HideTimeout:=7000;
  end;
end;

procedure TMainForm.AnalyzeSectionClick(Sender: TObject);
var S:String; B:Boolean;
begin
  PrintToolbtn.Enabled:=False;
  CutToolbtn.Enabled:=False;
  CopyToolbtn.Enabled:=False;
  PasteToolbtn.Enabled:=False;
  FindToolbtn.Enabled:=False;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  if TableIsQualitative then
  begin
    ShowMessage('For qualitative data types there is no analysis available.');
    GotoActiveSheet;
    Exit;
  end;
  if VListCount<=1 then
  begin
    ShowMessage('There is no enough data to analyze.');
    TableSection.Click;
    TableSection.Down:=True;
    ToolTable.Down:=True;
    IConTable.Down:=True;
    MTableSection.Checked:=True;
  end
  else
  begin
    AnalyzePanel.Visible:=True;
    AnalyzePanel.BringToFront;
    IconAnalyze.Down:=True;
    ToolAnalyze.Down:=True;
    MAnalyzeSection.Checked:=True;
    ActiveSheet:=8;
    Application.ProcessMessages;
    AnalyzeData;
    {After AnalyzeData section, ValueList array is sorted ascending}
    anaNLabel.Caption:=IntToStr(VListCount);
    S:=FloatToStr(ValueList[1]);
    if not(Pos('E',S)>0) and not(Pos('e',S)>0) and (Length(Copy(S,1,Length(S)-Pos('.',S)))>4) then
      S:=Copy(S,1,Pos('.',S)+4);
    anaMinLabel.Caption:=S;
    anaMaxLabel.Caption:=FloatToStr(ValueList[VListCount]);
    S:=FloatToStr(ValueList[VListCount]-ValueList[1]);
    if not(Pos('E',S)>0) and not(Pos('e',S)>0) and (Length(Copy(S,1,Length(S)-Pos('.',S)))>4) then
      S:=Copy(S,1,Pos('.',S)+4);
    anaRangeLabel.Caption:=S;
  end;
end;

procedure TMainForm.ModeListTimerTimer(Sender: TObject);
begin
{Timer}
  if (ModeList.Height<=5) and (ModeListVisible=False) then
  begin
    ModeListTimer.Enabled:=False;
    ModeList.Visible:=False;
    Exit;
  end
  else if (ModeList.Height>=89) and (ModeListVisible=True) then
  begin
    ModeListTimer.Enabled:=False;
    Exit;
  end;
  if ModeListVisible then
  begin
    ModeList.Top:=ModeList.Top-4;
    ModeList.Height:=ModeList.Height+4;
  end
  else
  begin
    ModeList.Top:=ModeList.Top+4;
    ModeList.Height:=ModeList.Height-4;
  end;  
end;

procedure TMainForm.SpeedButton1Click(Sender: TObject);
begin
  ShowModeList(True);
end;

procedure TMainForm.ModeBtnClick(Sender: TObject);
begin
  if ModeList.Visible then
    ShowModeList(False)
  else
    ShowModeList(True);
end;


procedure TMainForm.ExplorerButton8Click(Sender: TObject);
begin
  (Sender as TExplorerButton).Down:=not((Sender as TExplorerButton).Down);
end;

procedure TMainForm.Splitter2CanResize(Sender: TObject;
  var NewSize: Integer; var Accept: Boolean);
begin
  Accept:=False;
end;

procedure TMainForm.Splitter3Moved(Sender: TObject);
begin
  (Sender as TSplitter).Left:=72;
end;

procedure TMainForm.SizeComboChange(Sender: TObject);
begin
  ChartTitleEdit.Font.Size:=StrToInt(SizeCombo.Text);
  Chart1.Title.Font:=ChartTitleEdit.Font;
  ChartForm.Chart1.Title.Font:=ChartTitleEdit.Font;
end;

procedure TMainForm.ColorBtn1ColorSelected(Sender: TObject;
  AColor: TColor);
begin
  ChartTitleEdit.Font.Color:=AColor;
  Chart1.Title.Font:=ChartTitleEdit.Font;
  ChartForm.Chart1.Title.Font:=ChartTitleEdit.Font;
end;

procedure TMainForm.BoldBtn1Click(Sender: TObject);
begin
  if BoldBtn1.Down then
    ChartTitleEdit.Font.Style:=ChartTitleEdit.Font.Style+[fsBold]
  else
    ChartTitleEdit.Font.Style:=ChartTitleEdit.Font.Style-[fsBold];
  Chart1.Title.Font:=ChartTitleEdit.Font;
  ChartForm.Chart1.Title.Font:=ChartTitleEdit.Font;  
end;

procedure TMainForm.ItalicBtn1Click(Sender: TObject);
begin
  if ItalicBtn1.Down then
    ChartTitleEdit.Font.Style:=ChartTitleEdit.Font.Style+[fsItalic]
  else
    ChartTitleEdit.Font.Style:=ChartTitleEdit.Font.Style-[fsItalic];
  Chart1.Title.Font:=ChartTitleEdit.Font;
  ChartForm.Chart1.Title.Font:=ChartTitleEdit.Font;
end;

procedure TMainForm.UnderBtn1Click(Sender: TObject);
begin
  if UnderBtn1.Down then
    ChartTitleEdit.Font.Style:=ChartTitleEdit.Font.Style+[fsUnderLine]
  else
    ChartTitleEdit.Font.Style:=ChartTitleEdit.Font.Style-[fsUnderLine];
  Chart1.Title.Font:=ChartTitleEdit.Font;
  ChartForm.Chart1.Title.Font:=ChartTitleEdit.Font;
end;

procedure TMainForm.LeftAlClick(Sender: TObject);
begin
  if LeftAl.Down then
  begin
    ChartTitleEdit.Alignment:=taLeftJustify;
    Chart1.Title.Alignment:=taLeftJustify;
    ChartForm.Chart1.Title.Alignment:=taLeftJustify;
  end;
end;

procedure TMainForm.CenterAlClick(Sender: TObject);
begin
  if CenterAl.Down then
  begin
    ChartTitleEdit.Alignment:=taCenter;
    Chart1.Title.Alignment:=taCenter;
    ChartForm.Chart1.Title.Alignment:=taCenter;
  end;  
end;

procedure TMainForm.RightAlClick(Sender: TObject);
begin
  if RightAl.Down then
  begin
    ChartTitleEdit.Alignment:=taRightJustify;
    Chart1.Title.Alignment:=taRightJustify;
    ChartForm.Chart1.Title.Alignment:=taRightJustify;
  end;  
end;

procedure TMainForm.SpeedButton9Click(Sender: TObject);
begin
  Memo:=ChartTitleEdit;
  FD1.Font:=ChartTitleEdit.Font;
  if FD1.Execute then
  begin
    ChartTitleEdit.Font:=FD1.Font;
    if fsBold in FD1.Font.Style then
      BoldBtn1.Down:=True
    else
      BoldBtn1.Down:=False;
    if fsItalic in FD1.Font.Style then
      ItalicBtn1.Down:=True
    else
      ItalicBtn1.Down:=False;
    if fsUnderLine in FD1.Font.Style then
      UnderBtn1.Down:=True
    else
      UnderBtn1.Down:=False;
    ColorBtn1.SelectedColor:=FD1.Font.Color;
    SizeCombo.ItemIndex:=SizeCombo.Items.IndexOf(IntToStr(FD1.Font.Size));
    Chart1.Title.Font:=ChartTitleEdit.Font;
    ChartForm.Chart1.Title.Font:=ChartTitleEdit.Font;
  end;
end;

procedure TMainForm.FD1Apply(Sender: TObject; Wnd: HWND);
begin
  Memo.Font:=FD1.Font;
  if fsBold in FD1.Font.Style then
    BoldBtn1.Down:=True
  else
    BoldBtn1.Down:=False;
  if fsItalic in FD1.Font.Style then
    ItalicBtn1.Down:=True
  else
    ItalicBtn1.Down:=False;
  if fsUnderLine in FD1.Font.Style then
    UnderBtn1.Down:=True
  else
    UnderBtn1.Down:=False;
  ColorBtn1.SelectedColor:=FD1.Font.Color;
  SizeCombo.ItemIndex:=SizeCombo.Items.IndexOf(IntToStr(FD1.Font.Size));
end;

procedure TMainForm.BoldBtn2Click(Sender: TObject);
begin
  if BoldBtn2.Down then
    ChartFooterEdit.Font.Style:=ChartFooterEdit.Font.Style+[fsBold]
  else
    ChartFooterEdit.Font.Style:=ChartFooterEdit.Font.Style-[fsBold];
  Chart1.Foot.Font:=ChartFooterEdit.Font;
  Chartform.Chart1.Foot.Font:=ChartFooterEdit.Font;
end;

procedure TMainForm.ItalicBtn2Click(Sender: TObject);
begin
  if Italicbtn2.Down then
    ChartFooterEdit.Font.Style:=ChartFooterEdit.Font.Style+[fsItalic]
  else
    ChartFooterEdit.Font.Style:=ChartFooterEdit.Font.Style-[fsItalic];
  Chart1.Foot.Font:=ChartFooterEdit.Font;
  ChartForm.Chart1.Foot.Font:=ChartFooterEdit.Font;
end;

procedure TMainForm.UnderBtn2Click(Sender: TObject);
begin
  if UnderBtn2.Down then
    ChartFooterEdit.Font.Style:=ChartFooterEdit.Font.Style+[fsUnderLine]
  else
    ChartFooterEdit.Font.Style:=ChartFooterEdit.Font.Style-[fsUnderLine];
  Chart1.Foot.Font:=ChartFooterEdit.Font;
  Chartform.Chart1.Foot.Font:=ChartFooterEdit.Font;
end;

procedure TMainForm.LeftAl2Click(Sender: TObject);
begin
  if LeftAl2.Down then
  begin
    ChartFooterEdit.Alignment:=taLeftJustify;
    Chart1.Foot.Alignment:=taLeftJustify;
    ChartForm.Chart1.Foot.Alignment:=taLeftJustify;
  end;
end;

procedure TMainForm.CenterAl2Click(Sender: TObject);
begin
  if CenterAl2.Down then
  begin
    ChartFooterEdit.Alignment:=taCenter;
    Chart1.Foot.Alignment:=taCenter;
    ChartForm.Chart1.Foot.Alignment:=taCenter;
  end;
end;

procedure TMainForm.RightAl2Click(Sender: TObject);
begin
  if RightAl2.Down then
  begin
    ChartFooterEdit.Alignment:=taRightJustify;
    Chart1.Foot.Alignment:=taRightJustify;
    Chartform.Chart1.Foot.Alignment:=taRightJustify;
  end;  
end;

procedure TMainForm.SizeCombo2Change(Sender: TObject);
begin
  ChartFooterEdit.Font.Size:=StrToInt(SizeCombo2.Text);
  Chart1.Foot.Font:=ChartFooterEdit.Font;
  ChartForm.Chart1.Foot.Font:=ChartFooterEdit.Font;
end;

procedure TMainForm.ColorBtn2ColorSelected(Sender: TObject;
  AColor: TColor);
begin
  ChartFooterEdit.Font.Color:=AColor;
  Chart1.Foot.Font:=ChartFooterEdit.Font;
  ChartForm.Chart1.Foot.Font:=ChartFooterEdit.Font;
end;

procedure TMainForm.SpeedButton10Click(Sender: TObject);
begin
  Memo:=ChartFooterEdit;
  FD1.Font:=ChartFooterEdit.Font;
  if FD1.Execute then
  begin
    ChartFooterEdit.Font:=FD1.Font;
    if fsBold in FD1.Font.Style then
      BoldBtn2.Down:=True
    else
      BoldBtn2.Down:=False;
    if fsItalic in FD1.Font.Style then
      ItalicBtn2.Down:=True
    else
      ItalicBtn2.Down:=False;
    if fsUnderLine in FD1.Font.Style then
      UnderBtn2.Down:=True
    else
      UnderBtn2.Down:=False;
    ColorBtn2.SelectedColor:=FD1.Font.Color;
    SizeCombo2.ItemIndex:=SizeCombo2.Items.IndexOf(IntToStr(FD1.Font.Size));
    Chart1.Foot.Font:=ChartFooterEdit.Font;
    ChartForm.Chart1.Foot.Font:=ChartFooterEdit.Font;
  end;
end;

procedure TMainForm.BoldBtn3Click(Sender: TObject);
begin
  if BoldBtn3.Down then
    AxisTitleEdit.Font.Style:=AxisTitleEdit.Font.Style+[fsBold]
  else
    AxisTitleEdit.Font.Style:=AxisTitleEdit.Font.Style-[fsBold];
  Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
  ChartForm.Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
end;

procedure TMainForm.ItalicBtn3Click(Sender: TObject);
begin
  if ItalicBtn3.Down then
    AxisTitleEdit.Font.Style:=AxisTitleEdit.Font.Style+[fsItalic]
  else
    AxisTitleEdit.Font.Style:=AxisTitleEdit.Font.Style-[fsItalic];
  Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
  ChartForm.Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
end;

procedure TMainForm.UnderBtn3Click(Sender: TObject);
begin
  if UnderBtn3.Down then
    AxisTitleEdit.Font.Style:=AxisTitleEdit.Font.Style+[fsUnderLine]
  else
    AxisTitleEdit.Font.Style:=AxisTitleEdit.Font.Style-[fsUnderLine];
  Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
  ChartForm.Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
end;

procedure TMainForm.SizeCombo3Change(Sender: TObject);
begin
  AxisTitleEdit.Font.Size:=StrToInt(SizeCombo3.Text);
  Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
  Chartform.Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
end;

procedure TMainForm.ColorBtn3ColorSelected(Sender: TObject;
  AColor: TColor);
begin
  AxisTitleEdit.Font.Color:=AColor;
  Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
  ChartForm.Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
end;

procedure TMainForm.SpeedButton11Click(Sender: TObject);
begin
  Memo:=AxisTitleEdit;
  FD1.Font:=AxisTitleEdit.Font;
  if FD1.Execute then
  begin
    AxisTitleEdit.Font:=FD1.Font;
    if fsBold in FD1.Font.Style then
      BoldBtn3.Down:=True
    else
      BoldBtn3.Down:=False;
    if fsItalic in FD1.Font.Style then
      ItalicBtn3.Down:=True
    else
      ItalicBtn3.Down:=False;
    if fsUnderLine in FD1.Font.Style then
      UnderBtn3.Down:=True
    else
      UnderBtn3.Down:=False;
    ColorBtn3.SelectedColor:=FD1.Font.Color;
    SizeCombo3.ItemIndex:=SizeCombo3.Items.IndexOf(IntToStr(FD1.Font.Size));
    Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
    ChartForm.Chart1.LeftAxis.Title.Font:=AxisTitleEdit.Font;
  end;
end;

procedure TMainForm.ChartBtnClick(Sender: TObject);
begin
  if ExplorerPopup1.Visible then
  begin
    ChartBtn.Down:=False;
    Exit;
  end;  
  if not(ChartBtn.Down) then
  begin
    Chart1.Left:=7;
    Chart1.Width:=595;
    Chart1.Height:=357;
    //ChartPageControl1.Visible:=False;
    ChartPageControl1.Parent:=ExplorerPopup1;
    ChartPageControl1.Top:=2;
    ChartPageControl1.Left:=4;
  end
  else
  begin
    Chart1.Left:=312;
    Chart1.Width:=291;
    Chart1.Height:=349;
    //ChartPageControl1.Visible:=True;
    ChartPageControl1.Parent:=ChartPanel;
    ChartPageControl1.Top:=44;
    ChartPageControl1.Left:=9;
  end;
end;

procedure TMainForm.DetachBtnClick(Sender: TObject);
begin
  if ChartForm.Visible then
  begin
    ChartForm.Visible:=False;
    DetachBtn.Glyph.Assign(Detachbmp.Picture.Bitmap);
    DetachBtn.Hint:='Detach Chart';
  end
  else
  begin
    ChartForm.Position:=poDesktopCenter;
    ChartForm.Visible:=True;
    DetachBtn.Glyph.Assign(Atachbmp.Picture.Bitmap);
    DetachBtn.Hint:='Attach Chart';
    ExplorerPopup1.BringToFront;    
  end;
end;

procedure TMainForm.xpCheckBox7Click(Sender: TObject);
begin
  Chart1.Legend.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Legend.Visible:=(Sender as TxpCheckBox).Checked;
  LegendOptions.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.LegendPosChange(Sender: TObject);
begin
  case LegendPos.ItemIndex of
    0:begin
        Chart1.Legend.Alignment:=laLeft;
        ChartForm.Chart1.Legend.Alignment:=laLeft;
      end;
    1:begin
        Chart1.Legend.Alignment:=laRight;
        ChartForm.Chart1.Legend.Alignment:=laRight;
      end;
    2:begin
        Chart1.Legend.Alignment:=laTop;
        Chartform.Chart1.Legend.Alignment:=laTop;
      end;
    3:begin
        Chart1.Legend.Alignment:=laBottom;
        ChartForm.Chart1.Legend.Alignment:=laBottom;
      end;
  end;
  if LegendPos.ItemIndex<=1 then
  begin
    HorizMarg.Enabled:=True;
    VertMarg.Enabled:=False;
  end
  else
  begin
    HorizMarg.Enabled:=False;
    VertMarg.Enabled:=True;
  end;
end;

procedure TMainForm.ColorSelector3ChangeColor(Sender: TObject);
begin
  Chart1.Legend.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Legend.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.SpinEdit34Change(Sender: TObject);
begin
  Chart1.Legend.ColorWidth:=LegendWidth.Value;
  ChartForm.Chart1.Legend.ColorWidth:=LegendWidth.Value;
end;

procedure TMainForm.xpButton2Click(Sender: TObject);
begin
  FD2.Font:=Chart1.Legend.Font;
  if FD2.Execute then
  begin
    Chart1.Legend.Font:=FD2.Font;
    ChartForm.Chart1.Legend.Font:=FD2.Font;
    LegendFont.Font:=FD2.Font;
  end;
end;

procedure TMainForm.LegendStyleChange(Sender: TObject);
begin
  case LegendStyle.ItemIndex of
    0:begin
        Chart1.Legend.TextStyle:=ltsLeftValue;
        Chartform.Chart1.Legend.TextStyle:=ltsLeftValue;
      end;
    1:begin
        Chart1.Legend.TextStyle:=ltsLeftPercent;
        Chartform.Chart1.Legend.TextStyle:=ltsLeftPercent;
      end;
  end;
  LegendStyle.Hint:=LegendStyle.Items.Strings[LegendStyle.ItemIndex];
end;

procedure TMainForm.xpCheckBox8Click(Sender: TObject);
begin
  Chart1.Legend.Inverted:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Legend.Inverted:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox9Click(Sender: TObject);
begin
  Chart1.Legend.ResizeChart:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Legend.ResizeChart:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox10Click(Sender: TObject);
begin
  Chart1.Legend.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  Chartform.Chart1.Legend.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  lcolor.Enabled:=(Sender as TxpCheckBox).Checked;
  iwidth2.Enabled:=(Sender as TxpCheckBox).Checked;
  istyle2.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.lwidth2Change(Sender: TObject);
begin
  Chart1.Legend.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  Chartform.Chart1.Legend.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
end;

procedure TMainForm.lstyle2Change(Sender: TObject);
begin
  Chart1.Legend.Frame.Style:=(Sender as TPenStyleCombo).Selection;
  ChartForm.Chart1.Legend.Frame.Style:=(Sender as TPenStyleCombo).Selection;
end;

procedure TMainForm.HorizMargChange(Sender: TObject);
begin
  Chart1.Legend.HorizMargin:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.Legend.HorizMargin:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.VertMargChange(Sender: TObject);
begin
  Chart1.Legend.VertMargin:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.Legend.VertMargin:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.xpCheckBox11Click(Sender: TObject);
begin
  if (Sender as TxpCheckBox).Checked then
  begin
    Chart1.Legend.ShadowSize:=ShadowSize.Value;
    ChartForm.Chart1.Legend.ShadowSize:=ShadowSize.Value;
  end
  else
  begin
    Chart1.Legend.ShadowSize:=0;
    ChartForm.Chart1.Legend.ShadowSize:=0;
  end;
  ShadowColor.Enabled:=(Sender as TxpCheckBox).Checked;
  ShadowSize.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.View3DCheckClick(Sender: TObject);
begin
  Chart1.View3DWalls:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.View3DWalls:=(Sender as TxpCheckBox).Checked;
  LeftWall.Enabled:=(Sender as TxpCheckBox).Checked;
  BackWall.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.SpinEdit5Change(Sender: TObject);
begin
  Chart1.LeftWall.Size:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftWall.Size:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.SpinEdit6Change(Sender: TObject);
begin
  Chart1.BackWall.Size:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.BackWall.Size:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.xpCheckBox13Click(Sender: TObject);
begin
  Chart1.LeftWall.Pen.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.LeftWall.Pen.Visible:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox14Click(Sender: TObject);
begin
  Chart1.BackWall.Pen.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.BackWall.Pen.Visible:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.penwidthcombo3Change(Sender: TObject);
begin
  Chart1.LeftWall.Pen.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.LeftWall.Pen.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    penstylecombo2.Enabled:=False
  else
    penstylecombo2.Enabled:=True;
  Label48.Enabled:=penstylecombo2.Enabled;
end;

procedure TMainForm.penwidthcombo4Change(Sender: TObject);
begin
  Chart1.BackWall.Pen.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.BackWall.Pen.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    penstylecombo3.Enabled:=False
  else
    penstylecombo3.Enabled:=True;
  Label53.Enabled:=penstylecombo3.Enabled;
end;

procedure TMainForm.penstylecombo2Change(Sender: TObject);
begin
  Chart1.LeftWall.Pen.Style:=(Sender as tpenstylecombo).Selection;
  ChartForm.Chart1.LeftWall.Pen.Style:=(Sender as tpenstylecombo).Selection;
end;

procedure TMainForm.penstylecombo3Change(Sender: TObject);
begin
  Chart1.BackWall.Pen.Style:=(Sender as tpenstylecombo).Selection;
  ChartForm.Chart1.BackWall.Pen.Style:=(Sender as tpenstylecombo).Selection;
end;

procedure TMainForm.UseColorsClick(Sender: TObject);
begin
  if (Sender as TRadioButton).Checked then
  begin
    Chart1.Monochrome:=False;
    ChartForm.Chart1.Monochrome:=False;
    ColorOptions.Enabled:=True;
  end;
end;

procedure TMainForm.RadioButton1Click(Sender: TObject);
begin
  if (Sender as TRadioButton).Checked then
  begin
    Chart1.Monochrome:=True;
    ChartForm.Chart1.Monochrome:=True;
    ColorOptions.Enabled:=False;
  end;
end;

procedure TMainForm.xpCheckBox6Click(Sender: TObject);
begin
  Chart1.View3D:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.View3D:=(Sender as TxpCheckBox).Checked;
  View3D.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.ZoomTrackChange(Sender: TObject);
begin
  Chart1.View3DOptions.Zoom:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.View3DOptions.Zoom:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.Chart3DTrackChange(Sender: TObject);
begin
  Chart1.Chart3DPercent:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.Chart3DPercent:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.XRChange(Sender: TObject);
begin
  Chart1.View3DOptions.Elevation:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.View3DOptions.Elevation:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.YRChange(Sender: TObject);
begin
  Chart1.View3DOptions.Rotation:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.View3DOptions.Rotation:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.ZRChange(Sender: TObject);
begin
  Chart1.View3DOptions.Tilt:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.View3DOptions.Tilt:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.PerspectiveTrackChange(Sender: TObject);
begin
  Chart1.View3DOptions.Perspective:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.View3DOptions.Perspective:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.NormalViewClick(Sender: TObject);
begin
  Chart1.View3DOptions.Orthogonal:=(Sender as TRadioButton).Checked;
  ChartForm.Chart1.View3DOptions.Orthogonal:=(Sender as TRadioButton).Checked;
  XR.Enabled:=not((Sender as TRadioButton).Checked);
  YR.Enabled:=not((Sender as TRadioButton).Checked);
  ZR.Enabled:=not((Sender as TRadioButton).Checked);
  PerspectiveTrack.Enabled:=not((Sender as TRadioButton).Checked);
end;

procedure TMainForm.CustomizedViewClick(Sender: TObject);
begin
  Chart1.View3DOptions.Orthogonal:=not((Sender as TRadioButton).Checked);
  ChartForm.Chart1.View3DOptions.Orthogonal:=not((Sender as TRadioButton).Checked);
  XR.Enabled:=(Sender as TRadioButton).Checked;
  YR.Enabled:=(Sender as TRadioButton).Checked;
  ZR.Enabled:=(Sender as TRadioButton).Checked;
  PerspectiveTrack.Enabled:=(Sender as TRadioButton).Checked;
end;

procedure TMainForm.xpCheckBox15Click(Sender: TObject);
begin
  Chart1.Title.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Title.Visible:=(Sender as TxpCheckBox).Checked;
  ResizeT.Enabled:=(Sender as TxpCheckBox).Checked;
  ColorT.Enabled:=(Sender as TxpCheckBox).Checked;
  TitleText.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox17Click(Sender: TObject);
begin
  Chart1.Foot.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Foot.Visible:=(Sender as TxpCheckBox).Checked;
  ResizeF.Enabled:=(Sender as TxpCheckBox).Checked;
  ColorF.Enabled:=(Sender as TxpCheckBox).Checked;
  FooterText.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.ResizeTClick(Sender: TObject);
begin
  Chart1.Title.AdjustFrame:=not((Sender as TxpCheckBox).Checked);
  ChartForm.Chart1.Title.AdjustFrame:=not((Sender as TxpCheckBox).Checked);
end;

procedure TMainForm.ResizeFClick(Sender: TObject);
begin
  Chart1.Foot.AdjustFrame:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Foot.AdjustFrame:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox20Click(Sender: TObject);
begin
  Chart1.Foot.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Foot.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  bcolorf.Enabled:=(Sender as TxpCheckBox).Checked;
  bwidthf.Enabled:=(Sender as TxpCheckBox).Checked;
  bstyle.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox16Click(Sender: TObject);
begin
  Chart1.Title.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Title.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  bcolor.Enabled:=(Sender as TxpCheckBox).Checked;
  bwidth.Enabled:=(Sender as TxpCheckBox).Checked;
  bstyle.Enabled:=(Sender as TxpCheckBox).Checked;
  TDefBorder.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.bcolorfClick(Sender: TObject);
begin
  Chart1.Foot.Frame.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Foot.Frame.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.bwidthChange(Sender: TObject);
begin
  Chart1.Title.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.Title.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    bstyle.Enabled:=False
  else
    bstyle.Enabled:=True;
  Label57.Enabled:=bstyle.Enabled;  
end;

procedure TMainForm.bwidthfChange(Sender: TObject);
begin
  Chart1.Foot.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.Foot.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    bstylef.Enabled:=False
  else
    bstylef.Enabled:=True;
  Label61.Enabled:=bstylef.Enabled;
end;

procedure TMainForm.bstylefChange(Sender: TObject);
begin
  Chart1.Foot.Frame.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.Foot.Frame.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.bstyleChange(Sender: TObject);
begin
  Chart1.Title.Frame.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.Title.Frame.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.ChartTitleEditChange(Sender: TObject);
begin
  Chart1.Title.Text:=ChartTitleEdit.Lines;
  ChartForm.Chart1.Title.Text:=ChartTitleEdit.Lines;
end;

procedure TMainForm.ChartFooterEditChange(Sender: TObject);
begin
  Chart1.Foot.Text:=ChartFooterEdit.Lines;
  ChartForm.Chart1.Foot.Text:=ChartFooterEdit.Lines;
end;

procedure TMainForm.RadioButton8Click(Sender: TObject);
begin
  Chart1.LeftAxis.Automatic:=(Sender as TRadioButton).Checked;
  Automatic.Enabled:=not((Sender as TRadioButton).Checked);
end;

procedure TMainForm.RadioButton9Click(Sender: TObject);
begin
  Chart1.LeftAxis.Automatic:=not((Sender as TRadioButton).Checked);
  ChartForm.Chart1.LeftAxis.Automatic:=not((Sender as TRadioButton).Checked);
  Automatic.Enabled:=(Sender as TRadioButton).Checked;
  Minspin.Value:=Round(Chart1.LeftAxis.Minimum);
  Maxspin.Value:=Round(Chart1.LeftAxis.Maximum);
  Increment.Value:=Round(Chart1.LeftAxis.Increment);
end;

procedure TMainForm.SpinEdit10Change(Sender: TObject);
begin
  Chart1.LeftAxis.PositionPercent:=(Sender as TSpinEdit).Value;
  cHARTfORM.Chart1.LeftAxis.PositionPercent:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.SpinEdit11Change(Sender: TObject);
begin
  Chart1.LeftAxis.EndPosition:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.EndPosition:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.SpinEdit12Change(Sender: TObject);
begin
  Chart1.LeftAxis.StartPosition:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.StartPosition:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.xpCheckBox21Click(Sender: TObject);
begin
  Chart1.LeftAxis.Inverted:=(Sender as TxpcheckBox).Checked;
  ChartForm.Chart1.LeftAxis.Inverted:=(Sender as TxpcheckBox).Checked;
end;

procedure TMainForm.xpCheckBox22Click(Sender: TObject);
begin
  Chart1.LeftAxis.Visible:=(Sender as TxpcheckBox).Checked;
  ChartForm.Chart1.LeftAxis.Visible:=(Sender as TxpcheckBox).Checked;
  LeftAxis.Enabled:=(Sender as TxpcheckBox).Checked;
end;

procedure TMainForm.xpCheckBox23Click(Sender: TObject);
begin
  Chart1.LeftAxis.Labels:=(Sender as TxpcheckBox).Checked;
  ChartForm.Chart1.LeftAxis.Labels:=(Sender as TxpcheckBox).Checked;
  Labelcustom.Enabled:=(Sender as TxpcheckBox).Checked;
  mmlabel.Enabled:=(Sender as TxpcheckBox).Checked;
  mmtrack.Enabled:=(Sender as TxpcheckBox).Checked;
end;

procedure TMainForm.labelcustomClick(Sender: TObject);
begin
  FD2.Font:=Chart1.LeftAxis.LabelsFont;
  if FD2.Execute then
  begin
    Chart1.LeftAxis.LabelsFont:=FD2.Font;
    ChartForm.Chart1.LeftAxis.LabelsFont:=FD2.Font;
    Labelsfont.Font:=FD2.Font;
  end;  
end;

procedure TMainForm.mmtrackChange(Sender: TObject);
begin
  Chart1.LeftAxis.LabelsAngle:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.LeftAxis.LabelsAngle:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.mmlabelClick(Sender: TObject);
begin
  Chart1.LeftAxis.LabelsOnAxis:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.LeftAxis.LabelsOnAxis:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox24Click(Sender: TObject);
var i:Integer;
begin
  if ((Sender as TxpCheckBox).Checked) and (AxisTitleEdit.Lines.Count>0) then
  begin
    Chart1.LeftAxis.Title.Caption:=AxistitleEdit.Lines.Strings[0];
    ChartForm.Chart1.LeftAxis.Title.Caption:=AxistitleEdit.Lines.Strings[0];
    for i:=1 to (AxisTitleEdit.Lines.Count-1) do
    begin
      Chart1.LeftAxis.Title.Caption:=Chart1.LeftAxis.Title.Caption+Chr(13)+AxisTitleEdit.Lines.Strings[i];
      ChartForm.Chart1.LeftAxis.Title.Caption:=Chart1.LeftAxis.Title.Caption+Chr(13)+AxisTitleEdit.Lines.Strings[i];
    end;
  end
  else
  begin
    Chart1.LeftAxis.Title.Caption:='';
    ChartForm.Chart1.LeftAxis.Title.Caption:='';
  end;
  AxisTitle.Enabled:=(Sender as TxpCheckBox).Checked;
  TitleSize.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.AxisTitleRChange(Sender: TObject);
begin
  Chart1.LeftAxis.Title.Angle:=(Sender as TExTrackBar).Position;
  ChartForm.Chart1.LeftAxis.Title.Angle:=(Sender as TExTrackBar).Position;
end;

procedure TMainForm.disSpinChange(Sender: TObject);
begin
  Chart1.LeftAxis.TitleSize:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.TitleSize:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.xpCheckBox26Click(Sender: TObject);
begin
  Chart1.LeftAxis.Ticks.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.LeftAxis.Ticks.Visible:=(Sender as TxpCheckBox).Checked;
  MLength.Enabled:=(Sender as TxpCheckBox).Checked;
  MColor.Enabled:=(Sender as TxpCheckBox).Checked;
  MWidth.Enabled:=(Sender as TxpCheckBox).Checked;
  MStyle.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox28Click(Sender: TObject);
begin
  Chart1.LeftAxis.TicksInner.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.LeftAxis.TicksInner.Visible:=(Sender as TxpCheckBox).Checked;
  ILength.Enabled:=(Sender as TxpCheckBox).Checked;
  IColor.Enabled:=(Sender as TxpCheckBox).Checked;
  IWidth.Enabled:=(Sender as TxpCheckBox).Checked;
  IStyle.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox27Click(Sender: TObject);
begin
  Chart1.LeftAxis.MinorTicks.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.LeftAxis.MinorTicks.Visible:=(Sender as TxpCheckBox).Checked;
  miCount.Enabled:=(Sender as TxpCheckBox).Checked;
  miLength.Enabled:=(Sender as TxpCheckBox).Checked;
  miColor.Enabled:=(Sender as TxpCheckBox).Checked;
  miWidth.Enabled:=(Sender as TxpCheckBox).Checked;
  miStyle.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.miCountChange(Sender: TObject);
begin
  Chart1.LeftAxis.MinorTickCount:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.MinorTickCount:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.miLengthChange(Sender: TObject);
begin
  Chart1.LeftAxis.MinorTickLength:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.MinorTickLength:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.miWidthChange(Sender: TObject);
begin
  Chart1.LeftAxis.MinorTicks.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.LeftAxis.MinorTicks.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    miStyle.Enabled:=False
  else
    miStyle.Enabled:=True;
  Label83.Enabled:=miStyle.Enabled;
end;

procedure TMainForm.miStyleChange(Sender: TObject);
begin
  Chart1.LeftAxis.MinorTicks.Style:=(Sender as Tpenstylecombo).Selection;
  chartForm.Chart1.LeftAxis.MinorTicks.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.MLengthChange(Sender: TObject);
begin
  Chart1.LeftAxis.TickLength:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.TickLength:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.ILengthChange(Sender: TObject);
begin
  Chart1.LeftAxis.TickInnerLength:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.TickInnerLength:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.IWidthChange(Sender: TObject);
begin
  Chart1.LeftAxis.TicksInner.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.LeftAxis.TicksInner.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    IStyle.Enabled:=False
  else
    IStyle.Enabled:=True;
  Label87.Enabled:=IStyle.Enabled;
end;

procedure TMainForm.MWidthChange(Sender: TObject);
begin
  Chart1.LeftAxis.Ticks.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.LeftAxis.Ticks.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    MStyle.Enabled:=False
  else
    MStyle.Enabled:=True;
  Label82.Enabled:=MStyle.Enabled;
end;

procedure TMainForm.MStyleChange(Sender: TObject);
begin
  Chart1.LeftAxis.Ticks.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.LeftAxis.Ticks.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.IStyleChange(Sender: TObject);
begin
  Chart1.LeftAxis.TicksInner.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.LeftAxis.TicksInner.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.IColorChangeColor(Sender: TObject);
begin
  Chart1.LeftAxis.TicksInner.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftAxis.TicksInner.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.MColorChangeColor(Sender: TObject);
begin
  Chart1.LeftAxis.Ticks.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftAxis.Ticks.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.miColorChangeColor(Sender: TObject);
begin
  Chart1.LeftAxis.MinorTicks.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftAxis.MinorTicks.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.MLengthKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.ILengthKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.miCountKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.miLengthKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.disSpinKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.minspinKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.SpinEdit5KeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.HorizMargKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.LegendWidthKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.LineBtnClick(Sender: TObject);
var i:Integer;
begin
  for i:=1 to 7 do
  begin
    Chart1.Series[i-1].Active:=False;
    ChartForm.Chart1.Series[i-1].Active:=False;
  end;
  WallSheet.Enabled:=True;
  AxisSheet.Enabled:=True;
  case (Sender as TSpeedButton).Tag of
    0:begin
        Series1.Active:=True;
        ChartForm.Series1.Active:=True;
      end;
    1:begin
        Series2.Active:=True;
        ChartForm.Series2.Active:=True;
      end;
    2:begin
        Series3.Active:=True;
        ChartForm.Series3.Active:=True;
      end;
    3:begin
        Series4.Active:=True;
        ChartForm.Series4.Active:=True;
      end;
    4:begin
        Series5.Active:=True;
        ChartForm.Series5.Active:=True;
      end;
    5:begin
        Series6.Active:=True;
        ChartForm.Series6.Active:=True;
        WallSheet.Enabled:=False;
        AxisSheet.Enabled:=False;
      end;
    6:begin
        Series7.Active:=True;
        ChartForm.Series7.Active:=True;
      end;
  end;
  if NormalView.Checked then
      Chart1.View3DOptions.Orthogonal:=True;
end;

procedure TMainForm.bcClick(Sender: TObject);
begin
  if (Sender as TxpCheckBox).Checked then
  begin
    Chart1.BackColor:=BackColor.Color;
    ChartForm.Chart1.BackColor:=BackColor.Color;
  end
  else
  begin
    Chart1.BackColor:=clTeeColor;
    ChartForm.Chart1.BackColor:=clTeeColor;
  end;
  BackColor.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.BackColorChangeColor(Sender: TObject);
begin
  Chart1.BackColor:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.BackColor:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.bfClick(Sender: TObject);
begin
  if (Sender as TxpCheckBox).Checked then
  begin
    Chart1.Frame.Color:=FrameColor.Color;
    ChartForm.Chart1.Frame.Color:=FrameColor.Color;
  end
  else
  begin
    Chart1.Frame.Color:=clTeeColor;
    ChartForm.Chart1.Frame.Color:=clTeeColor;
  end;
  Chart1.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Frame.Visible:=(Sender as TxpCheckBox).Checked;
  FrameColor.Enabled:=(Sender as TxpCheckBox).Checked;    
end;

procedure TMainForm.FrameColorChangeColor(Sender: TObject);
begin
  Chart1.Frame.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Frame.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.cgClick(Sender: TObject);
begin
  Chart1.Gradient.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.Gradient.Visible:=(Sender as TxpCheckBox).Checked;
  StartColor.Enabled:=(Sender as TxpCheckBox).Checked;
  EndColor.Enabled:=(Sender as TxpCheckBox).Checked;
  GLabel.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.StartColorChange(Sender: TObject);
begin
  Chart1.Gradient.StartColor:=(Sender as TColorBox).Selected;
  ChartForm.Chart1.Gradient.StartColor:=(Sender as TColorBox).Selected;
end;

procedure TMainForm.EndColorChange(Sender: TObject);
begin
  Chart1.Gradient.EndColor:=(Sender as TColorBox).Selected;
  ChartForm.Chart1.Gradient.EndColor:=(Sender as TColorBox).Selected;
end;

procedure TMainForm.UseBackImageClick(Sender: TObject);
begin
  if not((Sender as TxpCheckBox).Checked) then
  begin
    Chart1.BackImage.Assign(nil);
    ChartForm.Chart1.BackImage.Assign(nil);
    BackImage.Picture.Assign(nil);
  end;
  BrowseImage.Enabled:=(Sender as TxpCheckBox).Checked;
  Style.Enabled:=(Sender as TxpCheckBox).Checked;
  PutInside.Enabled:=(Sender as TxpCheckBox).Checked;
  BrowseImage.Repaint;
  PutInside.Repaint;
end;

procedure TMainForm.LineBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Linear Shaped Chart';
end;

procedure TMainForm.ColorTChangeColor(Sender: TObject);
begin
  Chart1.Title.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Title.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.bColorChangeColor(Sender: TObject);
begin
  Chart1.Title.Frame.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Title.Frame.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.StaticText25Click(Sender: TObject);
begin
  ColorT.Color:=clDefault;
  Chart1.Title.Color:=clDefault;
  ChartForm.Chart1.Title.Color:=clDefault;
end;

procedure TMainForm.TDefBorderClick(Sender: TObject);
begin
  Chart1.Title.Frame.Color:=clDefault;
  ChartForm.Chart1.Title.Frame.Color:=clDefault;
  Chart1.Title.Frame.Width:=1;
  ChartForm.Chart1.Title.Frame.Width:=1;
  Chart1.Title.Frame.Style:=psSolid;
  ChartForm.Chart1.Title.Frame.Style:=psSolid;
  bColor.Color:=clDefault;
  bWidth.ItemIndex:=0;
  bstyle.ItemIndex:=0;
end;

procedure TMainForm.StaticText26Click(Sender: TObject);
begin
  ColorF.Color:=clBtnFace;
  Chart1.Foot.Color:=clBtnFace;
  ChartForm.Chart1.Foot.Color:=clBtnFace;
end;

procedure TMainForm.StaticText29Click(Sender: TObject);
begin
  Chart1.Foot.Frame.Color:=clDefault;
  ChartForm.Chart1.Foot.Frame.Color:=clDefault;
  Chart1.Foot.Frame.Width:=1;
  ChartForm.Chart1.Foot.Frame.Width:=1;
  Chart1.Foot.Frame.Style:=psSolid;
  ChartForm.Chart1.Foot.Frame.Style:=psSolid;
  bColorf.Color:=clDefault;
  bWidthf.ItemIndex:=0;
  bstylef.ItemIndex:=0;
end;

procedure TMainForm.FRUnderClick(Sender: TObject);
begin
  if FRUnder.Down then
    FreqReportEdit.Font.Style:=FreqReportEdit.Font.Style+[fsUnderline]
  else
    FreqReportEdit.Font.Style:=FreqReportEdit.Font.Style-[fsUnderline];
end;

procedure TMainForm.lcolorChangeColor(Sender: TObject);
begin
  Chart1.Legend.Frame.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Legend.Frame.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.ShadowColorChangeColor(Sender: TObject);
begin
  Chart1.Legend.ShadowColor:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Legend.ShadowColor:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.ShadowSizeChange(Sender: TObject);
begin
  Chart1.Legend.ShadowSize:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.Legend.ShadowSize:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.ColorSelector6ChangeColor(Sender: TObject);
begin
  Chart1.LeftWall.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftWall.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.ColorSelector9ChangeColor(Sender: TObject);
begin
  Chart1.BackWall.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.BackWall.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.ColorSelector8ChangeColor(Sender: TObject);
begin
  Chart1.BackWall.Pen.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.BackWall.Pen.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.ColorSelector7ChangeColor(Sender: TObject);
begin
  Chart1.LeftWall.Pen.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftWall.Pen.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.StaticText30Click(Sender: TObject);
begin
  ColorSelector6.Color:=$0080FFFF;
  Chart1.LeftWall.Color:=$0080FFFF;
  ChartForm.Chart1.LeftWall.Color:=$0080FFFF;
end;

procedure TMainForm.StaticText31Click(Sender: TObject);
begin
  ColorSelector9.Color:=clBtnFace;
  Chart1.BackWall.Color:=clBtnFace;
  ChartForm.Chart1.BackWall.Color:=clBtnFace;
end;

procedure TMainForm.MaxSpinKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.minspinChange(Sender: TObject);
begin
  Chart1.LeftAxis.Minimum:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.Minimum:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.MaxSpinChange(Sender: TObject);
begin
  Chart1.LeftAxis.Maximum:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.Maximum:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.IncrementChange(Sender: TObject);
begin
  Chart1.LeftAxis.Increment:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.Increment:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.AxisTitleEditChange(Sender: TObject);
var i:Integer; S:String;
begin
  if AxisTitleEdit.Lines.Count>0 then
    S:=AxisTitleEdit.Lines.Strings[0]
  else
    S:='';
  if AxisTitleEdit.Lines.Count>1 then
    for i:= 1 to (AxisTitleEdit.Lines.Count-1) do
      S:=S+Chr(13)+AxisTitleEdit.Lines.Strings[i];
  Chart1.LeftAxis.Title.Caption:=S;
  ChartForm.Chart1.LeftAxis.Title.Caption:=S;
  Chart1.LeftAxis.Title.Font:=(Sender as TMemo).Font;
  ChartForm.Chart1.LeftAxis.Title.Font:=(Sender as TMemo).Font;
end;

procedure TMainForm.TitleSizeKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.TitleSizeChange(Sender: TObject);
begin
  Chart1.LeftAxis.TitleSize:=(Sender as TSpinEdit).Value;
  ChartForm.Chart1.LeftAxis.TitleSize:=(Sender as TSpinEdit).Value;
end;

procedure TMainForm.xpCheckBox29Click(Sender: TObject);
begin
  if (Sender as TxpCheckBox).Checked then
  begin
    Chart1.LeftAxis.Axis.Width:=StrToInt(axwidth.Items.Strings[axwidth.ItemIndex]);
    ChartForm.Chart1.LeftAxis.Axis.Width:=StrToInt(axwidth.Items.Strings[axwidth.ItemIndex]);
  end
  else
  begin
    Chart1.LeftAxis.Axis.Width:=0;
    ChartForm.Chart1.LeftAxis.Axis.Width:=0;
  end;
//  Chart1.LeftAxis.Visible:=(Sender as TxpCheckBox).Checked;
  axcolor.Enabled:=(Sender as TxpCheckBox).Checked;
  axwidth.Enabled:=(Sender as TxpCheckBox).Checked;
  axstyle.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.axcolorChangeColor(Sender: TObject);
begin
  Chart1.LeftAxis.Axis.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftAxis.Axis.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.gridcolorChangeColor(Sender: TObject);
begin
  Chart1.LeftAxis.Grid.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.LeftAxis.Grid.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.axwidthChange(Sender: TObject);
begin
  Chart1.LeftAxis.Axis.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.LeftAxis.Axis.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    axstyle.Enabled:=False
  else
    axstyle.Enabled:=True;
  Label90.Enabled:=axstyle.Enabled;
end;

procedure TMainForm.gridwidthChange(Sender: TObject);
begin
  Chart1.LeftAxis.Grid.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.LeftAxis.Grid.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    gridstyle.Enabled:=False
  else
    gridstyle.Enabled:=True;
  Label93.Enabled:=gridstyle.Enabled;
end;

procedure TMainForm.axstyleChange(Sender: TObject);
begin
  Chart1.LeftAxis.Axis.Style:=(Sender as tpenstylecombo).Selection;
  ChartForm.Chart1.LeftAxis.Axis.Style:=(Sender as tpenstylecombo).Selection;
end;

procedure TMainForm.gridstyleChange(Sender: TObject);
begin
  Chart1.LeftAxis.Grid.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.LeftAxis.Grid.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.xpCheckBox30Click(Sender: TObject);
begin
  Chart1.LeftAxis.Grid.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.LeftAxis.Grid.Visible:=(Sender as TxpCheckBox).Checked;
  gridcolor.Enabled:=(Sender as TxpCheckBox).Checked;
  gridwidth.Enabled:=(Sender as TxpCheckBox).Checked;
  gridstyle.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.xpCheckBox1Click(Sender: TObject);
begin
  Chart1.BottomAxis.Grid.Visible:=(Sender as TxpCheckBox).Checked;
  ChartForm.Chart1.BottomAxis.Grid.Visible:=(Sender as TxpCheckBox).Checked;
  gridcolorv.Enabled:=(Sender as TxpCheckBox).Checked;
  gridwidthv.Enabled:=(Sender as TxpCheckBox).Checked;
  gridstylev.Enabled:=(Sender as TxpCheckBox).Checked;
end;

procedure TMainForm.gridstylevChange(Sender: TObject);
begin
  Chart1.BottomAxis.Grid.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.BottomAxis.Grid.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.gridwidthvChange(Sender: TObject);
begin
  Chart1.BottomAxis.Grid.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.BottomAxis.Grid.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    gridstylev.Enabled:=False
  else
    gridstylev.Enabled:=True;
  Label79.Enabled:=gridstylev.Enabled;
end;

procedure TMainForm.gridcolorvChangeColor(Sender: TObject);
begin
  Chart1.BottomAxis.Grid.Color:=(Sender as TColorSelector).Color;
  Chartform.Chart1.BottomAxis.Grid.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.PrintBtnClick(Sender: TObject);
begin
  if PSD1.Execute then
    Chart1.Print;
end;

procedure TMainForm.CopyAsBitmap1Click(Sender: TObject);
begin
  Chart1.CopyToClipboardBitmap;
end;

procedure TMainForm.CopyAsMetafile1Click(Sender: TObject);
begin
  Chart1.CopyToClipboardMetafile(False);
end;

procedure TMainForm.UseColors2Click(Sender: TObject);
begin
  UseColors.Checked:=(Sender as TRadioButton).Checked;
end;

procedure TMainForm.RadioButton6Click(Sender: TObject);
begin
RadioButton1.Checked:=(Sender as TRadioButton).Checked;
end;

procedure TMainForm.ExplorerPopup1Open(Sender: TObject);
begin
  if ChartBtn.Down then
  begin
    ChartBtn.Down:=False;
    Chart1.Left:=7;
    Chart1.Width:=595;
    Chart1.Height:=357;
    //ChartPageControl1.Visible:=False;
    ChartPageControl1.Parent:=ExplorerPopup1;
    ChartPageControl1.Top:=2;
    ChartPageControl1.Left:=4;
  end;
end;

procedure TMainForm.RBoldClick(Sender: TObject);
begin
  if RBold.Down then
    ReportEdit.Font.Style:=ReportEdit.Font.Style+[fsBold]
  else
    ReportEdit.Font.Style:=ReportEdit.Font.Style-[fsBold];
end;

procedure TMainForm.RItalicClick(Sender: TObject);
begin
  if RItalic.Down then
    ReportEdit.Font.Style:=ReportEdit.Font.Style+[fsItalic]
  else
    ReportEdit.Font.Style:=ReportEdit.Font.Style-[fsItalic];
end;

procedure TMainForm.RUnderClick(Sender: TObject);
begin
  if RUnder.Down then
    ReportEdit.Font.Style:=ReportEdit.Font.Style+[fsUnderline]
  else
    ReportEdit.Font.Style:=ReportEdit.Font.Style-[fsUnderline];
end;

procedure TMainForm.RLeftClick(Sender: TObject);
begin
  if RLeft.Down then
    ReportEdit.Alignment:=taLeftJustify;
end;

procedure TMainForm.RCenterClick(Sender: TObject);
begin
  if RCenter.Down then
    ReportEdit.Alignment:=taCenter;
end;

procedure TMainForm.RRightClick(Sender: TObject);
begin
  if RRight.Down then
    ReportEdit.Alignment:=taRightJustify;
end;

procedure TMainForm.RSizeComboChange(Sender: TObject);
begin
  ReportEdit.Font.Size:=StrToInt(RSizeCombo.Text);
end;

procedure TMainForm.RColorBtnColorSelected(Sender: TObject;
  AColor: TColor);
begin
  ReportEdit.Font.Color:=AColor;
end;

procedure TMainForm.RBuildClick(Sender: TObject);
begin
  FD3.Font:=ReportEdit.Font;
  if FD3.Execute then
  begin
    ReportEdit.Font:=FD3.Font;
    if fsBold in FD3.Font.Style then
      RBold.Down:=True
    else
      RBold.Down:=False;
    if fsItalic in FD3.Font.Style then
      RItalic.Down:=True
    else
      RItalic.Down:=False;
    if fsUnderLine in FD3.Font.Style then
      RUnder.Down:=True
    else
      RUnder.Down:=False;
    RColorBtn.SelectedColor:=FD3.Font.Color;
    RSizeCombo.ItemIndex:=RSizeCombo.Items.IndexOf(IntToStr(FD3.Font.Size));
  end;
end;

procedure TMainForm.xpButton3Click(Sender: TObject);
begin
  ReportEdit.SelectAll;
  ReportEdit.CopyToClipboard;
  ReportEdit.SelLength:=0;
end;

procedure TMainForm.xpButton1Click(Sender: TObject);
begin
  if PSD1.Execute then
    ReportEdit.Print('Data Analyzing Report');
end;

procedure TMainForm.Load1Click(Sender: TObject);
begin
  if OpenGrid.Execute then
    LoadGridFromFile(OpenGrid.FileName,False);
end;

procedure TMainForm.Save1Click(Sender: TObject);
begin
  if SaveGrid.Execute then
    SaveGridToFile(SaveGrid.FileName);
end;

procedure TMainForm.ExplorerButton11DropDownClick(Sender: TObject);
begin
  FGrid:=SGrid4;
end;

procedure TMainForm.ExplorerButton12DropDownClick(Sender: TObject);
begin
  FGrid:=SGrid3;
end;

procedure TMainForm.ExplorerButton13DropDownClick(Sender: TObject);
begin
  FGrid:=SGrid2;
end;

procedure TMainForm.ExplorerButton14DropDownClick(Sender: TObject);
begin
  FGrid:=SGrid1;
end;

procedure TMainForm.ExplorerButton15DropDownClick(Sender: TObject);
begin
  FGrid:=SGrid5;
end;

procedure TMainForm.ChartSectionClick(Sender: TObject);

  function GetChartHistory(Index:Byte):TChart;
  begin
    case Index of
        1: Result:=HistoryChart1;
        2: Result:=HistoryChart2;
        3: Result:=HistoryChart3;
        4: Result:=HistoryChart4;
        5: Result:=HistoryChart5;
    end;
  end;

var i,j:Integer; B:Boolean;
    HChart,HFrom,HTo:TChart;
begin
  PrintToolbtn.Enabled:=False;
  CutToolbtn.Enabled:=False;
  CopyToolbtn.Enabled:=False;
  PasteToolbtn.Enabled:=False;
  FindToolbtn.Enabled:=False;

  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
  if CanCreateAChart then
  begin
    if NeedChartRefresh then
    begin
      ChartGrid.RowCount:=TableGrid.RowCount-1;
      ColorGrid.RowCount:=ChartGrid.RowCount;
      for i:=1 to (ChartGrid.RowCount-1) do
      begin
        ChartGrid.Cells[1,i]:=TableGrid.Cells[0,i];
        ChartGrid.Cells[2,i]:=TableGrid.Cells[1,i];
      end;
      if ChartHistoryCount<5 then
      begin
        i:=ChartHistoryCount+1;
        Inc(ChartHistoryCount);
      end
      else
      begin
        for i:=1 to 4 do
        begin
          HFrom:=GetChartHistory(i+1);
          HTo:=GetChartHistory(i);
          for j:=1 to HTo.Series[0].Count do
            HTo.Series[0].Delete(0);
          for j:=0 to (HFrom.Series[0].Count-1) do
            HTo.Series[0].Add(HFrom.Series[0].YValue[j]);
        end;
        i:=5;
      end;
      HChart:=GetChartHistory(i);
      for i:=1 to HChart.Series[0].Count do
        HChart.Series[0].Delete(0);
      SetLength(ChartColors,ChartGrid.RowCount-1);
      ResetChart(Chart1);
      ResetChart(ChartForm.Chart1);
      for i:=1 to (ChartGrid.RowCount-1) do
      begin
        Series1.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        Series2.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        Series3.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        Series4.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        Series5.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        Series6.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        Series7.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series1.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series2.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series3.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series4.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series5.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series6.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        ChartForm.Series7.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i]);
        HChart.Series[0].Add(StrToInt(ChartGrid.Cells[2,i]));
      end;
      for i:=0 to (Series6.Count-1) do
        ChartColors[i]:=Series6.ValueColor[i];
      NeedChartRefresh:=False;
    end;
    ChartPanel.Visible:=True;
    ChartPanel.BringToFront;
    IconChart.Down:=True;
    ToolChart.Down:=True;
    MChartSection.Checked:=True;
    ActiveSheet:=7;
  end
  else
  begin
    ShowMessage('There is no frequency table available to create a chart for it.');
    TableSection.Click;
    TableSection.Down:=True;
    ToolTable.Down:=True;
    IConTable.Down:=True;
    MTableSection.Checked:=True;
  end;
end;

procedure TMainForm.SGrid5MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i,j:Integer; R:TRect; S:String;
begin
  for i:=0 to (SGrid5.ColCount-1) do
  begin
    R:=SGrid5.CellRect(i,0);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S5IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=0 to (SGrid5.RowCount-2) do
  begin
    R:=Sgrid5.CellRect(0,i);
    if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
    begin
      S5IDEdit.Caption:='000';
      Exit;
    end;
  end;
  for i:=SGrid5.TopRow to (SGrid5.RowCount-2) do
    for j:=1 to (SGrid5.ColCount-1) do
    begin
      R:=SGrid5.CellRect(j,i);
      if (X>=R.Left) and (X<=R.Right) and (Y>=R.Top) and (Y<=R.Bottom) then
      begin
        S:=Sgrid5.Cells[0,i]+SGrid5.Cells[j,0];
        if S[1]='0' then
          S:=Copy(S,2,Length(S)-1);
        S5IDEdit.Caption:=S;
        Exit;
      end;
    end;
  S5IDEdit.Caption:='000';
end;

procedure TMainForm.S4FilterbtnClick(Sender: TObject);
begin
  if Sheet4.TypeIndex=1 then
    FilterGridValues(SGrid4,Copy(S4FilterLabel.Caption,6,Length(S4FilterLabel.Caption)-5),S4PSpin.Value,False)
  else if Sheet4.TypeIndex=2 then
    FilterGridValues(SGrid4,Copy(S4FilterLabel.Caption,6,Length(S4FilterLabel.Caption)-5),S4PSpin.Value,True);
  RebuildFreqTable:=True;
end;

procedure TMainForm.S4PSpinKeyPress(Sender: TObject; var Key: Char);
begin
  Key:=Chr(0);
end;

procedure TMainForm.S1FilterbtnClick(Sender: TObject);
begin
  if Sheet1.TypeIndex=1 then
    FilterGridValues(SGrid1,Copy(S1FilterLabel.Caption,6,Length(S1FilterLabel.Caption)-5),S1PSpin.Value,False)
  else if Sheet4.TypeIndex=2 then
    FilterGridValues(SGrid4,Copy(S4FilterLabel.Caption,6,Length(S4FilterLabel.Caption)-5),S4PSpin.Value,True);
  RebuildFreqTable:=True;
end;

procedure TMainForm.S3FilterbtnClick(Sender: TObject);
begin
  if Sheet3.TypeIndex=1 then
    FilterGridValues(SGrid3,Copy(S3FilterLabel.Caption,6,Length(S3FilterLabel.Caption)-5),S3PSpin.Value,False)
  else if Sheet3.TypeIndex=2 then
    FilterGridValues(SGrid3,Copy(S3FilterLabel.Caption,6,Length(S3FilterLabel.Caption)-5),S3PSpin.Value,True);
  RebuildFreqTable:=True;
end;

procedure TMainForm.S5FilterbtnClick(Sender: TObject);
begin
  if Sheet5.TypeIndex=1 then
    FilterGridValues(SGrid5,Copy(S5FilterLabel.Caption,6,Length(S5FilterLabel.Caption)-5),S5PSpin.Value,False)
  else if Sheet5.TypeIndex=2 then
    FilterGridValues(SGrid5,Copy(S5FilterLabel.Caption,6,Length(S5FilterLabel.Caption)-5),S5PSpin.Value,True);
  RebuildFreqTable:=True;
end;

procedure TMainForm.S2FilterbtnClick(Sender: TObject);
begin
  if Sheet2.TypeIndex=1 then
    FilterGridValues(SGrid2,Copy(S2FilterLabel.Caption,6,Length(S2FilterLabel.Caption)-5),S2PSpin.Value,False)
  else if Sheet2.TypeIndex=2 then
    FilterGridValues(SGrid2,Copy(S2FilterLabel.Caption,6,Length(S2FilterLabel.Caption)-5),S2PSpin.Value,True);
  RebuildFreqTable:=True;
end;

procedure TMainForm.SGrid5SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var S:String;
begin
  if SGrid5.Tag=5 then
  begin
    CanSelect:=True;
    Exit;
  end;
  GridSelectCell(Sheet5,5,SGrid5,S5Combo,S5NLabel,ACol,ARow);
  if CheckGridCell(SGrid5,S5TypeLabel)=False then
  begin
    CanSelect:=False;
    Exit;
  end;
  case Sheet5.TypeIndex of
    1:S:=IntToStr(Sheet5.DValue1);
    2:begin
        S:=FloatToStr(Sheet5.DValue2);
        if (Pos('E',S)>0) or (Pos('e',S)>0) then S:='0.00';
      end;
    3:S:=Sheet5.DValue3;
    4:S:='['+IntToStr(Sheet5.LBoundI)+','+IntToStr(Sheet5.UBoundI)+']';
  end;
  if S5Combo.Visible then S5Combo.Visible:=False;
  if {not(goEditing in SGrid5.Options)}Sheet5.ForceValue and CanEditCell(ACol,ARow,ActiveSheet) then
  begin
    if not(DVCheck.Checked) then
    begin
      if S5Combo.Items.IndexOf(S)>0 then
        S5Combo.Items.Delete(S5Combo.Items.IndexOf(S));
      if S5Combo.Items.Count>0 then
        S5Combo.ItemIndex:=0;
    end;
    if DVCheck.Checked then
    begin
      if not(S5Combo.Items.IndexOf(S)>=0) then
        S5Combo.Items.Add(S);
      if (Length(SGrid5.Cells[ACol,ARow])>0) and (S5Combo.Items.IndexOf(SGrid5.Cells[ACol,ARow])>=0) then
        S5Combo.ItemIndex:=S5Combo.Items.IndexOf(SGrid5.Cells[ACol,ARow])
      else
        S5Combo.ItemIndex:=S5Combo.Items.IndexOf(S);
    end;
    S5Combo.Visible:=False;
    MoveCombo(SGrid5,S5Combo,ACol,ARow);
    S5Combo.Visible:=True;
    S5Combo.SetFocus;
    S5Combo.SelectAll;
  end;
  if DVCheck.Checked and (Length(SGrid5.Cells[ACol,ARow])=0) then
    SGrid5.Cells[ACol,ARow]:=S;
  if Length(SGrid5.Cells[ACol,ARow])>0 then
    S5Combo.Hint:='Current value: '+SGrid5.Cells[ACol,ARow]
  else
    S5Combo.Hint:='No value is set';
end;

procedure TMainForm.AnalyzeReportbtnDropDownClick(Sender: TObject);
var i,l:Integer;
begin
  ReportEdit.Lines.Clear;
  ReportEdit.Lines.Append('Number Of Data= '+anaNLabel.Caption);
  ReportEdit.Lines.Append('Minimum= '+anaMinLabel.Caption+' Maximum= '+anaMaxLabel.Caption+' Value Range= '+anaRangeLabel.Caption);
//  ReportEdit.Lines.Append('*************************************************************');
//  ReportEdit.Lines.Append('*************************************************************');
  ReportEdit.Lines.Append('Mean Of Data= '+MeanLabel.Caption);
  ReportEdit.Lines.Append('Middle Of Data= '+MiddleLabel.Caption+'      Mode Of Data= '+ModeLabel.Caption);
//  ReportEdit.Lines.Append('**************************************************************');
  ReportEdit.Lines.Append('Average Of Deviation(AD)= '+ADLabel.Caption+'      Variance= '+VarLabel.Caption);
  ReportEdit.Lines.Append('Standard Deviation= '+SDLabel.Caption+'      Coefficient Of Variation(CV)= '+CVLabel.Caption);
  ReportEdit.Lines.Append('');
  l:=Length(ReportEdit.Lines.Strings[0]);
  for i:=1 to (ReportEdit.Lines.Count-1) do
    if Length(ReportEdit.Lines.Strings[i])>l then
      l:=Length(ReportEdit.Lines.Strings[i]);
  ReportEdit.Lines.Append(StringOfChar('*',l));
  ReportEdit.Lines.Insert(2,StringOfChar('*',l));
  ReportEdit.Lines.Insert(3,StringOfChar('*',l));
  ReportEdit.Lines.Insert(6,StringOfChar('*',l));
  ReportEdit.Lines.Append('>End Of Report');
end;

procedure TMainForm.SaveAsPicture1Click(Sender: TObject);
var S:String;
begin
  SD1.Filter:='Bitmap Files(*.BMP)|*.BMP|All Files|*.*';
  SD1.FilterIndex:=1;
  if SD1.Execute then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.BMP' then
      S:='.bmp'
    else
      S:='';
    Chart1.SaveToBitmapFile(SD1.FileName+S);
  end;
end;

procedure TMainForm.SaveAsMetafile1Click(Sender: TObject);
var S:String;
begin
  SD1.Filter:='Metafiles(*.WMF)|*.WMF|All Files|*.*';
  SD1.FilterIndex:=1;
  if SD1.Execute then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.WMF' then
      S:='.wmf'
    else
      S:='';
    Chart1.SaveToMetafile(SD1.FileName+S);
  end;
end;

procedure TMainForm.Add1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  HelpLabel.Caption:='Add Button:'+Chr(13)+Chr(13)+'   Adds the new value to the value list';
end;

procedure TMainForm.FxLabelMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  HelpLabel.Caption:='F(X):'+Chr(13)+Chr(13)+'Shows the filter formula of the selected data sheet(Double click on it to edit formula)';
end;

procedure TMainForm.ForceValueMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  HelpLabel.Caption:='Enables/Disables using a value list for the values of the selected data sheet';
end;

procedure TMainForm.UseFilterMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  HelpLabel.Caption:='Enables/Disables using a formula filtering for the selected data sheet';
end;

procedure TMainForm.TypeComboDropDown(Sender: TObject);
begin
  HelpLabel.Caption:='Select a data type from the list for the selected data sheet';
end;

procedure TMainForm.HelpLabelMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  if not(HelpLabel.Font.Color=clRed) then
    HelpLabel.Caption:='displays a short description of items';
end;

procedure TMainForm.ReportPopupOpen(Sender: TObject);
begin
  RsizeCombo.ItemIndex:=RSizeCombo.Items.IndexOf(IntToStr(ReportEdit.Font.Size));
  RColorBtn.SelectedColor:=ReportEdit.Font.Color;
end;

procedure TMainForm.Label15MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value describes the coefficient of variation of the data, which is calculated based on the formula you see:';
  Fml.Picture.Assign(CVFml.Picture);
  Fml.Tag:=4;
  Fml.Stretch:=False;
end;

procedure TMainForm.Label14MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value determines the standard deviation of the data, which is calculated based on the formula you see:';
  Fml.Picture.Assign(SFml.Picture);
  Fml.Tag:=3;
  Fml.Stretch:=False;
end;

procedure TMainForm.Label13MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value determines the variance of the data, which is calculated based on the formula you see:';
  Fml.Picture.Assign(S2Fml.Picture);
  Fml.Tag:=2;
  Fml.Stretch:=True;
end;

procedure TMainForm.Label12MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value describes the average of deviation(AD) of the data,which is calculated based on the formula you see:';
  Fml.Picture.Assign(ADFml.Picture);
  Fml.Tag:=1;
  Fml.Stretch:=False;
end;

procedure TMainForm.Label11MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value describes the mode value of the data, the value that has the greatest frequency in data list';
  Fml.Picture.Assign(nil);
  Fml.Tag:=-1;
end;

procedure TMainForm.Label10MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value shows the middle value of the data, the value that is in the middle position in data list';
  Fml.Picture.Assign(nil);
  Fml.Tag:=-1;
end;

procedure TMainForm.Label9MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Description: This value shows the mean value of the data';
  Fml.Picture.Assign(nil);
end;

procedure TMainForm.Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  anaDesLabel.Caption:='Description: This value shows the mean value of the data';
  if Fml.Tag<>0 then
    Fml.Picture.Assign(nil);
end;

procedure TMainForm.Image2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  if Fml.Tag=2 then Exit;
  anaDesLabel.Caption:='Description: This value determines the variance of the data, which is calculated based on the formula you see:';
  Fml.Picture.Assign(S2Fml.Picture);
  Fml.Tag:=2;
end;

procedure TMainForm.BarBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Bar Chart';
end;

procedure TMainForm.HBarBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Horizontal Bar Chart';
end;

procedure TMainForm.AreaBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Area Chart';
end;

procedure TMainForm.PointBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Point Chart:Like a linear bar, but without lines';
end;

procedure TMainForm.ChartLabelMouseEnter(Sender: TObject);
begin
  ChartLabel.Caption:='{Chart Type Description}';
end;

procedure TMainForm.PieBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Pie Chart: A chart without any axis';
end;

procedure TMainForm.FastBtnMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  ChartLabel.Caption:='Fast Line Chart:Like a linear chart, but only at 2 dimensions';
end;

procedure TMainForm.Print1Click(Sender: TObject);
begin
  if PSD1.Execute then
    Chart1.Print;
end;

procedure TMainForm.CopyAsPicture1Click(Sender: TObject);
begin
  Chart1.CopyToClipboardBitmap;
end;

procedure TMainForm.CopyAsMetafile2Click(Sender: TObject);
begin
  Chart1.CopyToClipboardMetafile(False);
end;

procedure TMainForm.istyle2Change(Sender: TObject);
begin
  Chart1.Legend.Frame.Style:=(Sender as Tpenstylecombo).Selection;
  ChartForm.Chart1.Legend.Frame.Style:=(Sender as Tpenstylecombo).Selection;
end;

procedure TMainForm.iwidth2Change(Sender: TObject);
begin
  Chart1.Legend.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  ChartForm.Chart1.Legend.Frame.Width:=StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex]);
  if StrToInt((Sender as Tpenwidthcombo).Items.Strings[(Sender as Tpenwidthcombo).ItemIndex])>1 then
    istyle2.Enabled:=False
  else
    istyle2.Enabled:=True;
  Label38.Enabled:=istyle2.Enabled;
end;

procedure TMainForm.G1SortClick(Sender: TObject);
begin
  RefreshSheetGrid(SGrid1);
end;

procedure TMainForm.G2SortClick(Sender: TObject);
begin
  RefreshSheetGrid(SGrid2);
end;

procedure TMainForm.G3SortClick(Sender: TObject);
begin
  RefreshSheetGrid(SGrid3);
end;

procedure TMainForm.G4SortClick(Sender: TObject);
begin
  RefreshSheetGrid(SGrid4);
end;

procedure TMainForm.G5SortClick(Sender: TObject);
begin
  RefreshSheetGrid(SGrid5);
end;

procedure TMainForm.StaticText32Click(Sender: TObject);
var m:TMouse;
begin
  m:=TMouse.Create;
  GridPopup1.Popup(m.CursorPos.X,m.CursorPos.Y);
  m.Free;
end;

procedure TMainForm.StaticText33Click(Sender: TObject);
var m:TMouse;
begin
  m:=TMouse.Create;
  GridPopup3.Popup(m.CursorPos.X,m.CursorPos.Y);
  m.Free;
end;

procedure TMainForm.StaticText34Click(Sender: TObject);
var m:TMouse;
begin
  m:=TMouse.Create;
  GridPopup5.Popup(m.CursorPos.X,m.CursorPos.Y);
  m.Free;
end;

procedure TMainForm.StaticText35Click(Sender: TObject);
var m:TMouse;
begin
  m:=TMouse.Create;
  GridPopup4.Popup(m.CursorPos.X,m.CursorPos.Y);
  m.Free;
end;

procedure TMainForm.StaticText36Click(Sender: TObject);
var m:TMouse;
begin
  m:=TMouse.Create;
  GridPopup2.Popup(m.CursorPos.X,m.CursorPos.Y);
  m.Free;
end;

procedure TMainForm.ListView1Click(Sender: TObject);
begin
  ListView.Click;
end;

procedure TMainForm.IconView1Click(Sender: TObject);
begin
  IconView.Click;
end;

procedure TMainForm.ToolBarView1Click(Sender: TObject);
begin
  ToolView.Click;
end;

procedure TMainForm.MenuView1Click(Sender: TObject);
begin
  MenuView.Click;
end;

procedure TMainForm.StatusBar1Click(Sender: TObject);
begin
  StatusControlBar.Visible:=StatusBar1.Checked;
  if StatusBar1.Checked then
    MainForm.Height:=MainForm.Height+StatusControlBar.Height
  else
    MainForm.Height:=MainForm.Height-StatusControlBar.Height;
end;

procedure TMainForm.About1Click(Sender: TObject);
begin
  MainAboutBox.Visible:=False;
  MainAboutBox.ShowModal;
end;

procedure TMainForm.ApplicationEvents1Message(var Msg: tagMSG;
  var Handled: Boolean);
begin
  if (Msg.message=WM_SYSCOMMAND) and (Msg.wParam=SC_AboutItem) then
  begin
    About1.Click;
    Handled:=True;
  end;
end;

procedure TMainForm.DataSettings1Click(Sender: TObject);
begin
  SettingSection.Click;
end;

procedure TMainForm.DataSheetMainMenuClick(Sender: TObject);
var B:Boolean;
begin
//  Sheet1Section.Click;
  B:=True;
  case ActiveSheet of
    1: begin
         MSort.OnClick:=G1Sort.OnClick;
         MGretestColumn.OnClick:=GreatestColumnWidth1.OnClick;
         MSmallestColumn.OnClick:=Smallest1.OnClick;
         MDefault.OnClick:=Default1.OnClick;
         MFind.OnClick:=G1Find.OnClick;
         MReplace.OnClick:=G1Replace.OnClick;
         MPrint.OnClick:=G1Print.OnClick;
         MFont.OnClick:=G1Font.OnClick;
{         MShowBackImage.OnClick:=G1ShowBackImage.OnClick;
         MChoosePicture.OnClick:=G1ChoosePicture.OnClick;
         MNoBackground.OnClick:=G1NoBackClick;
         MStyleDefault.Checked:=G1SDefault.Checked;
         MStyleClassic.Checked:=G1SClassic.Checked;
         MStyleFlat.Checked:=G1SFlat.Checked;
         MStyleDefault.OnClick:=G1SDefault.OnClick;
         MStyleClassic.OnClick:=G1SClassic.OnClick;
         MStyleFlat.OnClick:=G1SFlat.OnClick;
         MShowBackImage.Checked:=G1ShowBackImage.Checked;
         MShowBackImage.OnClick:=G1ShowBackImage.OnClick;
         MGBackPicture.Enabled:=G1ImageSub.Enabled;
         MImageDefault.Checked:=G1Default.Checked;
         MChoosePicture.Checked:=G1ChoosePicture.Checked;
         MNoBackground.Checked:=G1NoBack.Checked;
         MImageDefault.OnClick:=G1Default.OnClick;
         MChoosePicture.OnClick:=G1ChoosePicture.OnClick;
         MNoBackground.OnClick:=G1NoBack.OnClick;}
       end;
    2: begin
         MSort.OnClick:=G2Sort.OnClick;
         MGretestColumn.OnClick:=GreatestColumnWidth2.OnClick;
         MSmallestColumn.OnClick:=Smallest2.OnClick;
         MDefault.OnClick:=Default2.OnClick;
         MFind.OnClick:=G2Find.OnClick;
         MReplace.OnClick:=G2Replace.OnClick;
         MPrint.OnClick:=G2Print.OnClick;
         MFont.OnClick:=G2Font.OnClick;
{         MShowBackImage.OnClick:=G2ShowBackImage.OnClick;
         MChoosePicture.OnClick:=G2ChoosePicture.OnClick;
         MNoBackground.OnClick:=G2NoBackClick;}
       end;
    3: begin
         MSort.OnClick:=G3Sort.OnClick;
         MGretestColumn.OnClick:=GreatestColumnWidth3.OnClick;
         MSmallestColumn.OnClick:=Smallest3.OnClick;
         MDefault.OnClick:=Default3.OnClick;
         MFind.OnClick:=G3Find.OnClick;
         MReplace.OnClick:=G3Replace.OnClick;
         MPrint.OnClick:=G3Print.OnClick;
         MFont.OnClick:=G3Font.OnClick;
{         MShowBackImage.OnClick:=G3ShowBackImage.OnClick;
         MChoosePicture.OnClick:=G3ChoosePicture.OnClick;
         MNoBackground.OnClick:=G3NoBackClick;}
       end;
    4: begin
         MSort.OnClick:=G4Sort.OnClick;
         MGretestColumn.OnClick:=GreatestColumnWidth4.OnClick;
         MSmallestColumn.OnClick:=Smallest4.OnClick;
         MDefault.OnClick:=Default4.OnClick;
         MFind.OnClick:=G4Find.OnClick;
         MReplace.OnClick:=G4Replace.OnClick;
         MPrint.OnClick:=G4Print.OnClick;
         MFont.OnClick:=G4Font.OnClick;
{         MShowBackImage.OnClick:=G4ShowBackImage.OnClick;
         MGBackPicture.OnClick:=G4ChoosePicture.OnClick;
         MChoosePicture.OnClick:=G4NoBackClick;}
       end;
    5: begin
         MSort.OnClick:=G5Sort.OnClick;
         MGretestColumn.OnClick:=GreatestColumnWidth5.OnClick;
         MSmallestColumn.OnClick:=Smallest5.OnClick;
         MDefault.OnClick:=Default5.OnClick;
         MFind.OnClick:=G5Find.OnClick;
         MReplace.OnClick:=G5Replace.OnClick;
         MPrint.OnClick:=G5Print.OnClick;
         MFont.OnClick:=G5Font.OnClick;
{         MShowBackImage.OnClick:=G5ShowBackImage.OnClick;
         MGBackPicture.OnClick:=G5ChoosePicture.OnClick;
         MChoosePicture.OnClick:=G5NoBackClick;}
       end;
    else
    begin
      MSort.Enabled:=False;
      MResize.Enabled:=False;
      MFind.Enabled:=False;
      MReplace.Enabled:=False;
      MPrint.Enabled:=False;
      MGStyle.Enabled:=False;
      MFont.Enabled:=False;
{      MShowBackImage.Enabled:=False;
      MGBackPicture.Enabled:=False;}
    end;
  end;
end;

procedure TMainForm.DataSheet2Click(Sender: TObject);
begin
  PrintDataSheet.Click;
end;

procedure TMainForm.DataSheet3Click(Sender: TObject);
begin
  Sheet3Section.Click;
end;

procedure TMainForm.DataSheet4Click(Sender: TObject);
begin
  Sheet4Section.Click;
end;

procedure TMainForm.DataSheet5Click(Sender: TObject);
begin
  Sheet5Section.Click;
end;

procedure TMainForm.Charts1Click(Sender: TObject);
begin
  ChartSection.Click;
end;

procedure TMainForm.DataAnalyzing1Click(Sender: TObject);
begin
  AnalyzeSection.Click;
end;

procedure TMainForm.Chart3Click(Sender: TObject);
var B:Boolean;
begin
  if ChartSection.Down or ToolChart.Down or IConChart.Down or MChartSection.Checked then
    B:=True
  else
    B:=False;
  c1.Enabled:=B;
  PrintChart1.Enabled:=B;
  c3.Enabled:=B;
  c4.Enabled:=B;
  c5.Enabled:=B;
  c6.Enabled:=B;
end;

procedure TMainForm.c1Click(Sender: TObject);
begin
  DetachBtn.Click;
end;

procedure TMainForm.PrintChart1Click(Sender: TObject);
begin
  PrintBtn.Click;
end;

procedure TMainForm.c3Click(Sender: TObject);
begin
  Chart1.CopyToClipboardBitmap;
end;

procedure TMainForm.c4Click(Sender: TObject);
begin
  Chart1.CopyToClipboardMetafile(False);
end;

procedure TMainForm.c5Click(Sender: TObject);
begin
  SaveAsPicture1.Click;
end;

procedure TMainForm.c6Click(Sender: TObject);
begin
  SaveAsMetafile1.Click;
end;

procedure TMainForm.Exit1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TMainForm.ChartToolbar1Click(Sender: TObject);
begin
  ChartToolbar.Visible:=ChartToolbar1.Checked;
end;

procedure TMainForm.StatManHelp1Click(Sender: TObject);
begin
  Application.HelpFile:='STATMAN.HLP';
  Application.HelpCommand(HELP_FINDER,0);
  if PersianHelp.Checked then
    Application.HelpFile:='STATMANPERSIAN.HLP';
end;

procedure TMainForm.Default1Click(Sender: TObject);
begin
  SGrid1.DefaultColWidth:=64;
  SGrid1.ColWidths[0]:=22;
end;

procedure TMainForm.GreatestColumnWidth1Click(Sender: TObject);
begin
  ResizeGrid(grkGreatest,SGrid1);
end;

procedure TMainForm.Smallest1Click(Sender: TObject);
begin
  ResizeGrid(grkSmallest,SGrid1);
end;

procedure TMainForm.Default2Click(Sender: TObject);
begin
  SGrid2.DefaultColWidth:=64;
  SGrid2.ColWidths[0]:=22;
end;

procedure TMainForm.Default3Click(Sender: TObject);
begin
  SGrid3.DefaultColWidth:=64;
  SGrid3.ColWidths[0]:=22;
end;

procedure TMainForm.Default4Click(Sender: TObject);
begin
  SGrid4.DefaultColWidth:=64;
  SGrid4.ColWidths[0]:=22;
end;

procedure TMainForm.Default5Click(Sender: TObject);
begin
  SGrid5.DefaultColWidth:=64;
  SGrid5.ColWidths[0]:=22;
end;

procedure TMainForm.GreatestColumnWidth2Click(Sender: TObject);
begin
  ResizeGrid(grkGreatest,SGrid2);
end;

procedure TMainForm.GreatestColumnWidth3Click(Sender: TObject);
begin
  ResizeGrid(grkGreatest,SGrid3);
end;

procedure TMainForm.GreatestColumnWidth4Click(Sender: TObject);
begin
  ResizeGrid(grkGreatest,SGrid4);
end;

procedure TMainForm.GreatestColumnWidth5Click(Sender: TObject);
begin
  ResizeGrid(grkGreatest,SGrid5);
end;

procedure TMainForm.Smallest2Click(Sender: TObject);
begin
  ResizeGrid(grkSmallest,SGrid2);
end;

procedure TMainForm.Smallest3Click(Sender: TObject);
begin
  ResizeGrid(grkSmallest,SGrid3);
end;

procedure TMainForm.Smallest4Click(Sender: TObject);
begin
  ResizeGrid(grkSmallest,SGrid4);
end;

procedure TMainForm.Smallest5Click(Sender: TObject);
begin
  ResizeGrid(grkSmallest,SGrid5);
end;

procedure TMainForm.AppendFromFile1Click(Sender: TObject);
var Row,Col:Integer;
begin
  if OpenGrid.Execute then
  begin
    RefreshSheetGrid(FGrid);
    for Row:=1 to (FGrid.RowCount-2) do
    begin
      for Col:=1 to (FGrid.ColCount-1) do
      begin
        if Length(FGrid.Cells[Col,Row])=0 then
          Break;
      end;
      if (Col<FGrid.ColCount) and (Length(FGrid.Cells[Col,Row])=0) then
        Break;
    end;
    if (Row=FGrid.RowCount-1) and (Col=FGrid.ColCount) then
      Exit;
    LoadGridFromFile(OpenGrid.FileName,True,Col,Row);
  end;
end;

procedure TMainForm.SaveGridPopupPopup(Sender: TObject);
var TypeIndex,Count:Byte;
begin
  AppendEmpty.Visible:=False;
  AppendS1.Visible:=False;
  AppendS2.Visible:=False;
  AppendS3.Visible:=False;
  AppendS4.Visible:=False;
  AppendS5.Visible:=False;
  Count:=0;
  case ActiveSheet of
    1: TypeIndex:=Sheet1.TypeIndex;
    2: TypeIndex:=Sheet2.TypeIndex;
    3: TypeIndex:=Sheet3.TypeIndex;
    4: TypeIndex:=Sheet4.TypeIndex;
    5: TypeIndex:=Sheet5.TypeIndex;
  end;
  if (ActiveSheet<>1) and (Sheet1.Used) and (Sheet1.TypeIndex=TypeIndex) then
  begin
    Inc(Count);
    AppendS1.Visible:=True;
  end;
  if (ActiveSheet<>2) and (Sheet2.Used) and (Sheet2.TypeIndex=TypeIndex) then
  begin
    Inc(Count);
    AppendS2.Visible:=True;
  end;
  if (ActiveSheet<>3) and (Sheet3.Used) and (Sheet3.TypeIndex=TypeIndex) then
  begin
    Inc(Count);
    AppendS3.Visible:=True;
  end;
  if (ActiveSheet<>4) and (Sheet4.Used) and (Sheet4.TypeIndex=TypeIndex) then
  begin
    Inc(Count);
    AppendS4.Visible:=True;
  end;
  if (ActiveSheet<>5) and (Sheet5.Used) and (Sheet5.TypeIndex=TypeIndex) then
  begin
    Inc(Count);
    AppendS5.Visible:=True;
  end;
  if Count=0 then
    AppendEmpty.Visible:=True;
end;

procedure TMainForm.AppendS5Click(Sender: TObject);
begin
  case (Sender as TMenuItem).Tag of
    1:AppendDataSheet(SGrid1);
    2:AppendDataSheet(SGrid2);
    3:AppendDataSheet(SGrid3);
    4:AppendDataSheet(SGrid4);
    5:AppendDataSheet(SGrid5);
  end;
end;


procedure TMainForm.BrowseImageClick(Sender: TObject);
begin
  if OPD1.Execute then
  begin
    backimage.Picture.LoadFromFile(OPD1.FileName);
    Chart1.BackImage.LoadFromFile(OPD1.FileName);
    ChartForm.Chart1.BackImage.LoadFromFile(OPD1.FileName);
  end;
end;

procedure TMainForm.TileRadioClick(Sender: TObject);
begin
  case (Sender as TRadioButton).Tag of
    1:begin
        Chart1.BackImageMode:=pbmTile;
        ChartForm.Chart1.BackImageMode:=pbmTile;
      end;
    2:begin
        Chart1.BackImageMode:=pbmStretch;
        ChartForm.Chart1.BackImageMode:=pbmStretch;
      end;
    3:begin
        Chart1.BackImageMode:=pbmCenter;
        ChartForm.Chart1.BackImageMode:=pbmCenter;
      end;
  end;
end;

procedure TMainForm.PutInsideClick(Sender: TObject);
begin
  Chart1.BackImageInside:=putinside.Checked;
  ChartForm.Chart1.BackImageInside:=putinside.Checked;
end;

procedure TMainForm.ColorFChangeColor(Sender: TObject);
begin
  Chart1.Foot.Color:=(Sender as TColorSelector).Color;
  ChartForm.Chart1.Foot.Color:=(Sender as TColorSelector).Color;
end;

procedure TMainForm.Title0Click(Sender: TObject);
begin
  Chart1.LeftAxis.Title.Angle:=(Sender as TRadioButton).Tag;
  ChartForm.Chart1.LeftAxis.Title.Angle:=(Sender as TRadioButton).Tag;
end;

procedure TMainForm.AxisTitleRotationOpen(Sender: TObject);
begin
   AxisTitleR.Position:=cHART1.LeftAxis.Title.Angle;
end;

procedure TMainForm.SaveAsPicture2Click(Sender: TObject);
begin
  SaveAsPicture1.Click;
end;

procedure TMainForm.SaveAsMetafile2Click(Sender: TObject);
begin
  SaveAsMetafile1.Click;
end;

procedure TMainForm.SD1CanClose(Sender: TObject; var CanClose: Boolean);
var S:String; R:Integer;
begin
  if SD1.Filter='Bitmap Files(*.BMP)|*.BMP|All Files|*.*' then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.BMP' then
      S:='.bmp'
    else
      S:='';
    if FileExists(SD1.FileName+S) then
    begin
      r:=MessageDlg(SD1.FileName+S+' already exists.'+Chr(13)+'Do you want to replace it?',mtCustom,[mbYes,mbNo],0);
      if r=mrYes then
        CanClose:=True
      else
        CanClose:=False;
    end;
  end
  else if SD1.Filter='Metafiles(*.WMF)|*.WMF|All Files|*.*' then
  begin
    S:=RightStr(ExtractFileName(SD1.FileName),4);
    if UpperCase(S)<>'.WMF' then
      S:='.wmf'
    else
      S:='';
    if FileExists(SD1.FileName+S) then
    begin
      r:=MessageDlg(SD1.FileName+S+' already exists.'+Chr(13)+'Do you want to replace it?',mtCustom,[mbYes,mbNo],0);
      if r=mrYes then
        CanClose:=True
      else
        CanClose:=False;
    end;
  end;
end;

procedure TMainForm.EnglishHelpClick(Sender: TObject);
begin
  if EnglishHelp.Checked then
    PersianHelp.Checked:=False
  else if not(PersianHelp.Checked) then
    EnglishHelp.Checked:=True;
end;

procedure TMainForm.PersianHelpClick(Sender: TObject);
begin
  if PersianHelp.Checked then
    EnglishHelp.Checked:=False
  else if not(EnglishHelp.Checked) then
    PersianHelp.Checked:=True;
end;

procedure TMainForm.StatManHelp2Click(Sender: TObject);
begin
  Application.HelpFile:='STATMANPERSIAN.HLP';
  Application.HelpCommand(HELP_FINDER,0);
  if EnglishHelp.Checked then
    Application.HelpFile:='STATMAN.HLP';
end;

procedure TMainForm.HideWhenMinimized1Click(Sender: TObject);
begin
  TTIcon.MinimizeToTray:=(Sender as TMenuItem).Checked;
end;

procedure TMainForm.OpenStatMan1Click(Sender: TObject);
begin
  if IsMinimized then
  begin
    FloatingRectangles(False, True);
    TTIcon.ShowMainForm;
  end;
  TTIcon.IconVisible:=False;
  TrayTimer.Enabled:=False;
  IsMinimized:=False;
//  if MainAboutBox.Visible then
  //  MainAboutBox.Hide;
//  Application.Restore;
end;

procedure TMainForm.Help2Click(Sender: TObject);
begin
  Application.HelpCommand(HELP_FINDER,0);
end;

procedure TMainForm.About2Click(Sender: TObject);
begin
  MainAboutBox.Show;
end;

procedure TMainForm.Exit2Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TMainForm.TTIconBalloonHintClick(Sender: TObject);
begin
  OpenStatMan1.Click;
end;

procedure TMainForm.TrayTimerTimer(Sender: TObject);
begin
  TTIcon.HideBalloonHint;
  TrayTimer.Enabled:=False;
end;

procedure TMainForm.TTIconClick(Sender: TObject);
begin
  if IsMinimized then
    TTIcon.PopupAtCursor;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
var i:Integer;
begin
  SetLength(ChartColors,0);
  for i:=1 to 5 do
    SetLength(LockedCells[i],0);
  WriteConfiguration;
end;
procedure TMainForm.SettingGridMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  RowSizing:=True;
end;

procedure TMainForm.SettingGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  RowSizing:=False;
end;

procedure TMainForm.SettingGridMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    ArrangeCheckBoxes;
end;

procedure TMainForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var m:TMouse;
begin
  m:=TMouse.Create;
  if Key=VK_F1 then
  begin
    if EnglishHelp.Checked then
    begin
      SetHelpTopic(lIDEngish);
      Application.HelpFile:=ExeDir+'STATMAN.HLP'
    end
    else if PersianHelp.Checked then
    begin
      SetHelpTopic(lIDPersian);
      Application.HelpFile:=ExeDir+'STATMANPERSIAN.HLP';
    end;
    if PromptForLan.Checked then
    begin
      if PromptAsList.Checked then
      begin
        HelpLangForm.Left:=m.CursorPos.X-2;
        if (HelpLangForm.Width-(Screen.Width-m.CursorPos.X))>0 then
          HelpLangForm.Left:=HelpLangForm.Left-(HelpLangForm.Width-(Screen.Width-m.CursorPos.X));
        HelpLangForm.Top:=m.CursorPos.Y;
        if (HelpLangForm.Height-(Screen.Height-m.CursorPos.Y))>0 then
          HelpLangForm.Top:=HelpLangForm.Top-(HelpLangForm.Height-(Screen.Height-m.CursorPos.Y));
        if EnglishHelp.Checked then
          HelpLangForm.LangList.ItemIndex:=0
        else if PersianHelp.Checked then
          HelpLangForm.LangList.ItemIndex:=1;
        HelpLangForm.Show;
        HelpLangForm.LangList.SetFocus;
      end
      else if PromptAsMenu.Checked then
        HelpLangMenu.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    end
    else
      Application.HelpContext(HelpTopic);
  end;
  m.Free;
  if (Key=VK_END) and (ssAlt in Shift) then
    GotoNextSection
  else if (Key=VK_HOME) and (ssAlt in Shift) then
    GotoPreviousSection;
  if (Key=Ord('P')) and (ssShift in Shift) and (ssCtrl in Shift) then
    Clipboard.Assign(MainForm.GetFormImage);
end;

procedure TMainForm.MemberPane2Click(Sender: TObject);
begin
  MemberGroup.Visible:=(Sender as TMenuItem).Checked;
  ResizeForm;
end;

procedure TMainForm.TTIconMinimizeToTray(Sender: TObject);
begin
  FloatingRectangles(True,True);
  TTIcon.IconVisible:=True;
  TTIcon.ShowBalloonHint('StatMan','Click here to open StatMan.',bitInfo,10);
  TTIcon.Icon:=TrayIconImage.Picture.Icon;
  TrayTimer.Enabled:=True;
  IsMinimized:=True;
end;

procedure TMainForm.PromptForLanClick(Sender: TObject);
begin
  if ShowDefaultLan.Checked then
    ShowDefaultLan.Checked:=False;

  if not(ShowDefaultLan.Checked) then
  begin
    PromptForLan.Checked:=True;
    HelpLangForm.PromptCheck.Checked:=False;
  end;  
end;

procedure TMainForm.ShowDefaultLanClick(Sender: TObject);
begin
  if PromptForLan.Checked then
    PromptForLan.Checked:=False;
  if not(PromptForLan.Checked) then
    ShowDefaultLan.Checked:=True;
end;

procedure TMainForm.English1Click(Sender: TObject);
begin
  VisibleMenuHeaders(False);
  SwitchProgramLanguage(lIDEngish);
  VisibleMenuHeaders(True);
  English1.Checked:=True;
  Persian1.Checked:=False;
  FastCalcToolbar.Repaint;
end;

procedure TMainForm.Persian1Click(Sender: TObject);
begin
  if (Persian1.Tag>0) or (MessageDlg('If Persian is not one of your system regional languages,or '+
    'your system does not have Persian fonts,don`t change '+
    'program language to Persian.'+Chr(13)+'If you can see a correct Persian sentence below change the program language to Persian.'+
    Chr(13)+'Sample Sentence: ".Çíä Ìãáå ÝÇÑÓí ÈØæÑ ÏÑÓÊí äãÇíÔ ÏÇÏå ÔÏå ÇÓÊ" '+Chr(13)+
    'Now,do you want to change program language to Persian?',mtConfirmation,[mbYes,mbNo],0)=mrYes) then
  begin
    Persian1.Tag:=1;
    VisibleMenuHeaders(False);
    SwitchProgramLanguage(lIDPersian);
    VisibleMenuHeaders(True);
    English1.Checked:=False;
    Persian1.Checked:=True;
    FastCalcToolbar.Repaint;
  end;  
end;

procedure TMainForm.StrLimitsTableSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  if ACol=0 then
  begin
    CanSelect:=False;
    Exit;
  end;
  if (ACol=1) and (ARow=2) then
  begin
    CharStyleCombo.Visible:=False;
    CharStyleCombo.ItemIndex:=CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[ACol,ARow]);
    MoveCombo(StrLimitsTable,CharStyleCombo,ACol,ARow);
    CharStyleCombo.Visible:=True;
    StrLimitsTable.OnExit:=nil;
    CharStyleCombo.SetFocus;
    StrLimitsTable.OnExit:=StrLimitsTableExit;
  end
  else
    CharStyleCombo.Visible:=False;
  if Length(StrLimitsTable.Cells[1,1])=0 then
    StrLimitsTable.Cells[1,1]:='5'
  else if not(IsValidInt(StrLimitsTable.Cells[1,1])) then
  begin
    ShowMessage('"'+StrLimitsTable.Cells[1,1]+'" is not a valid number.');
    StrLimitsTable.Cells[1,1]:='5';
  end
  else if StrToInt(StrLimitsTable.Cells[1,1])>255 then
  begin
    ShowMessage('Please enter a value between 1 and 255 for the length of the string.');
    StrLimitsTable.Cells[1,1]:='255';
  end
  else if StrToInt(StrLimitsTable.Cells[1,1])<=0 then
  begin
    ShowMessage('Please enter a non-zero positive value for the length of the string.');
    StrLimitsTable.Cells[1,1]:='1';
  end;
end;

procedure TMainForm.CharStyleComboChange(Sender: TObject);
var AllUpper,AllLower,NoChange:Boolean; i:Integer;
begin
  AllUpper:=True;
  AllLower:=True;
  if ForceValue.Checked and (List1.Items.Count>0) then
    for i:=0 to (List1.Items.Count-1) do
    begin
      if UpperCase(List1.Items.Strings[i])<>List1.Items.Strings[i] then
        AllUpper:=False;
      if LowerCase(List1.Items.Strings[i])<>List1.Items.Strings[i] then
        AllLower:=False;
      if not(AllUpper) and not(AllLower) then
        Break;
    end;
  NoChange:=False;
  case CharStyleCombo.ItemIndex of
    0: begin
         SettingGrid.Cells[3,SettingGrid.Row]:=UpperCase(SettingGrid.Cells[3,SettingGrid.Row]);
         if not(AllUpper) then
           if MessageDlg('You have some items in the force value list that are not in uppercase style.If you change the character style of string, those values in the force value list will change.'+Chr(13)+
                  'Do you want to continue?',mtConfirmation,[mbYes,mbNo],0)=mrYes then
           for i:=0 to (List1.Items.Count-1) do
             List1.Items.Strings[i]:=UpperCase(List1.Items.Strings[i])
           else
             NoChange:=True;
       end;
    1: begin
         SettingGrid.Cells[3,SettingGrid.Row]:=LowerCase(SettingGrid.Cells[3,SettingGrid.Row]);
         if not(AllLower) then
           if MessageDlg('You have some items in the force value list that are not in lowercase style.If you change the character style of string, those values in the force value list will change.'+Chr(13)+
                  'Do you want to continue?',mtConfirmation,[mbYes,mbNo],0)=mrYes then
           for i:=0 to (List1.Items.Count-1) do
             List1.Items.Strings[i]:=LowerCase(List1.Items.Strings[i])
           else
             NoChange:=True;
       end;
    2: SettingGrid.Cells[3,SettingGrid.Row]:=UpperCase(SettingGrid.Cells[3,SettingGrid.Row][1])+LowerCase(Copy(SettingGrid.Cells[3,SettingGrid.Row],2,Length(SettingGrid.Cells[3,SettingGrid.Row])-1));
  end;
  if NoChange then
    CharStyleCombo.ItemIndex:=CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[StrLimitsTable.Col,StrLimitsTable.Row])
  else
    StrLimitsTable.Cells[StrLimitsTable.Col,
          StrLimitsTable.Row]:=CharStyleCombo.Items.Strings[
          CharStyleCombo.ItemIndex];
end;

procedure TMainForm.StrLimitsTableExit(Sender: TObject);
var B:Boolean;
begin
  StrLimitsTable.OnSelectCell(StrLimitsTable,StrLimitsTable.Col,StrLimitsTable.Row,B);
end;

procedure TMainForm.SGrid4KeyPress(Sender: TObject; var Key: Char);
var S:String[1];
begin
  if Sheet4.TypeIndex=3 then
  begin
    if (Length(SGrid4.Cells[SGrid4.Col,SGrid4.Row])=Sheet4.MaxStringLength) and (Ord(Key)<>VK_BACK) then
      Key:=Chr(0);
    case Sheet4.CharStyle of
      csUppercase: begin
                     S:=UpperCase(Key);
                     Key:=S[1];
                   end;
      csLowercase: begin
                     S:=LowerCase(Key);
                     Key:=S[1];
                   end;
    end;
  end;
end;

procedure TMainForm.SGrid1KeyPress(Sender: TObject; var Key: Char);
var S:String;
begin
  if Sheet1.TypeIndex=3 then
  begin
    if (Length(SGrid1.Cells[SGrid1.Col,SGrid1.Row])=Sheet1.MaxStringLength) and (Ord(Key)<>VK_BACK) then
      Key:=Chr(0);
    case Sheet1.CharStyle of
      csUppercase: begin
                     S:=UpperCase(Key);
                     Key:=S[1];// SGrid4.Cells[SGrid4.Col,SGrid4.Row]:=UpperCase(SGrid4.Cells[SGrid4.Col,SGrid4.Row]);
                   end;
      csLowercase: begin
                     S:=LowerCase(Key);
                     Key:=S[1];// SGrid4.Cells[SGrid4.Col,SGrid4.Row]:=LowerCase(SGrid4.Cells[SGrid4.Col,SGrid4.Row]);
                   end;
    end;
  end;
end;

procedure TMainForm.SGrid3KeyPress(Sender: TObject; var Key: Char);
var S:String;
begin
  if Sheet3.TypeIndex=3 then
  begin
    if (Length(SGrid3.Cells[SGrid3.Col,SGrid3.Row])=Sheet3.MaxStringLength) and (Ord(Key)<>VK_BACK) then
      Key:=Chr(0);
    case Sheet3.CharStyle of
      csUppercase: begin
                     S:=UpperCase(Key);
                     Key:=S[1];
                   end;
      csLowercase: begin
                     S:=LowerCase(Key);
                     Key:=S[1];
                   end;
    end;
  end;
end;

procedure TMainForm.SGrid5KeyPress(Sender: TObject; var Key: Char);
var S:String;
begin
  if Sheet5.TypeIndex=3 then
  begin
    if (Length(SGrid5.Cells[SGrid5.Col,SGrid5.Row])=Sheet5.MaxStringLength) and (Ord(Key)<>VK_BACK) then
      Key:=Chr(0);
    case Sheet5.CharStyle of
      csUppercase: begin
                     S:=UpperCase(Key);
                     Key:=S[1];
                   end;
      csLowercase: begin
                     S:=LowerCase(Key);
                     Key:=S[1];
                   end;
    end;
  end;
end;

procedure TMainForm.SGrid2KeyPress(Sender: TObject; var Key: Char);
var S:String;
begin
  if Sheet2.TypeIndex=3 then
  begin
    if (Length(SGrid2.Cells[SGrid2.Col,SGrid2.Row])=Sheet2.MaxStringLength) and (Ord(Key)<>VK_BACK) then
      Key:=Chr(0);
    case Sheet2.CharStyle of
      csUppercase: begin
                     S:=UpperCase(Key);
                     Key:=S[1];
                   end;
      csLowercase: begin
                     S:=LowerCase(Key);
                     Key:=S[1];
                   end;
    end;
  end;
end;

procedure TMainForm.NewWorkProject1Click(Sender: TObject);
var i:Integer;
begin
  if Changed then
  begin
    i:=MessageDlg('The current project you were working on is not saved.'+Chr(13)+'Do you want to save it now?',mtConfirmation,[mbYes,mbNo,mbCancel],0);
    if i=mrCancel then
      Exit
    else
    if i=mrYes then
    begin
      SaveWorkProject1.Click;
      Changed:=False;
    end;
  end;
  ResetProgram;
end;

procedure TMainForm.SetSheetDefaults;
begin
  with Sheet1 do
  begin
    FieldName:='Field 1';
    Used:=False;
    DesIndex:=0;
    ForceValue:=False;
    Filter:=False;
    FilterIndex:=5;
    VListIndex:=1;
    MaxStringLength:=5;
    CharStyle:=csNone;
  end;
  with Sheet2 do
  begin
    FieldName:='Field 2';
    Used:=False;
    DesIndex:=1;
    ForceValue:=False;
    Filter:=False;
    FilterIndex:=6;
    VListIndex:=2;
    MaxStringLength:=5;
    CharStyle:=csNone;
  end;
  with Sheet3 do
  begin
    FieldName:='Field 3';
    Used:=False;
    DesIndex:=2;
    ForceValue:=False;
    Filter:=False;
    FilterIndex:=7;
    VListIndex:=3;
    MaxStringLength:=5;
    CharStyle:=csNone;
  end;
  with Sheet4 do
  begin
    FieldName:='Field 4';
    Used:=False;
    DesIndex:=3;
    ForceValue:=False;
    Filter:=False;
    FilterIndex:=8;
    VListIndex:=4;
    MaxStringLength:=5;
    CharStyle:=csNone;
  end;
  with Sheet5 do
  begin
    FieldName:='Field 5';
    Used:=False;
    DesIndex:=4;
    ForceValue:=False;
    Filter:=False;
    FilterIndex:=9;
    VListIndex:=5;
    MaxStringLength:=5;
    CharStyle:=csNone;
  end;
end;

procedure TMainForm.ResetChart(Chart: TChart);
var i,j:Integer;
begin
  for i:=0 to (Chart.SeriesCount-1) do
    for j:=0 to (Chart.Series[i].Count-1) do
    begin
      Chart.Series[i].Delete(0);
    end;
end;

procedure TMainForm.xpButton4Click(Sender: TObject);
begin
  if PSD1.Execute then
    FreqReportEdit.Print('Frequency Table Report');
end;

procedure TMainForm.xpButton5Click(Sender: TObject);
begin
  FreqReportEdit.SelectAll;
  FreqReportEdit.CopyToClipboard;
  FreqReportEdit.SelLength:=0;
end;

procedure TMainForm.FRBoldClick(Sender: TObject);
begin
  if FRBold.Down then
    FreqReportEdit.Font.Style:=FreqReportEdit.Font.Style+[fsBold]
  else
    FreqReportEdit.Font.Style:=FreqReportEdit.Font.Style-[fsBold];
end;

procedure TMainForm.FRItalicClick(Sender: TObject);
begin
  if FRItalic.Down then
    FreqReportEdit.Font.Style:=FreqReportEdit.Font.Style+[fsItalic]
  else
    FreqReportEdit.Font.Style:=FreqReportEdit.Font.Style-[fsItalic];
end;

procedure TMainForm.FRLeftClick(Sender: TObject);
begin
  if FRLeft.Down then
    FreqReportEdit.Alignment:=taLeftJustify;
end;

procedure TMainForm.FRCenterClick(Sender: TObject);
begin
  if FRCenter.Down then
    FreqReportEdit.Alignment:=taCenter;
end;

procedure TMainForm.FRRightClick(Sender: TObject);
begin
  if FRRight.Down then
    FreqReportEdit.Alignment:=taRightJustify;
end;

procedure TMainForm.FRSizeComboChange(Sender: TObject);
begin
  FreqReportEdit.Font.Size:=StrToInt(FRSizeCombo.Text);
end;

procedure TMainForm.FRColorBtnColorSelected(Sender: TObject;
  AColor: TColor);
begin
  FreqReportEdit.Font.Color:=AColor;
end;

procedure TMainForm.SpeedButton12Click(Sender: TObject);
begin
  FD3.Font:=FreqReportEdit.Font;
  if FD3.Execute then
  begin
    FreqReportEdit.Font:=FD3.Font;
    if fsBold in FD3.Font.Style then
      FRBold.Down:=True
    else
      FRBold.Down:=False;
    if fsItalic in FD3.Font.Style then
      FRItalic.Down:=True
    else
      FRItalic.Down:=False;
    if fsUnderLine in FD3.Font.Style then
      FRUnder.Down:=True
    else
      FRUnder.Down:=False;
    FRColorBtn.SelectedColor:=FD3.Font.Color;
    FRSizeCombo.ItemIndex:=FRSizeCombo.Items.IndexOf(IntToStr(FD3.Font.Size));
  end;
end;

procedure TMainForm.ReportFreqTablebtnDropDownClick(Sender: TObject);
var ColLength:array of Integer;
    i,j,k:Integer;  Line,S:String;
begin
  FRSizeCombo.ItemIndex:=FRSizeCombo.Items.IndexOf(IntToStr(FreqReportEdit.Font.Size));
  SetLength(ColLength,TableGrid.ColCount);
  for i:=0 to (TableGrid.ColCount-1) do
  begin
    ColLength[i]:=0;
    for j:=0 to (TableGrid.RowCount-1) do
      if Length(TableGrid.Cells[i,j])>ColLength[i] then
        ColLength[i]:=Length(TableGrid.Cells[i,j]);
  end;
  FreqReportEdit.Lines.Clear;
  for i:=0 to (TableGrid.RowCount-1) do
  begin
    Line:='';
    for j:=0 to (TableGrid.ColCount-1) do
    begin
      S:=TableGrid.Cells[j,i];
      if Length(S)<ColLength[j] then
        for k:=0 to (ColLength[j]-Length(S)+1) do
          S:=S+' ';
      S:=S+'  ';
      Line:=Line+S;
    end;
    FreqReportEdit.Lines.Append(Line);
    FreqReportEdit.Lines.Add('')
  end;
end;

procedure TMainForm.CreateFrqTableColumns;
var fi,Total,i,j:Integer; fpi,pi,pci:Double; S:String;
begin
  Total:=0;
  for i:=1 to (TableGrid.RowCount-2) do
  begin
    Total:=Total+StrToInt(TableGrid.Cells[1,i]);
    TableGrid.Cells[4,i]:=IntToStr(Total);
  end;  
  for i:=1 to (TableGrid.RowCount-2) do
    for j:=2 to 6 do
    begin
      fi:=StrToInt(TableGrid.Cells[1,i]);
      fpi:=fi/Total;
      pi:=fpi*100;
      S:=FloatToStrF(fpi,ffFixed,18,4);//FloatToStr(fpi);
      TableGrid.Cells[2,i]:=S;//FloatToStrF(fpi,ffFixed,18,2);//
      S:=FloatToStr(pi);
      if Pos('E',S)>0 then
        S:=S+'%'
      else
      begin
        if pos('.',S)>0 then
          S:=Copy(S,1,pos('.',S)+2);
        S:=S+'%';
      end;
      TableGrid.Cells[3,i]:=S;

      fi:=StrToInt(TableGrid.Cells[4,i]);
      fpi:=fi/Total;
      pci:=fpi*100;
      TableGrid.Cells[5,i]:=FloatToStrF(fpi,ffFixed,18,4);//FloatToStr(fpi);
      S:=FloatToStr(pci);
      if Pos('E',S)>0 then
        S:=S+'%'
      else
      begin
        if pos('.',S)>0 then
          S:=Copy(S,1,pos('.',S)+2);
        S:=S+'%';
      end;
      TableGrid.Cells[6,i]:=S;
    end;
  TableGrid.Cells[0,TableGrid.RowCount-1]:='Total';
  TableGrid.Cells[1,TableGrid.RowCount-1]:=IntToStr(Total);
  for i:=2 to 6 do
    TableGrid.Cells[i,TableGrid.RowCount-1]:='********';
end;

procedure TMainForm.SGrid4SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
  if not(Changed) then
    Changed:=True;
  RebuildFreqTable:=True;
end;

procedure TMainForm.SGrid1SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
  if not(Changed) then
    Changed:=True;
  RebuildFreqTable:=True;
end;

procedure TMainForm.SGrid3SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
  if not(Changed) then
    Changed:=True;
  RebuildFreqTable:=True;
end;

procedure TMainForm.SGrid5SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
  if not(Changed) then
    Changed:=True;
  RebuildFreqTable:=True;
end;

procedure TMainForm.SGrid2SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
  if not(Changed) then
    Changed:=True;
  RebuildFreqTable:=True;
end;

procedure TMainForm.Label8MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Determines the number of data(N) '+
    'which is used in calculating statistical indicators.';
  Fml.Picture.Assign(nil);
end;

procedure TMainForm.Label5MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Determines the minimum value of data';
  Fml.Picture.Assign(nil);
end;

procedure TMainForm.Label6MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Determines the maximum value of data';
  Fml.Picture.Assign(nil);
end;

procedure TMainForm.Label7MouseEnter(Sender: TObject);
begin
  anaDesLabel.Caption:='Determines the range of the data '+
    'which is equal to (Max - Min)';
  Fml.Picture.Assign(nil);
end;

procedure TMainForm.GotoActiveSheet;
begin
  case ActiveSheet of
    0: begin
         SettingSection.Down:=True;
         MSettingSection.Checked:=True;
         ToolSettings.Down:=True;
         IconSettings.Down:=True;
       end;
    1: begin
         Sheet1Section.Down:=True;
         MSheet1Section.Checked:=True;
         ToolSheet1.Down:=True;
         IconSheet1.Down:=True;
       end;
    2: begin
         Sheet2Section.Down:=True;
         MSheet2Section.Checked:=True;
         ToolSheet2.Down:=True;
         IconSheet2.Down:=True;
       end;
    3: begin
         Sheet3Section.Down:=True;
         MSheet3Section.Checked:=True;
         ToolSheet3.Down:=True;
         IconSheet3.Down:=True;
       end;
    4: begin
         Sheet4Section.Down:=True;
         MSheet4Section.Checked:=True;
         ToolSheet4.Down:=True;
         IconSheet4.Down:=True;
       end;
    5: begin
         Sheet5Section.Down:=True;
         MSheet5Section.Checked:=True;
         ToolSheet5.Down:=True;
         IconSheet5.Down:=True;
       end;
    6: begin
         TableSection.Down:=True;
         MTableSection.Checked:=True;
         ToolTable.Down:=True;
         IconTable.Down:=True;
       end;
    7: begin
         ChartSection.Down:=True;
         MChartSection.Checked:=True;
         ToolChart.Down:=True;
         IconChart.Down:=True;
       end;
    8: begin
         AnalyzeSection.Down:=True;
         MAnalyzeSection.Checked:=True;
         ToolAnalyze.Down:=True;
         IconAnalyze.Down:=True;
       end;
  end;
end;

procedure TMainForm.G4ChoosePictureClick(Sender: TObject);
begin
  if OPD1.Execute then
  begin
    if G4ShowBackImage.Checked then
      SGrid4.BackGround.Bitmap.LoadFromFile(OPD1.FileName);
    GridImagePath[4]:=OPD1.FileName;
    G4ChoosePicture.Checked:=True;
    G4NoBack.Checked:=False;
    G4Default.Checked:=False;
    SGrid4.Repaint;
  end;
end;

procedure TMainForm.G4DefaultClick(Sender: TObject);
begin
  Sgrid4.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  G4ChoosePicture.Checked:=False;
  G4NoBack.Checked:=False;
  G4Default.Checked:=True;
  SGrid4.Repaint;
end;

procedure TMainForm.G4ShowBackImageClick(Sender: TObject);
begin
  G4ImageSub.Enabled:=G4ShowBackImage.Checked;
  if G4ImageSub.Enabled then
  begin
    if G4ChoosePicture.Checked then
      SGrid4.BackGround.Bitmap.LoadFromFile(GridImagePath[4])
    else if G4Default.Checked then
      SGrid4.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  end
  else
    SGrid4.BackGround.Bitmap.Assign(nil);
  SGrid4.Repaint;
end;

procedure TMainForm.G4FontClick(Sender: TObject);
begin
  FD3.Font:=SGrid4.Font;
  if FD3.Execute then
    SGrid4.Font:=FD3.Font;
end;

procedure TMainForm.G4SDefaultClick(Sender: TObject);
begin
  SGrid4.Look:=glXP;
  if not(G4SFlat.Checked) and not(G4SClassic.Checked) then
    G4SDefault.Checked:=True;
end;

procedure TMainForm.G4SClassicClick(Sender: TObject);
begin
  SGrid4.Look:=glClassic;
  if not(G4SDefault.Checked) and not(G4SFlat.Checked) then
    G4SClassic.Checked:=True;
end;

procedure TMainForm.G4SFlatClick(Sender: TObject);
begin
  SGrid4.Look:=glSoft;
  if not(G4SDefault.Checked) and not(G4SClassic.Checked) then
    G4SFlat.Checked:=True;
end;

procedure TMainForm.G4ReplaceClick(Sender: TObject);
begin
  GridReplaceForm.Grid:=SGrid4;
  GridReplaceForm.Show;
end;

procedure TMainForm.G4FindClick(Sender: TObject);
begin
  GridFindForm.Grid:=SGrid4;
  GridFindForm.Show;
end;

procedure TMainForm.G4PrintClick(Sender: TObject);
begin
  if G4PrintSettingsD.Execute then
    SGrid4.Print;
end;

procedure TMainForm.G1DefaultClick(Sender: TObject);
begin
  Sgrid1.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  G1ChoosePicture.Checked:=False;
  G1NoBack.Checked:=False;
  G1Default.Checked:=True;
  SGrid1.Repaint;
end;

procedure TMainForm.G2DefaultClick(Sender: TObject);
begin
  Sgrid2.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  G2ChoosePicture.Checked:=False;
  G2NoBack.Checked:=False;
  G2Default.Checked:=True;
  SGrid2.Repaint;
end;

procedure TMainForm.G3DefaultClick(Sender: TObject);
begin
  Sgrid3.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  G3ChoosePicture.Checked:=False;
  G3NoBack.Checked:=False;
  G3Default.Checked:=True;
  SGrid3.Repaint;
end;

procedure TMainForm.G5DefaultClick(Sender: TObject);
begin
  Sgrid5.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  G5ChoosePicture.Checked:=False;
  G5NoBack.Checked:=False;
  G5Default.Checked:=True;
  SGrid5.Repaint;
end;

procedure TMainForm.G1ChoosePictureClick(Sender: TObject);
begin
  if OPD1.Execute then
  begin
    if G1ShowBackImage.Checked then
      SGrid1.BackGround.Bitmap.LoadFromFile(OPD1.FileName);
    GridImagePath[1]:=OPD1.FileName;
    G1ChoosePicture.Checked:=True;
    G1Default.Checked:=False;
    G1NoBack.Checked:=False;
    SGrid1.Repaint;
  end;
end;

procedure TMainForm.G2ChoosePictureClick(Sender: TObject);
begin
  if OPD1.Execute then
  begin
    if G2ShowBackImage.Checked then
      SGrid2.BackGround.Bitmap.LoadFromFile(OPD1.FileName);
    GridImagePath[2]:=OPD1.FileName;
    G2ChoosePicture.Checked:=True;
    G2NoBack.Checked:=False;
    G2Default.Checked:=False;
    SGrid2.Repaint;
  end;
end;

procedure TMainForm.G3ChoosePictureClick(Sender: TObject);
begin
  if OPD1.Execute then
  begin
    if G3ShowBackImage.Checked then
      SGrid3.BackGround.Bitmap.LoadFromFile(OPD1.FileName);
    GridImagePath[3]:=OPD1.FileName;
    G3ChoosePicture.Checked:=True;
    G3NoBack.Checked:=False;
    G3Default.Checked:=False;
    SGrid3.Repaint;
  end;
end;

procedure TMainForm.G5ChoosePictureClick(Sender: TObject);
begin
  if OPD1.Execute then
  begin
    if G5ShowBackImage.Checked then
      SGrid5.BackGround.Bitmap.LoadFromFile(OPD1.FileName);
    GridImagePath[5]:=OPD1.FileName;
    G5ChoosePicture.Checked:=True;
    G5NoBack.Checked:=False;
    G5Default.Checked:=False;
    SGrid5.Repaint;
  end;
end;

procedure TMainForm.G1ShowBackImageClick(Sender: TObject);
begin
  G1ImageSub.Enabled:=G1ShowBackImage.Checked;
  if G1ImageSub.Enabled then
  begin
    if G1ChoosePicture.Checked then
      SGrid1.BackGround.Bitmap.LoadFromFile(GridImagePath[1])
    else if G1Default.Checked then
      SGrid1.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  end
  else
    SGrid1.BackGround.Bitmap.Assign(nil);
  SGrid1.Repaint;
end;

procedure TMainForm.G2ShowBackImageClick(Sender: TObject);
begin
  G2ImageSub.Enabled:=G2ShowBackImage.Checked;
  if G2ImageSub.Enabled then
  begin
    if G2ChoosePicture.Checked then
      SGrid2.BackGround.Bitmap.LoadFromFile(GridImagePath[2])
    else if G2Default.Checked then
      SGrid2.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  end
  else
    SGrid2.BackGround.Bitmap.Assign(nil);
  SGrid2.Repaint;
end;

procedure TMainForm.G3ShowBackImageClick(Sender: TObject);
begin
  G3ImageSub.Enabled:=G3ShowBackImage.Checked;
  if G3ImageSub.Enabled then
  begin
    if G3ChoosePicture.Checked then
      SGrid3.BackGround.Bitmap.LoadFromFile(GridImagePath[3])
    else if G3Default.Checked then
      SGrid3.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  end
  else
    SGrid3.BackGround.Bitmap.Assign(nil);
  SGrid3.Repaint;
end;

procedure TMainForm.G5ShowBackImageClick(Sender: TObject);
begin
  G5ImageSub.Enabled:=G5ShowBackImage.Checked;
  if G5ImageSub.Enabled then
  begin
    if G5ChoosePicture.Checked then
      SGrid5.BackGround.Bitmap.LoadFromFile(GridImagePath[5])
    else if G5Default.Checked then
      SGrid5.BackGround.Bitmap:=GridBackImage.Picture.Bitmap;
  end
  else
    SGrid5.BackGround.Bitmap.Assign(nil);
  SGrid5.Repaint;
end;

procedure TMainForm.G1FontClick(Sender: TObject);
begin
  FD3.Font:=SGrid1.Font;
  if FD3.Execute then
  begin
    SGrid1.Font:=FD3.Font;
    SGrid1.Repaint;
    AutoSizeGridRows(SGrid1);
  end;
end;

procedure TMainForm.G2FontClick(Sender: TObject);
begin
  FD3.Font:=SGrid2.Font;
  if FD3.Execute then
    SGrid2.Font:=FD3.Font;
end;

procedure TMainForm.G3FontClick(Sender: TObject);
begin
  FD3.Font:=SGrid3.Font;
  if FD3.Execute then
    SGrid3.Font:=FD3.Font;
end;

procedure TMainForm.G5FontClick(Sender: TObject);
begin
  FD3.Font:=SGrid5.Font;
  if FD3.Execute then
    SGrid5.Font:=FD3.Font;
end;

procedure TMainForm.G1PrintClick(Sender: TObject);
begin
  if G1PrintSettingsD.Execute then
    SGrid1.Print;
end;

procedure TMainForm.G2PrintClick(Sender: TObject);
begin
  if G2PrintSettingsD.Execute then
    SGrid2.Print;
end;

procedure TMainForm.G3PrintClick(Sender: TObject);
begin
  if G3PrintSettingsD.Execute then
    SGrid3.Print;
end;

procedure TMainForm.G5PrintClick(Sender: TObject);
begin
  if G5PrintSettingsD.Execute then
    SGrid5.Print;
end;

procedure TMainForm.G1FindClick(Sender: TObject);
begin
  GridFindForm.Grid:=SGrid1;
  GridFindForm.Show;
end;

procedure TMainForm.G1ReplaceClick(Sender: TObject);
begin
  GridReplaceForm.Grid:=SGrid1;
  GridReplaceForm.Show;
end;

procedure TMainForm.G2FindClick(Sender: TObject);
begin
  GridFindForm.Grid:=SGrid2;
  GridFindForm.Show;
end;

procedure TMainForm.G2ReplaceClick(Sender: TObject);
begin
  GridReplaceForm.Grid:=SGrid2;
  GridReplaceForm.Show;
end;

procedure TMainForm.G3FindClick(Sender: TObject);
begin
  GridFindForm.Grid:=SGrid3;
  GridFindForm.Show;
end;

procedure TMainForm.G3ReplaceClick(Sender: TObject);
begin
  GridReplaceForm.Grid:=SGrid3;
  GridReplaceForm.Show;
end;

procedure TMainForm.G5FindClick(Sender: TObject);
begin
  GridFindForm.Grid:=SGrid5;
  GridFindForm.Show;
end;

procedure TMainForm.G5ReplaceClick(Sender: TObject);
begin
  GridReplaceForm.Grid:=SGrid5;
  GridReplaceForm.Show;
end;

procedure TMainForm.G1SDefaultClick(Sender: TObject);
begin
  SGrid1.Look:=glXP;
  if not(G1SFlat.Checked) and not(G1SClassic.Checked) then
    G1SDefault.Checked:=True;
end;

procedure TMainForm.G1SClassicClick(Sender: TObject);
begin
  SGrid1.Look:=glClassic;
  if not(G1SDefault.Checked) and not(G1SFlat.Checked) then
    G1SClassic.Checked:=True;
end;

procedure TMainForm.G1SFlatClick(Sender: TObject);
begin
  SGrid1.Look:=glSoft;
  if not(G1SDefault.Checked) and not(G1SClassic.Checked) then
    G1SFlat.Checked:=True;
end;

procedure TMainForm.G2SDefaultClick(Sender: TObject);
begin
  SGrid2.Look:=glXP;
  if not(G2SFlat.Checked) and not(G2SClassic.Checked) then
    G2SDefault.Checked:=True;
end;

procedure TMainForm.G2SClassicClick(Sender: TObject);
begin
  SGrid2.Look:=glClassic;
  if not(G2SDefault.Checked) and not(G2SFlat.Checked) then
    G2SClassic.Checked:=True;
end;

procedure TMainForm.G2SFlatClick(Sender: TObject);
begin
  SGrid2.Look:=glSoft;
  if not(G2SDefault.Checked) and not(G2SClassic.Checked) then
    G2SFlat.Checked:=True;
end;

procedure TMainForm.G3SDefaultClick(Sender: TObject);
begin
  SGrid3.Look:=glXP;
  if not(G3SFlat.Checked) and not(G3SClassic.Checked) then
    G3SDefault.Checked:=True;
end;

procedure TMainForm.G3sClassicClick(Sender: TObject);
begin
  SGrid3.Look:=glClassic;
  if not(G3SDefault.Checked) and not(G3SFlat.Checked) then
    G3SClassic.Checked:=True;
end;

procedure TMainForm.G3SFlatClick(Sender: TObject);
begin
  SGrid3.Look:=glSoft;
  if not(G3SDefault.Checked) and not(G3SClassic.Checked) then
    G3SFlat.Checked:=True;
end;

procedure TMainForm.G5SDefaultClick(Sender: TObject);
begin
  SGrid5.Look:=glXP;
  if not(G5SFlat.Checked) and not(G5SClassic.Checked) then
    G5SDefault.Checked:=True;
end;

procedure TMainForm.G5SClassicClick(Sender: TObject);
begin
  SGrid5.Look:=glClassic;
  if not(G5SDefault.Checked) and not(G5SFlat.Checked) then
    G5SClassic.Checked:=True;
end;

procedure TMainForm.G5SFlatClick(Sender: TObject);
begin
  SGrid5.Look:=glSoft;
  if not(G5SDefault.Checked) and not(G5SClassic.Checked) then
    G5SFlat.Checked:=True;
end;

procedure TMainForm.SGrid4DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if ARow=(SGrid4.RowCount-1) then
  begin
    SGrid4.Canvas.Brush.Color:=clWhite;
    SGrid4.Canvas.FillRect(Rect);
  end;
end;

procedure TMainForm.SGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if ARow=(SGrid1.RowCount-1) then
  begin
    SGrid1.Canvas.Brush.Color:=clWhite;
    SGrid1.Canvas.FillRect(Rect);
  end;
end;

procedure TMainForm.SGrid3DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if ARow=(SGrid3.RowCount-1) then
  begin
    SGrid3.Canvas.Brush.Color:=clWhite;
    SGrid3.Canvas.FillRect(Rect);
  end;
end;

procedure TMainForm.SGrid5DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if ARow=(SGrid5.RowCount-1) then
  begin
    SGrid5.Canvas.Brush.Color:=clWhite;
    SGrid5.Canvas.FillRect(Rect);
  end;
end;

procedure TMainForm.SGrid2DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if ARow=(SGrid2.RowCount-1) then
  begin
    SGrid2.Canvas.Brush.Color:=clWhite;
    SGrid2.Canvas.FillRect(Rect);
  end;
end;

procedure TMainForm.G5NoBackClick(Sender: TObject);
begin
  G5NoBack.Checked:=True;
  G5Default.Checked:=False;
  G5ChoosePicture.Checked:=False;
  SGrid5.BackGround.Bitmap.Assign(nil);
  SGrid5.Repaint;
  SGrid5.Repaint;
end;

procedure TMainForm.ApplySheetSettings(SheetID:Byte;var Sheet:SheetSettings);
begin
  if Sheet.Used then
  begin
    case SheetID of
      1:begin
          S1Check.OnClick:=nil;
          S1Check.Checked:=Sheet.Used;
          S1Check.OnClick:=S1CheckClick;
          S1TypeLabel.Caption:=TypeCombo.Items.Strings[Sheet.TypeIndex-1];
        end;
      2:begin
          S2Check.OnClick:=nil;
          S2Check.Checked:=Sheet.Used;
          S2Check.OnClick:=S2CheckClick;
          S2TypeLabel.Caption:=TypeCombo.Items.Strings[Sheet.TypeIndex-1];
        end;
      3:begin
          S3Check.OnClick:=nil;
          S3Check.Checked:=Sheet.Used;
          S3Check.OnClick:=S3CheckClick;
          S3TypeLabel.Caption:=TypeCombo.Items.Strings[Sheet.TypeIndex-1];
        end;
      4:begin
          S4Check.OnClick:=nil;
          S4Check.Checked:=Sheet.Used;
          S4Check.OnClick:=S4CheckClick;
          S4TypeLabel.Caption:=TypeCombo.Items.Strings[Sheet.TypeIndex-1];
        end;
      5:begin
          S5Check.OnClick:=nil;
          S5Check.Checked:=Sheet.Used;
          S5Check.OnClick:=S5CheckClick;
          S5TypeLabel.Caption:=TypeCombo.Items.Strings[Sheet.TypeIndex-1];
        end;
    end;
    SettingGrid.Cells[1,SheetID]:=Sheet.FieldName;
    SettingGrid.Cells[2,SheetID]:=TypeCombo.Items.Strings[Sheet.TypeIndex-1];
    case Sheet.TypeIndex of
      1:SettingGrid.Cells[3,SheetID]:=IntToStr(Sheet.DValue1);
      2:SettingGrid.Cells[3,SheetID]:=FloatToStr(Sheet.DValue2);
      3:SettingGrid.Cells[3,SheetID]:=Sheet.DValue3;
      4:SettingGrid.Cells[3,SheetID]:='['+IntToStr(Sheet.LBoundI)+','+IntToStr(Sheet.UBoundI)+']';
    end;
    SettingGrid.Cells[4,SheetID]:=StrTemp.Items.Strings[Sheet.DesIndex];
  end;
end;

procedure TMainForm.LoadGridFromList(DatList:TListBox;Grid:TStringGrid;var Index:Integer;Count:Integer);
var Col,Row:Integer; i:integer;
begin
  i:=0;
  for Row:=1 to (Grid.RowCount-2) do
    for Col:=1 to (Grid.ColCount-1) do
    begin
      Grid.Cells[Col,Row]:=DatList.Items.Strings[Index];
      Inc(Index);
      Inc(i);
      if i=Count then Exit;
    end;
end;

procedure TMainForm.ReadSheetList(SheetID:Byte;DatList:TListBox;var Index:Integer;Count:Integer);
var List:TListBox; i:Integer;
begin
  case SheetID of
    1:case Sheet1.TypeIndex of
        1:List:=IntList1;
        2:List:=DecList1;
        3:List:=StrList1;
      end;
    2:case Sheet2.TypeIndex of
        1:List:=IntList2;
        2:List:=DecList2;
        3:List:=StrList2;
      end;
    3:case Sheet3.TypeIndex of
        1:List:=IntList3;
        2:List:=DecList3;
        3:List:=StrList3;
      end;
    4:case Sheet4.TypeIndex of
        1:List:=IntList4;
        2:List:=DecList4;
        3:List:=StrList4;
      end;
    5:case Sheet5.TypeIndex of
        1:List:=IntList5;
        2:List:=DecList5;
        3:List:=StrList5;
      end;
  end;
  List.Items.Clear;
  for i:=1 to Count do
  begin
    List.Items.Append(DatList.Items.Strings[Index]);
    Inc(Index);
  end;
end;

procedure TMainForm.LoadWork(DatList:TListBox;const Rec:array of WorkSettings);
var S:String; Index,i:Integer;
begin
  Index:=0;
  for i:=1 to StrTemp.Items.Count do
  begin
    StrTemp.Items.Strings[i-1]:=DatList.Items.Strings[Index];
    Inc(Index);
  end;
  S:=DatList.Items.Strings[Index];
  SpanList.Items.Clear;
  Inc(Index);
  for i:=1 to StrToInt(S) do
  begin
    SpanList.Items.Append(DatList.Items.Strings[Index]);
    Inc(Index);
  end;
  if Sheet1.Used and Sheet1.ForceValue then
    ReadSheetList(1,DatList,Index,Rec[0].ForceCount);
  if Sheet2.Used and Sheet2.ForceValue then
    ReadSheetList(2,DatList,Index,Rec[1].ForceCount);
  if Sheet3.Used and Sheet3.ForceValue then
    ReadSheetList(3,DatList,Index,Rec[2].ForceCount);
  if Sheet4.Used and Sheet4.ForceValue then
    ReadSheetList(4,DatList,Index,Rec[3].ForceCount);
  if Sheet5.Used and Sheet5.ForceValue then
    ReadSheetList(5,DatList,Index,Rec[4].ForceCount);
  if Sheet1.Used then
    LoadGridFromList(DatList,SGrid1,Index,Rec[0].ValueCount);
  if Sheet2.Used then
    LoadGridFromList(DatList,SGrid2,Index,Rec[1].ValueCount);
  if Sheet3.Used then
    LoadGridFromList(DatList,SGrid3,Index,Rec[2].ValueCount);
  if Sheet4.Used then
    LoadGridFromList(DatList,SGrid4,Index,Rec[3].ValueCount);
  if Sheet5.Used then
    LoadGridFromList(DatList,SGrid5,Index,Rec[4].ValueCount);

  ApplySheetSettings(1,Sheet1);
  ApplySheetSettings(2,Sheet2);
  ApplySheetSettings(3,Sheet3);
  ApplySheetSettings(4,Sheet4);
  ApplySheetSettings(5,Sheet5);
  if Sheet1.Used then
    LoadSheet(1)
  else if Sheet2.Used then
    LoadSheet(2)
  else if Sheet3.Used then
    LoadSheet(3)
  else if Sheet4.Used then
    LoadSheet(4)
  else
    LoadSheet(5)
{  SettingGrid.Col:=1;
  SettingGrid.Row:=1;}
//  ForceValue.SetFocus;
end;

procedure TMainForm.OpenWorkProject1Click(Sender: TObject);
var WF:WorkFile;
    FName,S:String;
    i:Integer;
begin
  ODExt.Filter:='StatMan Work Files(*.wrk)|*.wrk|All Files|*.*';
  ODExt.FilterIndex:=1;
  if ODExt.Execute then
  begin
    FName:=ODExt.FileName;
    if LowerCase(RightStr(FName,4))<>'.wrk' then
    begin
      ShowMessage('Invalid work file name.');
      Exit;
    end;
    FName:=LeftStr(FName,Length(FName)-4);
    if Changed then
    begin
      i:=MessageDlg('The current project you were working on is not saved.'+Chr(13)+'Do you want to save it now?',mtConfirmation,[mbYes,mbNo,mbCancel],0);
      if i=mrCancel then
        Exit
      else
        if i=mrYes then
        begin
          SaveWorkProject1.Click;
          Changed:=False;
        end;
      end;
    try
      Changed:=False;
      OpenWorkProject(FName);

//      LoadWork(FileList,Rec);
//      SetLength(Rec,0);
      IconSheet1.Enabled:=S1Check.Checked;
      ToolSheet1.Enabled:=S1Check.Checked;
      Sheet1Section.Enabled:=S1Check.Checked;
      IconSheet2.Enabled:=S2Check.Checked;
      ToolSheet2.Enabled:=S2Check.Checked;
      Sheet2Section.Enabled:=S2Check.Checked;
      IconSheet3.Enabled:=S3Check.Checked;
      ToolSheet3.Enabled:=S3Check.Checked;
      Sheet3Section.Enabled:=S3Check.Checked;
      IconSheet4.Enabled:=S4Check.Checked;
      ToolSheet4.Enabled:=S4Check.Checked;
      Sheet4Section.Enabled:=S4Check.Checked;
      IconSheet5.Enabled:=S5Check.Checked;
      ToolSheet5.Enabled:=S5Check.Checked;
      Sheet5Section.Enabled:=S5Check.Checked;
    except
      ShowMessage('Occured an error while opening the specified work project.');
    end;
  end;
end;


procedure TMainForm.File1Click(Sender: TObject);
var B:Boolean;
begin
  if (ActiveSheet<1) or (ActiveSheet>5) then
    B:=False
  else
    B:=True;
  LoadDataSheet1.Enabled:=B;
  AppendDataSheet1.Enabled:=B;
  SaveDataSheetAs1.Enabled:=B;
  LoadDataSheetXLS.Enabled:=B;  
end;

function TMainForm.GetSGrid(No: Byte): TAdvStringGrid;
begin
  case No of
    1: Result:=SGrid1;
    2: Result:=SGrid2;
    3: Result:=SGrid3;
    4: Result:=SGrid4;
    5: Result:=SGrid5;
  end;
end;

function TMainForm.GetXLSFileName(var FName:String):Boolean;
var Ext:String;
    R:Integer;
begin
  Result:=False;
  SDExt.Filter:='MS Excel Spread Sheet(*.xls)|*.xls|All Files|*.*';
  if SDExt.Execute then
  begin
    Result:=True;
    FName:=SDExt.FileName;
    Ext:=RightStr(FName,4);
    if LowerCase(Ext)<>'.xls' then
      FName:=FName+'.xls';
    if FileExists(FName) then
    begin
      R:=MessageDlg('A file named "'+FName+'" already exists in this location.'+Chr(13)+'Do you want to replace it?',mtConfirmation,[mbYes,mbNo],0);
      if R=mrYes then
      begin
        if not(DeleteFile(FName)) then
        begin
          ShowMessage('Can not write in the specified file name.');
          Result:=False;
        end;
      end
      else
        Result:=False;
    end;
  end;
end;

procedure TMainForm.MicrosoftExcell1Click(Sender: TObject);
var S:String;
    SG:TAdvStringGrid;
    i,j:Integer;
begin
  if GetXLSFileName(S) then
  begin
    try
      SG:=GetSGrid(ActiveSheet);
      TempAdvGrid.RowCount:=SG.RowCount-2;
      TempAdvGrid.ColCount:=SG.ColCount-1;
      for i:=1 to (SG.RowCount-2) do
        for j:=1 to (SG.ColCount-1) do
          TempAdvGrid.Cells[j-1,i-1]:=SG.Cells[j,i];
      TempAdvGrid.SaveToXLS(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.MicrosoftExcel2Click(Sender: TObject);
var S:String;
    SG:TAdvStringGrid;
    i,j:Integer;
begin
  XLSSheetName.OKBtn.Caption:='Save...';
  if Length(XLSSheetName.SheetName.Text)>0 then
    XLSSheetName.OKBtn.Enabled:=True
  else
    XLSSheetName.OKBtn.Enabled:=False;
  if XLSSheetName.ShowModal=mrOk then
    if GetXLSFileName(S) then
    begin
      try
        SG:=GetSGrid(ActiveSheet);
        TempAdvGrid.RowCount:=SG.RowCount-2;
        TempAdvGrid.ColCount:=SG.ColCount-1;
        for i:=1 to (SG.RowCount-2) do
          for j:=1 to (SG.ColCount-1) do
            TempAdvGrid.Cells[j-1,i-1]:=SG.Cells[j,i];
        TempAdvGrid.SaveToXLSSheet(S,XLSSheetName.SheetName.Text);
      except
        on E:Exception do
          ShowMessage(E.Message);
      end;
    end;
end;

function TMainForm.GetWordDOCFileName(var FName: String): Boolean;
var Ext:String;
    R:Integer;
begin
  Result:=False;
  SDExt.Filter:='MS Word Documents(*.doc)|*.doc|All Files|*.*';
  if SDExt.Execute then
  begin
    Result:=True;
    FName:=SDExt.FileName;
    Ext:=RightStr(FName,4);
    if LowerCase(Ext)<>'.doc' then
      FName:=FName+'.doc';
    if FileExists(FName) then
    begin
      R:=MessageDlg('A file named "'+FName+'" already exists in this location.'+Chr(13)+'Do you want to replace it?',mtConfirmation,[mbYes,mbNo],0);
      if R=mrYes then
      begin
        if not(DeleteFile(FName)) then
        begin
          ShowMessage('Can not write in the specified file name.');
          Result:=False;
        end;
      end
      else
        Result:=False;
    end;
  end;
end;

function TMainForm.GetHTMLFileName(var FName: String): Boolean;
var Ext1,Ext2:String;
    R:Integer;
begin
  Result:=False;
  SDExt.Filter:='HTML Files|*.html;*.htm|All Files|*.*';
  if SDExt.Execute then
  begin
    Result:=True;
    FName:=SDExt.FileName;
    Ext1:=RightStr(FName,5);
    Ext2:=RightStr(FName,4);
    if (LowerCase(Ext1)<>'.html') and (LowerCase(Ext2)<>'.htm') then
      FName:=FName+'.html';
    if FileExists(FName) then
    begin
      R:=MessageDlg('A file named "'+FName+'" already exists in this location.'+Chr(13)+'Do you want to replace it?',mtConfirmation,[mbYes,mbNo],0);
      if R=mrYes then
      begin
        if not(DeleteFile(FName)) then
        begin
          ShowMessage('Can not write in the specified file name.');
          Result:=False;
        end;
      end
      else
        Result:=False;
    end;
  end;
end;

procedure TMainForm.WordDocument1Click(Sender: TObject);
var S:String;
    SG:TAdvStringGrid;
    i,j:Integer;
begin
  if GetWordDOCFileName(S) then
  begin
    try
      SG:=GetSGrid(ActiveSheet);
      TempAdvGrid.RowCount:=SG.RowCount-2;
      TempAdvGrid.ColCount:=SG.ColCount-1;
      for i:=1 to (SG.RowCount-2) do
        for j:=1 to (SG.ColCount-1) do
          TempAdvGrid.Cells[j-1,i-1]:=SG.Cells[j,i];
      TempAdvGrid.SaveToDOC(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.HTMLFile1Click(Sender: TObject);
var S:String;
    SG:TAdvStringGrid;
    i,j:Integer;
begin
  if GetHTMLFileName(S) then
  begin
    try
      SG:=GetSGrid(ActiveSheet);
      TempAdvGrid.RowCount:=SG.RowCount-2;
      TempAdvGrid.ColCount:=SG.ColCount-1;
      for i:=1 to (SG.RowCount-2) do
        for j:=1 to (SG.ColCount-1) do
          TempAdvGrid.Cells[j-1,i-1]:=SG.Cells[j,i];
      TempAdvGrid.SaveToHTML(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.TextFile1Click(Sender: TObject);
var S:String;
    R:Integer;
    SG:TAdvStringGrid;
    i,j:Integer;
begin
  SDExt.Filter:='Text Files(*.txt)|*.txt|All Files(*.*)|*.*';
  SDExt.FilterIndex:=0;
  if SDExt.Execute then
  begin
    S:=SDExt.FileName;
    if FileExists(S) then
    begin
      R:=MessageDlg('A file named "'+S+'" already exists in this location.'+Chr(13)+'Do you want to replace it?',mtConfirmation,[mbYes,mbNo],0);
      if R=mrYes then
      begin
        if not(DeleteFile(S)) then
        begin
          ShowMessage('Can not write in the specified file name.');
          Exit;
        end;
      end
      else
        Exit;
    end;
    if SDExt.FilterIndex=0 then
      if LowerCase(RightStr(S,4))<>'.txt' then
        S:=S+'.txt';
    try
      SG:=GetSGrid(ActiveSheet);
      TempAdvGrid.RowCount:=SG.RowCount-2;
      TempAdvGrid.ColCount:=SG.ColCount-1;
      for i:=1 to (SG.RowCount-2) do
        for j:=1 to (SG.ColCount-1) do
          TempAdvGrid.Cells[j-1,i-1]:=SG.Cells[j,i];
      TempAdvGrid.SaveToASCII(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.LoadDataSheetXLSClick(Sender: TObject);
var S:String;
begin
  XLSSheetName.OKBtn.Caption:='Load...';
  XLSSheetName.NoSheetName.Visible:=True;
  XLSSheetName.NoSheetName.Checked:=False;
  if Length(XLSSheetName.SheetName.Text)>0 then
    XLSSheetName.OKBtn.Enabled:=True
  else
    XLSSheetName.OKBtn.Enabled:=False;
  if XLSSheetName.ShowModal=mrOk then
  begin
    ODExt.Filter:='MS Excel Spread Sheet(*.xls)|*.xls|All Files|*.*';
    ODExt.Title:='Load Data Sheet From MS Excel File';
    if ODExt.Execute then
    begin
      try
        if XLSSheetName.NoSheetName.Checked then
          GetSGrid(ActiveSheet).LoadFromXLS(ODExt.FileName)
        else
          GetSGrid(ActiveSheet).LoadFromXLSSheet(ODExt.FileName,XLSSheetName.SheetName.Text);
      except
        XLSSheetName.NoSheetName.Checked:=False;
        XLSSheetName.NoSheetName.Visible:=False;
        if Length(XLSSheetName.SheetName.Text)=0 then
          XLSSheetName.SheetName.Text:='Sheet 1';
        Screen.Cursor:=crDefault;
        ShowMessage('Unable to save in the specified file.');
        Exit;
      end;
    end;
  end;
  XLSSheetName.NoSheetName.Checked:=False;
  XLSSheetName.NoSheetName.Visible:=False;
  if Length(XLSSheetName.SheetName.Text)=0 then
    XLSSheetName.SheetName.Text:='Sheet 1';
end;

procedure TMainForm.SaveWorkProject1Click(Sender: TObject);
begin
  try
    if WorkFileName<>'' then
    begin
      SaveWorkToFile(WorkFileName);
      Changed:=False;
    end  
    else
      SaveWorkAs.Click;
  except
    ShowMessage('Occured an error while saving work project.');
  end;      
end;

procedure TMainForm.SaveWorkAsClick(Sender: TObject);
var FName:String;
begin
  SDExt.Filter:='StatMan Work Files(*.wrk)|*.wrk|All Files|*.*';
  SDExt.Title:='Save Work Project As';
  if SDExt.Execute then
  begin
    FName:=SDExt.FileName;
    try
      SaveWorkToFile(FName);
      Changed:=False;
    except
      on E:EInOutError do
        ShowMessage('Occured an error while saving:'+E.Message)
      else
        Exit;
    end;
    WorkFileName:=FName;
  end;
end;

procedure TMainForm.WriteSheetList(SheetID:Byte;List:TListBox);
begin
  List.Items.Clear;
  case SheetID of
    1:case Sheet1.TypeIndex of
        1:List.Items:=IntList1.Items;
        2:List.Items:=DecList1.Items;
        3:List.Items:=StrList1.Items;
      end;
    2:case Sheet2.TypeIndex of
        1:List.Items:=IntList2.Items;
        2:List.Items:=DecList2.Items;
        3:List.Items:=StrList2.Items;
      end;
    3:case Sheet3.TypeIndex of
        1:List.Items:=IntList3.Items;
        2:List.Items:=DecList3.Items;
        3:List.Items:=StrList3.Items;
      end;
    4:case Sheet4.TypeIndex of
        1:List.Items:=IntList4.Items;
        2:List.Items:=DecList4.Items;
        3:List.Items:=StrList4.Items;
      end;
    5:case Sheet5.TypeIndex of
        1:List.Items:=IntList5.Items;
        2:List.Items:=DecList5.Items;
        3:List.Items:=StrList5.Items;
      end;
  end;
end;

procedure TMainForm.AppendGridToList(List:TListBox;Grid:TStringGrid;var Count:Integer);
var i,j:Integer;
begin
  Count:=0;
  for i:=1 to (Grid.RowCount-2) do
    for j:=1 to (Grid.ColCount-1) do
      if Length(Grid.Cells[j,i])>0 then
      begin
        List.Items.Append(Grid.Cells[j,i]);
        Inc(Count);
      end;
end;

procedure TMainForm.SaveWorkToFile(const FileName:String);
var WF:WorkFile;
    Rec:array of WorkSettings;
    Count:Integer;
    i:Integer;
    FName:String;
begin
  FName:=FileName;
  if LowerCase(RightStr(FName,4))='.wrk' then
    FName:=Copy(FName,1,Length(FName)-4);
  SetLength(Rec,5);
  Rec[0].Sheet:=Sheet1;
  Rec[1].Sheet:=Sheet2;
  Rec[2].Sheet:=Sheet3;
  Rec[3].Sheet:=Sheet4;
  Rec[4].Sheet:=Sheet5;
  Rec[0].StrTempCount:=StrTemp.Items.Count;
  Rec[0].SpanCount:=SpanList.Items.Count;

  Rec[0].IntCount:=IntList1.Items.Count;
  Rec[0].DecCount:=DecList1.Items.Count;
  Rec[0].StrCount:=StrList1.Items.Count;
  Rec[1].IntCount:=IntList2.Items.Count;
  Rec[1].DecCount:=DecList2.Items.Count;
  Rec[1].StrCount:=StrList2.Items.Count;
  Rec[2].IntCount:=IntList3.Items.Count;
  Rec[2].DecCount:=DecList3.Items.Count;
  Rec[2].StrCount:=StrList3.Items.Count;
  Rec[3].IntCount:=IntList4.Items.Count;
  Rec[3].DecCount:=DecList4.Items.Count;
  Rec[3].StrCount:=StrList4.Items.Count;
  Rec[4].IntCount:=IntList5.Items.Count;
  Rec[4].DecCount:=DecList5.Items.Count;
  Rec[4].StrCount:=StrList5.Items.Count;
  FileList.Items.Clear;
  FileList.Items.AddStrings(StrTemp.Items);
  FileList.Items.AddStrings(SpanList.Items);

  FileList.Items.AddStrings(IntList1.Items);
  FileList.Items.AddStrings(DecList1.Items);
  FileList.Items.AddStrings(StrList1.Items);
  FileList.Items.AddStrings(IntList2.Items);
  FileList.Items.AddStrings(DecList2.Items);
  FileList.Items.AddStrings(StrList2.Items);
  FileList.Items.AddStrings(IntList3.Items);
  FileList.Items.AddStrings(DecList3.Items);
  FileList.Items.AddStrings(StrList3.Items);
  FileList.Items.AddStrings(IntList4.Items);
  FileList.Items.AddStrings(DecList4.Items);
  FileList.Items.AddStrings(StrList4.Items);
  FileList.Items.AddStrings(IntList5.Items);
  FileList.Items.AddStrings(DecList5.Items);
  FileList.Items.AddStrings(StrList5.Items);
{  for i:=0 to (StrTemp.Items.Count-1) do
    FileList.Items.Append(StrTemp.Items.Strings[i]);
  for i:=0 to (SpanList.Items.Count-1) do
    FileList.Items.Append(SpanList.Items.Strings[i]);}
//  FValueList.Items.Clear;
{  if Sheet1.Used and Sheet1.ForceValue then
    WriteSheetList(1,FValueList);
  for i:=0 to (FValueList.Count-1) do
    FileList.Items.Append(FValueList.Items.Strings[i]);}
//  Rec[0].ForceCount:=FValueList.Items.Count;
{
  FValueList.Items.Clear;
  if Sheet2.Used and Sheet2.ForceValue then
    WriteSheetList(2,FValueList);
  for i:=0 to (FValueList.Count-1) do
    FileList.Items.Append(FValueList.Items.Strings[i]);
  Rec[1].ForceCount:=FValueList.Items.Count;

  FValueList.Items.Clear;
  if Sheet3.Used and Sheet3.ForceValue then
    WriteSheetList(3,FValueList);
  for i:=0 to (FValueList.Count-1) do
    FileList.Items.Append(FValueList.Items.Strings[i]);
  Rec[2].ForceCount:=FValueList.Items.Count;

  FValueList.Items.Clear;
  if Sheet4.Used and Sheet4.ForceValue then
    WriteSheetList(4,FValueList);
  for i:=0 to (FValueList.Count-1) do
    FileList.Items.Append(FValueList.Items.Strings[i]);
  Rec[3].ForceCount:=FValueList.Items.Count;

  FValueList.Items.Clear;
  if Sheet5.Used and Sheet5.ForceValue then
    WriteSheetList(5,FValueList);
  for i:=0 to (FValueList.Count-1) do
    FileList.Items.Append(FValueList.Items.Strings[i]);
  Rec[4].ForceCount:=FValueList.Items.Count;}

    AppendGridToList(FileList,SGrid1,Count);
    Rec[0].ValueCount:=Count;
    AppendGridToList(FileList,SGrid2,Count);
    Rec[1].ValueCount:=Count;
    AppendGridToList(FileList,SGrid3,Count);
    Rec[2].ValueCount:=Count;
    AppendGridToList(FileList,SGrid4,Count);
    Rec[3].ValueCount:=Count;
    AppendGridToList(FileList,SGrid5,Count);
    Rec[4].ValueCount:=Count;

  if FileExists(FName+'.dat') then
    DeleteFile(FName+'.dat');
  FileList.Items.SaveToFile(FName+'.dat');
  AssignFile(WF,FName+'.wrk');
  if FileExists(FName+'.wrk') then
    DeleteFile(FName+'.wrk');
  {$I-}
  Rewrite(WF);
  {$I+}
  if IOResult<>0 then
  begin
    ShowMessage('Occured an error while writing in the specified file');
    Exit;
  end;
  for i:=0 to 4 do
    Write(WF,Rec[i]);
  CloseFile(WF);
  SetLength(Rec,0);
end;

procedure TMainForm.LoadDataSheet1Click(Sender: TObject);
begin
  Load1.Click;
end;

procedure TMainForm.ApplyChartGridClick(Sender: TObject);
var i:Integer;
begin
  for i:=1 to (ChartGrid.RowCount-1) do
    if Length(ChartGrid.Cells[2,i])=0 then
      ChartGrid.Cells[2,i]:=FloatToStr(Series1.YValues[i-1]);
  ResetChart(Chart1);
  ResetChart(ChartForm.Chart1);
  for i:=1 to (ChartGrid.RowCount-1) do
  begin
    Series1.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    Series2.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    Series3.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    Series4.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    Series5.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    Series6.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    Series7.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series1.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series2.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series3.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series4.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series5.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series6.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
    ChartForm.Series7.Add(StrToInt(ChartGrid.Cells[2,i]),ChartGrid.Cells[1,i],ChartColors[i-1]);
  end;
end;

procedure TMainForm.ChartGridKeyPress(Sender: TObject; var Key: Char);
begin
  if ChartGrid.Col=2 then
    if not((Key in ['0'..'9','.']) or (Ord(Key)=Ord(VK_BACK))) then
      Key:=Chr(0);
end;

procedure TMainForm.MoveColorSelector(SGrid: TStringGrid;
  ColorSelector: TColorSelector;ACol,ARow:Integer);
var i:Integer;
begin
  ColorSelector.Visible:=False;
  ColorSelector.Left:=SGrid.Left;
  if ACol>0 then
    ColorSelector.Left:=ColorSelector.Left+SGrid.ColWidths[0];
  if ACol>1 then
    for i:=(SGrid.LeftCol+1) to (ACol-1) do
      ColorSelector.Left:=ColorSelector.Left+SGrid.ColWidths[i]+SGrid.GridLineWidth;
  ColorSelector.Top:=SGrid.Top;
  ColorSelector.Top:=ColorSelector.Top+SGrid.RowHeights[SGrid.TopRow];
  for i:=SGrid.TopRow to (ARow-1) do
    ColorSelector.Top:=ColorSelector.Top+SGrid.RowHeights[i]+SGrid.GridLineWidth;
  ColorSelector.Width:=SGrid.ColWidths[ACol]+1;
  ColorSelector.Height:=SGrid.RowHeights[ARow]+1;
  ColorSelector.Visible:=True;
  ColorSelector.BringToFront;
end;

procedure TMainForm.ChartGridColorSelectorChangeColor(Sender: TObject);
begin
  ChartColors[ColorGrid.Row-1]:=ChartGridColorSelector.Color;
  ColorGrid.Repaint;
  ChartGridColorSelector.Repaint;
end;

procedure TMainForm.ChartGridRowMoved(Sender: TObject; FromIndex,
  ToIndex: Integer);
var Temp:TColor;
begin
  Temp:=ChartColors[FromIndex-1];
  ChartColors[FromIndex-1]:=ChartColors[ToIndex-1];
  ChartColors[ToIndex-1]:=Temp;
end;

procedure TMainForm.xpButton7Click(Sender: TObject);
var i:Integer;
begin
  ChartGrid.RowCount:=TableGrid.RowCount-1;
  for i:=1 to (ChartGrid.RowCount-1) do
    ChartGrid.Cells[2,i]:=TableGrid.Cells[1,i];
end;

procedure TMainForm.xpButton6Click(Sender: TObject);
var i:Integer;
begin
  ChartGrid.RowCount:=TableGrid.RowCount-1;
  for i:=1 to (ChartGrid.RowCount-1) do
    ChartGrid.Cells[1,i]:=TableGrid.Cells[0,i];
end;

procedure TMainForm.ShowChartMarksClick(Sender: TObject);
var i:Integer;
begin
  for i:=0 to (Chart1.SeriesCount-1) do
  begin
    Chart1.Series[i].Marks.Visible:=ShowChartMarks.Checked;
    ChartForm.Chart1.Series[i].Marks.Visible:=ShowChartMarks.Checked;
  end;  
end;

procedure TMainForm.RadioButton2Click(Sender: TObject);
var Style:TSeriesMarksStyle;
    i:Integer;
begin
  case (Sender as TRadioButton).Tag of
    0: Style:=smsValue;
    1: Style:=smsPercent;
    2: Style:=smsLabel;
    3: Style:=smsLabelValue;
    4: Style:=smsLabelPercent;
    5: Style:=smsPercentTotal;
    6: Style:=smsLabelPercentTotal;
  end;
  for i:=0 to (Chart1.SeriesCount-1) do
  begin
    Chart1.Series[i].Marks.Style:=Style;
    ChartForm.Chart1.Series[i].Marks.Style:=Style;
  end;
end;

procedure TMainForm.ColorGridTopLeftChanged(Sender: TObject);
begin
  ChartGridColorSelector.Visible:=False;
  ChartGrid.TopRow:=ColorGrid.TopRow;
end;

procedure TMainForm.ChartGridTopLeftChanged(Sender: TObject);
begin
  ChartGridColorSelector.Visible:=False;
  ColorGrid.TopRow:=ChartGrid.TopRow;
end;

procedure TMainForm.ColorGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var ARect:TRect;
begin
  if ARow>0 then
  begin
    ARect:=Rect;
    ColorGrid.Canvas.Brush.Color:=ChartColors[ARow-1];
    ColorGrid.Canvas.FillRect(Classes.Rect(ARect.Left+2,ARect.Top+2,ARect.Right-2,ARect.Bottom-2));
  end;
end;

procedure TMainForm.ColorGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var i:Integer;
begin
  if ARow>0 then
  begin
    ChartGridColorSelector.Visible:=False;
    ChartGridColorSelector.Left:=ColorGrid.Left+2;
    ChartGridColorSelector.Top:=ColorGrid.Top+ColorGrid.RowHeights[0]+2;
    if ARow<>ColorGrid.TopRow then
      for i:=(ColorGrid.TopRow+1) to ARow do
        ChartGridColorSelector.Top:=ChartGridColorSelector.Top+ColorGrid.RowHeights[i];
    ChartGridColorSelector.Color:=ChartColors[ARow-1];
    ChartGridColorSelector.Width:=ColorGrid.ColWidths[0];
    ChartGridColorSelector.Visible:=True;
    ChartGridColorSelector.BringToFront;
  end;
end;

procedure TMainForm.ColorGridExit(Sender: TObject);
begin
  ChartGridColorSelector.Visible:=False;
end;

procedure TMainForm.ChartGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  ColorGrid.Row:=ARow;
  ChartGridColorSelector.Visible:=False;
end;

procedure TMainForm.ColorGridEnter(Sender: TObject);
begin
  ChartGridColorSelector.Visible:=True;
end;

procedure TMainForm.FastLabelClick(Sender: TObject);
begin
  FastCalcAllMenu.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
end;

procedure TMainForm.FastLabelMouseEnter(Sender: TObject);
begin
  (Sender as TLabel).Color:=clMenuHighlight;
end;

procedure TMainForm.FastLabelMouseLeave(Sender: TObject);
begin
  (Sender as TLabel).Color:=clGray;
end;

procedure TMainForm.FastCalc1Click(Sender: TObject);
begin
  FastCalcToolbar.Visible:=(Sender as TMenuItem).Checked;
end;

procedure TMainForm.CalcResultbtnClick(Sender: TObject);
var F:Extended; S:String;
begin
  if Length(FastExpression.Text)=0 then
  begin
    FastResult.Caption:='{No Expression}';
    FastResult.Hint:='There is no expression in the expression box.';
    Exit;
  end;
  try
    FastCalcParser.Expression:=FastExpression.Text;
  except
    on ESyntaxError do
    begin
      FastResult.Caption:='{Syntax Error}';
      FastResult.Hint:='Syntax error in the expression.'+Chr(13)+'Please check the expression syntax such as parentheses order.';
      Exit;
    end;
    on EInvalidOp do
    begin
      FastResult.Caption:='{Invalid Operation}';
      FastResult.Hint:='The expression caused an invalid operation while '+Chr(13)+'calculating the result(Such as "Division By Zero").';
      Exit;
    end;
    else
    begin
      FastResult.Caption:='{Error}';
      FastResult.Hint:='Occured an error while calculating the result';
      Exit;
    end;
  end;
  F:=FastCalcParser.Value;
  S:=FloatToStr(F-Frac(F))+Copy(FloatToStr(Frac(F)),2,Length(FloatToStr(Frac(F)))-1);
  if (Pos('E',FloatToStr(F))>0) or (Pos('e',FloatToStr(F))>0) then
    S:=FloatToStr(F);
  FastResult.Caption:=S;
  FastResult.Hint:='Result:'+Chr(13)+S;
end;

procedure TMainForm.SpeedButton26Click(Sender: TObject);
begin
  FastExpression.Text:='';
end;

procedure TMainForm.SpeedButton3Click(Sender: TObject);
begin
  FastExpression.Text:=FastExpression.Text+(Sender as TSpeedButton).Caption;
end;

procedure TMainForm.SpeedButton23Click(Sender: TObject);
begin
  FastExpression.Text:=FastExpression.Text+(Sender as TSpeedButton).Caption+'(';
end;

procedure TMainForm.SpeedButton33Click(Sender: TObject);
begin
  if Length(FastExpression.Text)>0 then
    FastExpression.Text:=Copy(FastExpression.Text,1,Length(FastExpression.Text)-1);
end;

procedure TMainForm.Calculate1Click(Sender: TObject);
begin
  CalcResultbtn.Click;
end;

procedure TMainForm.CopyToClipboard1Click(Sender: TObject);
var n:Integer;
    V:Extended;
begin
  CalcResultbtn.Click;
  Val(FastResult.Caption,V,n);
  if n<=0 then
    Clipboard.AsText:=FastResult.Caption;
end;

procedure TMainForm.FastExpressionKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    CalcResultbtn.Click;
    FastExpression.SelectAll;
    FastExpression.SetFocus;
  end
  else if Ord(Key)=Ord(VK_ESCAPE) then
  begin
    Key:=Chr(0);
    FastExpression.Text:='';
  end
  else if Key=' ' then
    Key:=Chr(0);
end;

procedure TMainForm.BitBtn3Click(Sender: TObject);
var i,j:Integer;
    a:Double;
begin
  for i:=1 to 8 do
  begin
    a:=Random(600)+76.6546685;
    for j:=0 to (Chart1.SeriesCount-1) do
      Chart1.Series[j].Add(a);
  end;    
end;

procedure TMainForm.PrinterSetup1Click(Sender: TObject);
begin
  PSD1.Execute;
end;

procedure TMainForm.RandomNumberProducer1Click(Sender: TObject);
var A:Integer;
begin
  A:=ActiveSheet;
  ActiveSheet:=10;
  RandomProducerForm.ShowModal;
  ActiveSheet:=A;
end;

procedure TMainForm.SGrid4MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button=mbRight then
  begin
    GridMouseUp(SGrid4,GridPopup4,GridCellPopup,X,Y);
  end;
end;

procedure TMainForm.GridMouseUp(Grid: TStringGrid; P1, P2: TPopupMenu; X,
  Y: Integer);
var B:Boolean;
    Row,Col:Integer;
begin
  B:=False;
  for Col:=0 to (Grid.ColCount-1) do
    if (X>Grid.CellRect(Col,0).Left) and (X<Grid.CellRect(Col,0).Right)
      and (Y>Grid.CellRect(Col,0).Top) and (Y<Grid.CellRect(Col,0).Bottom) then
      Break;
  if Col<Grid.ColCount then
  begin
    P1.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    Exit;
  end;
  for Row:=1 to (Grid.RowCount-2) do
    if (X>Grid.CellRect(Row,0).Left) and (X<Grid.CellRect(Row,0).Right)
      and (Y>Grid.CellRect(Row,0).Top) and (Y<Grid.CellRect(Row,0).Bottom) then
      Break;
  if Row<(Grid.RowCount-1) then
  begin
    P1.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    Exit;
  end;
  for Row:=1 to (Grid.RowCount-2) do
  begin
    for Col:=1 to (Grid.ColCount-1) do
      if (X>Grid.CellRect(Col,Row).Left) and (X<Grid.CellRect(Col,Row).Right)
        and (Y>Grid.CellRect(Col,Row).Top) and (Y<Grid.CellRect(Col,Row).Bottom) then
        begin
          B:=True;
          Break;
        end;
    if B then
      Break;
  end;
  if B then
  begin
    if (Grid.Col<>Col) or (Grid.Row<>Row) then
    begin
      Grid.Row:=Row;
      Grid.Col:=Col;
    end;  
    P2.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
  end;
end;

procedure TMainForm.CutCellClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  Clipboard.AsText:=SG.Cells[SG.Col,SG.Row];
  SG.Cells[SG.Col,SG.Row]:='';
end;

procedure TMainForm.CopyCellClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  Clipboard.AsText:=SG.Cells[SG.Col,SG.Row];
end;

procedure TMainForm.PasteCellClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  if Clipboard.HasFormat(CF_TEXT) then
  begin
    SG:=GetSGrid(ActiveSheet);
    SG.Cells[SG.Col,SG.Row]:=Clipboard.AsText;
  end;
end;

procedure TMainForm.GridCellPopupPopup(Sender: TObject);
var SG:TAdvStringGrid;
    S:String;
    Sheet:SheetSettings;
    i,j,Row,Col:Integer;
    B:Boolean;
begin
  case ActiveSheet of
    1: if S1Combo.Visible then S1Combo.Visible:=False;
    2: if S2Combo.Visible then S2Combo.Visible:=False;
    3: if S3Combo.Visible then S3Combo.Visible:=False;
    4: if S4Combo.Visible then S4Combo.Visible:=False;
    5: if S5Combo.Visible then S5Combo.Visible:=False;
  end;
  SG:=GetSGrid(ActiveSheet);
  if Clipboard.HasFormat(CF_TEXT) then
    PasteCell.Enabled:=True
  else
    PasteCell.Enabled:=False;
  Sheet:=GetSheet(ActiveSheet);
  FilterCell.Enabled:=Sheet.Filter;
  FilterCellValue.Enabled:=Sheet.Filter;
  if Length(SG.Cells[SG.Col,SG.Row])>0 then
  begin
    SearchCell.Enabled:=True;
    S:='"'+SG.Cells[SG.Col,SG.Row]+'"';
    if Length(S)>17 then
      S:=LeftStr(S,14)+'...';
    SearchCell.Caption:='Search For '+S;
    ResizeCell.Enabled:=True;
  end
  else
  begin
    SearchCell.Caption:='Search For "{Serach String}"';
    SearchCell.Enabled:=False;
    ResizeCell.Enabled:=False;
    FilterCellValue.Enabled:=False;
  end;
  if Clipboard.HasFormat(CF_TEXT) then
    PasteCell.Enabled:=True
  else
    PasteCell.Enabled:=False;
  if Length(LastFilter[ActiveSheet].Str)>0 then
    UndoCellFilter.Enabled:=True
  else
    UndoCellFilter.Enabled:=False;
  if High(LockedCells[ActiveSheet])>=0 then
    UnlockAllCell.Enabled:=True
  else
    UnlockAllCell.Enabled:=False;
  if UnlockAllCell.Enabled then
  begin
    B:=True;
    for i:=0 to High(LockedCells[ActiveSheet]) do
      if (LockedCells[ActiveSheet][i].Row=SG.Row)
        and (LockedCells[ActiveSheet][i].Col=SG.Col) then
        begin
          B:=False;
          Break;
        end;
    if B then
    begin
      LockUnlockCell.Caption:='Lock Cell Content';
      LockUnlockCell.Tag:=0;
    end
    else
    begin
      LockUnlockCell.Caption:='Unlock Cell Content';
      LockUnlockCell.Tag:=1;
    end;
  end;
end;

procedure TMainForm.DeleteCellClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  SG.Cells[SG.Col,SG.Row]:='';
end;

procedure TMainForm.ResizeCellClick(Sender: TObject);
var SG:TAdvStringGrid;
    i,MaxW:Integer;
begin
  SG:=GetSGrid(ActiveSheet);
  MaxW:=0;
  for i:=0 to (SG.RowCount-2) do
    if SG.Canvas.TextWidth(SG.Cells[SG.Col,i])>MaxW then
      MaxW:=SG.Canvas.TextWidth(SG.Cells[SG.Col,i]);
  SG.ColWidths[SG.Col]:=9+MaxW;
end;

procedure TMainForm.ColorCellClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  ColorD1.Color:=SG.Colors[SG.Col,SG.Row];
  if ColorD1.Execute then
    SG.Colors[SG.Col,SG.Row]:=ColorD1.Color;
end;

procedure TMainForm.FreqCellClick(Sender: TObject);
var i,j,Count:Integer;
    SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  Count:=0;
  for i:=1 to (SG.RowCount-2) do
    for j:=1 to (SG.ColCount-1) do
      if SG.Cells[j,i]=SG.Cells[SG.Col,SG.Row] then
        Inc(Count);
  FreqCellValue.Caption:='&'+IntToStr(Count);      
end;

procedure TMainForm.SearchCellClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  GridFindForm.TextCombo.Text:=SG.Cells[SG.Col,SG.Row];
  GridFindForm.Show;
end;

procedure TMainForm.FilterCellValueClick(Sender: TObject);
var SG:TAdvStringGrid;
    Sheet:SheetSettings;
    CanAdd:Boolean;
    S,SubS:String;
    Precision:Byte;
    K:Integer;
begin
  SG:=GetSGrid(ActiveSheet);
  Sheet:=GetSheet(ActiveSheet);
  case ActiveSheet of
    1: Precision:=S1PSpin.Value;
    2: Precision:=S2PSpin.Value;
    3: Precision:=S3PSpin.Value;
    4: Precision:=S4PSpin.Value;
    5: Precision:=S5PSpin.Value;
  end;
  FilterParser.Expression:=Copy(StrTemp.Items.Strings[Sheet.FilterIndex],6,Length(StrTemp.Items.Strings[Sheet.FilterIndex])-5);
  CanAdd:=True;
  S:=SG.Cells[SG.Col,SG.Row];
  with LastFilter[ActiveSheet] do
  begin
    Str:=S;
    Row:=SG.Row;
    Col:=SG.Col;
  end;
  if Sheet.TypeIndex=2 then
    FilterParser.X:=StrToFloat(S)
  else if Sheet.TypeIndex=1 then
    FilterParser.X:=StrToInt(S);
  SetRoundMode(rmDown);
  S:=FloatToStr(Round(FilterParser.Value))+Copy(FloatToStr(Frac(FilterParser.Value)),2,Length(FloatToStr(Frac(FilterParser.Value)))-1);
  if (Pos('E',S)>0) or (Pos('e',S)>0) then
    CanAdd:=False;
  if CanAdd then
  begin
    SubS:='';
    if Pos('.',S)>0 then
    begin
      SubS:=Copy(S,Pos('.',S)+1,Length(S)-Pos('.',S));
      S:=Copy(S,1,Pos('.',S)-1);
      if Length(SubS)>Precision then
        SubS:=Copy(Subs,1,Precision);
    end;
    if Precision=0 then
      SG.Cells[SG.Col,SG.Row]:=S
    else
    begin
      if Sheet.TypeIndex=2 then
        S:=S+'.'+SubS
      else
      begin
        for k:=1 to Length(SubS) do
          if SubS[k]<>'0' then
            CanAdd:=False;
      end;
      if CanAdd then
        SG.Cells[SG.Col,SG.Row]:=S;
    end;
    if CanAdd then
      RebuildFreqTable:=True;
  end;
end;

function TMainForm.GetSheet(No: Byte): SheetSettings;
begin
  case No of
    1: Result:=Sheet1;
    2: Result:=Sheet2;
    3: Result:=Sheet3;
    4: Result:=Sheet4;
    5: Result:=Sheet5;
  end;
end;

procedure TMainForm.UndoCellFilterClick(Sender: TObject);
var SG:TAdvStringGrid;
begin
  SG:=GetSGrid(ActiveSheet);
  with LastFilter[ActiveSheet] do
  begin
    SG.Cells[Col,Row]:=Str;
    Str:='';
  end;  
end;

procedure TMainForm.SGrid1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button=mbRight then
  begin
    GridMouseUp(SGrid1,GridPopup1,GridCellPopup,X,Y);
  end;
end;

procedure TMainForm.SGrid3MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button=mbRight then
  begin
    GridMouseUp(SGrid3,GridPopup3,GridCellPopup,X,Y);
  end;
end;

procedure TMainForm.SGrid5MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button=mbRight then
  begin
    GridMouseUp(SGrid5,GridPopup5,GridCellPopup,X,Y);
  end;
end;

procedure TMainForm.SGrid2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button=mbRight then
  begin
    GridMouseUp(SGrid2,GridPopup2,GridCellPopup,X,Y);
  end;
end;

procedure TMainForm.LockUnlockCellClick(Sender: TObject);
var SG:TAdvStringGrid;
    i,j:Integer;
begin
  SG:=GetSGrid(ActiveSheet);
  if LockUnlockCell.Tag=0 then
  begin
    SetLength(LockedCells[ActiveSheet],High(LockedCells[ActiveSheet])+2);
    with LockedCells[ActiveSheet][High(LockedCells[ActiveSheet])] do
    begin
      Row:=SG.Row;
      Col:=SG.Col;
    end;
  end else if LockUnlockCell.Tag=1 then
  begin
    for i:=0 to High(LockedCells[ActiveSheet]) do
      with LockedCells[ActiveSheet][i] do
        if (Row=SG.Row) and (Col=SG.Col) then
          Break;
    for j:=(i+1) to High(LockedCells[ActiveSheet]) do
      LockedCells[ActiveSheet][j-1]:=LockedCells[ActiveSheet][j];
    SetLength(LockedCells[ActiveSheet],
      High(LockedCells[ActiveSheet]));
  end;
end;

procedure TMainForm.UnlockAllCellClick(Sender: TObject);
begin
  SetLength(LockedCells[ActiveSheet],0);
end;

procedure TMainForm.SGrid4CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
var i:Integer;
begin
//  if not(Sheet4.ForceValue) then
    CanEdit:=CanEditCell(ACol,ARow,ActiveSheet);
//  else
//    CanEdit:=False;
end;

function TMainForm.CanEditCell(ACol, ARow: Integer; ASheet: Byte): Boolean;
var i:Integer;
begin
  Result:=True;
  if High(LockedCells[ASheet])>=0 then
    for i:=0 to High(LockedCells[ASheet]) do
      with(LockedCells[ASheet][i]) do
        if (Row=ARow) and (Col=ACol) then
        begin
          Result:=False;
          Exit;
        end;
end;

procedure TMainForm.SGrid1CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  CanEdit:=CanEditCell(ACol,ARow,ActiveSheet);
end;

procedure TMainForm.SGrid3CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  CanEdit:=CanEditCell(ACol,ARow,ActiveSheet);
end;

procedure TMainForm.SGrid5CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  CanEdit:=CanEditCell(ACol,ARow,ActiveSheet);
end;

procedure TMainForm.SGrid2CanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  CanEdit:=CanEditCell(ACol,ARow,ActiveSheet);
end;

procedure TMainForm.TipOfTheDayClick(Sender: TObject);
begin
  TipForm.ShowModal;
end;

procedure TMainForm.FormulaEditor1Click(Sender: TObject);
begin
  FormulaForm.FmlEdit.Text:='F(X)=X';
  FormulaForm.OKbtn.Enabled:=True;
  FormulaForm.OKbtn.Cancel:=True;
  FormulaForm.Cancelbtn.Visible:=False;
  FormulaForm.Edit1.Text:='';
  FormulaForm.REdit.Text:='0';
  FormulaForm.ShowModal;
  FormulaForm.OKbtn.Cancel:=False;
end;

procedure TMainForm.HistoryLabel1MouseEnter(Sender: TObject);
begin
  HistoryLabel1.Color:=clBtnFace;
  HistoryLabel1.Font.Color:=clWindowText;
  HistoryLabel1.Tag:=0;
  HistoryLabel2.Color:=clBtnFace;
  HistoryLabel2.Font.Color:=clWindowText;
  HistoryLabel2.Tag:=0;
  HistoryLabel3.Color:=clBtnFace;
  HistoryLabel3.Font.Color:=clWindowText;
  HistoryLabel3.Tag:=0;
  HistoryLabel4.Color:=clBtnFace;
  HistoryLabel4.Font.Color:=clWindowText;
  HistoryLabel4.Tag:=0;
  HistoryLabel5.Color:=clBtnFace;
  HistoryLabel5.Font.Color:=clWindowText;
  HistoryLabel5.Tag:=0;
  (Sender as TLabel).Color:=clMenuHighlight;
  (Sender as TLabel).Font.Color:=clMenuText;
  (Sender as TLabel).Tag:=1;
end;

procedure TMainForm.HistoryLabel1MouseLeave(Sender: TObject);
begin
  HistoryLabel1.Color:=clBtnFace;
  HistoryLabel1.Font.Color:=clWindowText;
  HistoryLabel2.Color:=clBtnFace;
  HistoryLabel2.Font.Color:=clWindowText;
  HistoryLabel3.Color:=clBtnFace;
  HistoryLabel3.Font.Color:=clWindowText;
  HistoryLabel4.Color:=clBtnFace;
  HistoryLabel4.Font.Color:=clWindowText;
  HistoryLabel5.Color:=clBtnFace;
  HistoryLabel5.Font.Color:=clWindowText;
end;

procedure TMainForm.HistoryLabel1Click(Sender: TObject);
begin
  ChartHistoryPopup.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
end;

procedure TMainForm.HistoryBtnDropDownClick(Sender: TObject);
begin
  if ChartHistoryCount<5 then
  begin
    HistoryLabel5.Enabled:=False;
    if ChartHistoryCount<4 then
    begin
      HistoryLabel4.Enabled:=False;
      if ChartHistoryCount<3 then
      begin
        HistoryLabel3.Enabled:=False;
        if ChartHistoryCount<2 then
        begin
          HistoryLabel2.Enabled:=False;
          if ChartHistoryCount<1 then
            HistoryLabel1.Enabled:=False;
        end;
      end;
    end;
  end;
  if ChartHistoryCount>0 then
    HistoryLabel1.Enabled:=True;
  if ChartHistoryCount>1 then
    HistoryLabel2.Enabled:=True;
  if ChartHistoryCount>2 then
    HistoryLabel3.Enabled:=True;
  if ChartHistoryCount>3 then
    HistoryLabel4.Enabled:=True;
  if ChartHistoryCount>4 then
    HistoryLabel5.Enabled:=True;
end;

procedure TMainForm.Load2Click(Sender: TObject);
var Chart:TChart;
    i,j:Integer;
begin
  if HistoryLabel1.Tag=1 then
    Chart:=HistoryChart1
  else if HistoryLabel2.Tag=1 then
    Chart:=HistoryChart2
  else if HistoryLabel3.Tag=1 then
    Chart:=HistoryChart3
  else if HistoryLabel4.Tag=1 then
    Chart:=HistoryChart4
  else if HistoryLabel5.Tag=1 then
    Chart:=HistoryChart5;
  for i:=0 to (Chart1.SeriesCount-1) do
    for j:=1 to Chart1.Series[i].Count do
    begin
      Chart1.Series[i].Delete(0);
      ChartForm.Chart1.Series[i].Delete(0);
    end;
  for i:=0 to (Chart.Series[0].Count-1) do
    for j:=0 to (Chart1.SeriesCount-1) do
    begin
      Chart1.Series[j].Add(Chart.Series[0].YValue[i]);
      ChartForm.Chart1.Series[j].Add(Chart.Series[0].YValue[i]);
    end;
end;

procedure TMainForm.Append1Click(Sender: TObject);
var Chart:TChart;
    i,j:Integer;
begin
  if HistoryLabel1.Tag=1 then
    Chart:=HistoryChart1
  else if HistoryLabel2.Tag=1 then
    Chart:=HistoryChart2
  else if HistoryLabel3.Tag=1 then
    Chart:=HistoryChart3
  else if HistoryLabel4.Tag=1 then
    Chart:=HistoryChart4
  else if HistoryLabel5.Tag=1 then
    Chart:=HistoryChart5;
  for i:=0 to (Chart.Series[0].Count-1) do
    for j:=0 to (Chart1.SeriesCount-1) do
    begin
      Chart1.Series[j].Add(Chart.Series[0].YValue[i]);
      ChartForm.Chart1.Series[j].Add(Chart.Series[0].YValue[i]);
    end;
end;

procedure TMainForm.Preview1Click(Sender: TObject);
var Chart:TChart;
    i,j:Integer;
begin
  if HistoryLabel1.Tag=1 then
    Chart:=HistoryChart1
  else if HistoryLabel2.Tag=1 then
    Chart:=HistoryChart2
  else if HistoryLabel3.Tag=1 then
    Chart:=HistoryChart3
  else if HistoryLabel4.Tag=1 then
    Chart:=HistoryChart4
  else if HistoryLabel5.Tag=1 then
    Chart:=HistoryChart5;
  for i:=0 to (ChartPrevForm.Chart1.SeriesCount-1) do
    for j:=1 to ChartPrevForm.Chart1.Series[i].Count do
      ChartPrevForm.Chart1.Series[i].Delete(0);
  for i:=0 to (Chart.Series[0].Count-1) do
    for j:=0 to (ChartPrevForm.Chart1.SeriesCount-1) do
      ChartPrevForm.Chart1.Series[j].Add(Chart.Series[0].YValue[i]);
  ChartPrevForm.Show;    
end;

procedure TMainForm.GridPopup1Popup(Sender: TObject);
begin
  if Sheet1.TypeIndex=3 then
    ViewAsChart1.Enabled:=True
  else
    ViewAsChart1.Enabled:=False;
end;

procedure TMainForm.ChartItemDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
  ACanvas.Brush.Color:=clBtnFace;
  ACanvas.FillRect(ARect);
  ACanvas.Draw(ARect.Left+4,ARect.Top,ViewAsChartImage.Picture.Bitmap);
end;

procedure TMainForm.ChartItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
  var Width, Height: Integer);
begin
  Width:=ViewAsChartImage.Picture.Bitmap.Width;
  Height:=ViewAsChartImage.Picture.Bitmap.Height;
end;

procedure TMainForm.AsPicture1Click(Sender: TObject);
begin
  TempChart.CopyToClipboardBitmap;
end;

procedure TMainForm.AsMetafile1Click(Sender: TObject);
begin
  TempChart.CopyToClipboardMetafile(False);
end;

procedure TMainForm.AsPicture2Click(Sender: TObject);
var S:String;
begin
  SDExt.Filter:='Bitmap Files(*.bmp)|*.bmp|All Files(*.*)|*.*';
  SDExt.FilterIndex:=0;
  if SDExt.Execute then
  begin
    S:=SDExt.FileName;
    if LowerCase(RightStr(S,4))<>'.bmp' then
      S:=S+'.bmp';
    TempChart.SaveToBitmapFile(S);
  end;
end;

procedure TMainForm.AsMetafile2Click(Sender: TObject);
var S:String;
begin
  SDExt.Filter:='Metafiles(*.wmf)|*.wmf|All Files(*.*)|*.*';
  SDExt.FilterIndex:=0;
  if SDExt.Execute then
  begin
    S:=SDExt.FileName;
    if LowerCase(RightStr(S,4))<>'.wmf' then
      S:=S+'.wmf';
    TempChart.SaveToMetafile(S);
  end;
end;

procedure TMainForm.PrintViewChartClick(Sender: TObject);
begin
  if PSD1.Execute then
    TempChart.Print;
end;

procedure TMainForm.SendToChartViewer1Click(Sender: TObject);
var i,j:Integer;
begin
  ChartPrevForm.Hide;
  for i:=0 to (ChartPrevForm.Chart1.SeriesCount-1) do
    for j:=1 to ChartPrevForm.Chart1.Series[i].Count do
      ChartPrevForm.Chart1.Series[i].Delete(0);
  for i:=0 to (TempChart.Series[0].Count-1) do
    for j:=0 to (ChartPrevForm.Chart1.SeriesCount-1) do
      ChartPrevForm.Chart1.Series[j].Add(TempChart.Series[0].YValue[i],TempChart.Series[0].XLabel[i],TempChart.Series[0].ValueColor[i]);
  ChartPrevForm.Show;
end;

procedure TMainForm.DesItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
  var Width, Height: Integer);
begin
  Width:=ViewAsChartImage.Picture.Bitmap.Width;
  Height:=ACanvas.TextHeight('{Chart View}')+4;
end;

procedure TMainForm.DesItemAdvancedDrawItem(Sender: TObject;
  ACanvas: TCanvas; ARect: TRect; State: TOwnerDrawState);
begin
  ACanvas.Brush.Color:=clMenu;
  ACanvas.FillRect(ARect);
  ACanvas.Pen.Color:=clRed;
  ACanvas.Rectangle(ARect);
  ACanvas.Font.Color:=clBlue;
  ACanvas.TextOut(ARect.Left+1,ARect.Top+1,'    {Chart View Of Data Sheet 1}');
end;

procedure TMainForm.GridPopup2Popup(Sender: TObject);
begin
  if Sheet2.TypeIndex=3 then
    ViewAsChart2.Enabled:=True
  else
    ViewAsChart2.Enabled:=False;
end;

procedure TMainForm.GridPopup3Popup(Sender: TObject);
begin
  if Sheet3.TypeIndex=3 then
    ViewAsChart3.Enabled:=True
  else
    ViewAsChart3.Enabled:=False;
end;

procedure TMainForm.GridPopup4Popup(Sender: TObject);
begin
  if Sheet4.TypeIndex=3 then
    ViewAsChart4.Enabled:=True
  else
    ViewAsChart4.Enabled:=False;  
end;

procedure TMainForm.GridPopup5Popup(Sender: TObject);
begin
  if Sheet5.TypeIndex=3 then
    ViewAsChart5.Enabled:=True
  else
    ViewAsChart5.Enabled:=False;
end;

procedure TMainForm.ViewAsChart(SG: TAdvStringGrid;ViewItem:TMenuItem);
var i,j,k,Count:Integer;
    S:String;
    StrValues:array of String;
    Add:Boolean;
begin
  for i:=1 to (SG.RowCount-2) do
    for j:=1 to (SG.ColCount-1) do
      if Length(SG.Cells[j,i])>0 then
      begin
        S:=SG.Cells[j,i];
        Add:=True;
        if High(StrValues)>=0 then
          for k:=0 to High(StrValues) do
            if StrValues[k]=S then
            begin
              Add:=False;
              Break;
            end;
        if Add then
        begin
          SetLength(StrValues,High(StrValues)+2);
          StrValues[High(StrValues)]:=S;
        end;
      end;
  if not(High(StrValues)>=0) then
  begin
    ViewItem.Enabled:=False;
    Exit;
  end;
  for i:=1 to TempChart.Series[0].Count do
    TempChart.Series[0].Delete(0);
  for i:=0 to High(StrValues) do
  begin
    Count:=0;
    S:=StrValues[i];
    for j:=1 to (SG.RowCount-2) do
      for k:=1 to (SG.ColCount-1) do
        if S=SG.Cells[k,j] then
          Inc(Count);
    TempChart.Series[0].Add(Count,S);
  end;
  TempChart.CopyToClipboardBitmap;
  ViewAsChartImage.Picture.Bitmap.Assign(Clipboard);
  SetLength(StrValues,0);
end;

procedure TMainForm.ViewAsChart1Click(Sender: TObject);
begin
  ViewAsChart(SGrid1,ViewAsChart1);
end;

procedure TMainForm.ViewAsChart2Click(Sender: TObject);
begin
  ViewAsChart(SGrid2,ViewAsChart2);
end;

procedure TMainForm.ViewAsChart3Click(Sender: TObject);
begin
  ViewAsChart(SGrid3,ViewAsChart3);
end;

procedure TMainForm.ViewAsChart4Click(Sender: TObject);
begin
  ViewAsChart(SGrid4,ViewAsChart4);
end;

procedure TMainForm.ViewAsChart5Click(Sender: TObject);
begin
  ViewAsChart(SGrid5,ViewAsChart5);
end;

procedure TMainForm.DesItem2AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; State: TOwnerDrawState);
begin
  ACanvas.Brush.Color:=clMenu;
  ACanvas.FillRect(ARect);
  ACanvas.Pen.Color:=clRed;
  ACanvas.Rectangle(ARect);
  ACanvas.Font.Color:=clBlue;
  ACanvas.TextOut(ARect.Left+1,ARect.Top+1,'    {Chart View Of Data Sheet 2}');
end;

procedure TMainForm.DesItem3AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; State: TOwnerDrawState);
begin
  ACanvas.Brush.Color:=clMenu;
  ACanvas.FillRect(ARect);
  ACanvas.Pen.Color:=clRed;
  ACanvas.Rectangle(ARect);
  ACanvas.Font.Color:=clBlue;
  ACanvas.TextOut(ARect.Left+1,ARect.Top+1,'    {Chart View Of Data Sheet 3}');
end;

procedure TMainForm.DesItem4AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; State: TOwnerDrawState);
begin
  ACanvas.Brush.Color:=clMenu;
  ACanvas.FillRect(ARect);
  ACanvas.Pen.Color:=clRed;
  ACanvas.Rectangle(ARect);
  ACanvas.Font.Color:=clBlue;
  ACanvas.TextOut(ARect.Left+1,ARect.Top+1,'    {Chart View Of Data Sheet 4}');
end;

procedure TMainForm.DesItem5AdvancedDrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; State: TOwnerDrawState);
begin
  ACanvas.Brush.Color:=clMenu;
  ACanvas.FillRect(ARect);
  ACanvas.Pen.Color:=clRed;
  ACanvas.Rectangle(ARect);
  ACanvas.Font.Color:=clBlue;
  ACanvas.TextOut(ARect.Left+1,ARect.Top+1,'    {Chart View Of Data Sheet 5}');
end;

procedure TMainForm.Chart4Click(Sender: TObject);
begin
  PrintChart.Click;
end;

procedure TMainForm.CutToolbtnClick(Sender: TObject);
begin
  CutCell.Click;
end;

procedure TMainForm.CopyToolbtnClick(Sender: TObject);
begin
  CopyCell.Click;
end;

procedure TMainForm.PasteToolbtnClick(Sender: TObject);
begin
  PasteCell.Click;
end;

procedure TMainForm.FindToolbtnClick(Sender: TObject);
begin
  if (ActiveSheet>0) and (ActiveSheet<6) then
  begin
    GridFindForm.Grid:=GetSGrid(ActiveSheet);
    GridFindForm.Show;
  end;  
end;

procedure TMainForm.ToolButton28Click(Sender: TObject);
begin
  if EnglishHelp.Checked then
    StatManHelp1.Click
  else If PersianHelp.Checked then
    StatManHelp2.Click;   
end;

procedure TMainForm.ToolButton24Click(Sender: TObject);
begin
  NewWorkProject1.Click;
end;

procedure TMainForm.ToolButton25Click(Sender: TObject);
begin
  OpenWorkProject1.Click;
end;

procedure TMainForm.ToolButton23Click(Sender: TObject);
begin
  SaveWorkProject1.Click;
end;

procedure TMainForm.GlobalToolbar1Click(Sender: TObject);
begin
  GlobalToolbar.Visible:=(Sender as TMenuItem).Checked;  
end;

procedure TMainForm.Print2Click(Sender: TObject);
begin
  if ChartHistoryCount>0 then
    PrintChart.Enabled:=True
  else
    PrintChart.Enabled:=False;
  if (ActiveSheet>0) and (ActiveSheet<6) then
    PrintDataSheet.Enabled:=True
  else
    PrintDataSheet.Enabled:=False;
end;

procedure TMainForm.PrintDataSheetClick(Sender: TObject);
begin
  case ActiveSheet of
    1: G1Print.Click;
    2: G2Print.Click;
    3: G3Print.Click;
    4: G4Print.Click;
    5: G5Print.Click;
  end;
end;

procedure TMainForm.PrintChartClick(Sender: TObject);
begin
  PrintBtn.Click;
end;

procedure TMainForm.PrintToolbtnClick(Sender: TObject);
begin
  if ChartHistoryCount>0 then
    Chart4.Enabled:=True
  else
    Chart4.Enabled:=False;
  if (ActiveSheet>0) and (ActiveSheet<6) then
    DataSheet2.Enabled:=True
  else
    DataSheet2.Enabled:=False;
end;

procedure TMainForm.ResizeForm;
begin
  if MemberGroup.Visible  and (MemberGroup.Parent=MainForm) then
  begin
    MemberGroup.Left:=3;
    MemberGroup.Top:=ControlBar1.Top+ControlBar1.Height+10;
  end;
  if MemberGroup.Visible and (MemberGroup.Parent=MainForm) then
    SettingPanel.Left:=MemberGroup.Left+MemberGroup.Width+20
  else
    SettingPanel.Left:=5;
  SettingPanel.Top:=ControlBar1.Top+ControlBar1.Height+5;
  Sheet1Panel.Left:=SettingPanel.Left;
  Sheet2Panel.Left:=SettingPanel.Left;
  Sheet3Panel.Left:=SettingPanel.Left;
  Sheet4Panel.Left:=SettingPanel.Left;
  Sheet5Panel.Left:=SettingPanel.Left;
  TablePanel.Left:=SettingPanel.Left;
  ChartPanel.Left:=SettingPanel.Left;
  AnalyzePanel.Left:=SettingPanel.Left;
  Sheet1Panel.Top:=SettingPanel.Top;
  Sheet2Panel.Top:=SettingPanel.Top;
  Sheet3Panel.Top:=SettingPanel.Top;
  Sheet4Panel.Top:=SettingPanel.Top;
  Sheet5Panel.Top:=SettingPanel.Top;
  TablePanel.Top:=SettingPanel.Top;
  ChartPanel.Top:=SettingPanel.Top;
  AnalyzePanel.Top:=SettingPanel.Top;
  SettingPanel.Width:=609;
  SettingPanel.Height:=419;
  if MemberGroup.Visible and (MemberGroup.Parent=MainForm) then
    MainForm.Width:=MemberGroup.Width+5+SettingPanel.Width+10+20
  else
    MainForm.Width:=5+SettingPanel.Width+10+20;
  MainForm.Height:=ControlBar1.Height+5+SettingPanel.Height+20+95//Round(StatusControlBar.Height*2.9)
end;

procedure TMainForm.ControlBar1Resize(Sender: TObject);
begin
  ResizeForm;
end;

procedure TMainForm.FormResize(Sender: TObject);
var DeltaW,DeltaH:Integer;
begin
  ResizeForm;
//  StatusBar.Repaint;
//  FastCalcToolbar.Repaint;
end;

procedure TMainForm.G1NoBackClick(Sender: TObject);
begin
  G1NoBack.Checked:=True;
  G1Default.Checked:=False;
  G1ChoosePicture.Checked:=False;
  SGrid1.BackGround.Bitmap.Assign(nil);
  SGrid1.Repaint;
  SGrid1.Repaint;
end;

procedure TMainForm.G2NoBackClick(Sender: TObject);
begin
  G2NoBack.Checked:=True;
  G2Default.Checked:=False;
  G2ChoosePicture.Checked:=False;
  SGrid2.BackGround.Bitmap.Assign(nil);
  SGrid2.Repaint;
  SGrid2.Repaint;
end;

procedure TMainForm.G3NoBackClick(Sender: TObject);
begin
  G3NoBack.Checked:=True;
  G3Default.Checked:=False;
  G3ChoosePicture.Checked:=False;
  SGrid3.BackGround.Bitmap.Assign(nil);
  SGrid3.Repaint;
  SGrid3.Repaint;
end;

procedure TMainForm.G4NoBackClick(Sender: TObject);
begin
  G4NoBack.Checked:=True;
  G4Default.Checked:=False;
  G4ChoosePicture.Checked:=False;
  SGrid4.BackGround.Bitmap.Assign(nil);
  SGrid4.Repaint;
  SGrid4.Repaint;
end;

procedure TMainForm.FreqCellValueAdvancedDrawItem(Sender: TObject;
  ACanvas: TCanvas; ARect: TRect; State: TOwnerDrawState);
begin
  ACanvas.Brush.Color:=clBtnFace;
  ACanvas.FillRect(ARect);
  ACanvas.Pen.Color:=clGreen;
  ACanvas.Pen.Width:=2;
  ACanvas.Rectangle(ARect);
  ACanvas.Font.Color:=clBlack;
  ACanvas.TextOut(ARect.Left+4,ARect.Top+1,Copy(FreqCellValue.Caption,2,Length(
    FreqCellValue.Caption)-1));
end;

procedure TMainForm.FreqCellValueMeasureItem(Sender: TObject;
  ACanvas: TCanvas; var Width, Height: Integer);
begin
  Width:=ACanvas.TextWidth('8')*(Length(FreqCellValue.Caption)-1);
  Height:=ACanvas.TextHeight('8')+4;
end;

procedure TMainForm.DettachAttachPaneClick(Sender: TObject);
begin
  if MemberGroup.Parent=MainForm then
  begin
    DettachAttachPane.Caption:='Attach Member Pane';
    MemberForm.Show;
    MemberGroup.Parent:=MemberForm;
    MemberGroup.Left:=1;
    MemberGroup.Top:=1;
    MemberPane1.Enabled:=False;
    MemberPane2.Enabled:=False;
    MemberButtonNormal.Visible:=False;
    MemberButtonPressed.Visible:=False;
    MemberGroup.PopupMenu:=MemberForm.DettachPopup;
  end
  else
  begin
    DettachAttachPane.Caption:='Dettach Member Pane';
    MemberForm.Hide;
    MemberGroup.Parent:=MainForm;
    MemberGroup.Left:=3;
    MemberGroup.Top:=50;
    MemberPane1.Enabled:=True;
    MemberPane2.Enabled:=True;
    MemberButtonNormal.Visible:=True;
    MemberButtonPressed.Visible:=True;
    MemberGroup.PopupMenu:=nil;
  end;
  ResizeForm;
  if FastCalcToolbar.Parent=StatusBar then
  begin
    FastCalcToolbar.Left:=StatusBar.Width-FastCalcToolbar.Width-105;
    FastCalcToolbar.Top:=3;
  end;  
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  MemberGroup.Parent:=MainForm;
end;

procedure TMainForm.Edit1Click(Sender: TObject);
begin
  MCut.Enabled:=CutToolbtn.Enabled;
  MCopy.Enabled:=CopyToolbtn.Enabled;
  MPaste.Enabled:=PasteToolbtn.Enabled;
end;

procedure TMainForm.MCutClick(Sender: TObject);
begin
  CutToolbtn.Click;
end;

procedure TMainForm.MCopyClick(Sender: TObject);
begin
  CopyToolbtn.Click;
end;

procedure TMainForm.MPasteClick(Sender: TObject);
begin
  PasteToolbtn.Click;
end;

procedure TMainForm.VisibleMenuHeaders(V: Boolean);
begin
  Help1.Visible:=V;
  Tools1.Visible:=V;
  Chart3.Visible:=V;
  DataSheetMainMenu.Visible:=V;
  View1.Visible:=V;
  Edit1.Visible:=V;
  File1.Visible:=V;
end;

procedure TMainForm.GotoNextSection;
begin
  case ActiveSheet of
    0: MSheet1Section.Click;
    1: MSheet2Section.Click;
    2: MSheet3Section.Click;
    3: MSheet4Section.Click;
    4: MSheet5Section.Click;
    5: MTableSection.Click;
    6: MChartSection.Click;
    7: MAnalyzeSection.Click;
    8: MSettingSection.Click;
  end;
end;

procedure TMainForm.GotoPreviousSection;
begin
  case ActiveSheet of
    2: MSheet1Section.Click;
    3: MSheet2Section.Click;
    4: MSheet3Section.Click;
    5: MSheet4Section.Click;
    6: MSheet5Section.Click;
    7: MTableSection.Click;
    8: MChartSection.Click;
    0: MAnalyzeSection.Click;
    1: MSettingSection.Click;
  end;
end;

procedure TMainForm.N60DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
  ACanvas.Brush.Color:=clBtnFace;
  ACanvas.FillRect(ARect);
  ACanvas.Draw(ARect.Left+4,ARect.Top,ViewAsChartImage.Picture.Bitmap);
end;

procedure TMainForm.N63DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
  ACanvas.Brush.Color:=clBtnFace;
  ACanvas.FillRect(ARect);
  ACanvas.Draw(ARect.Left+4,ARect.Top,ViewAsChartImage.Picture.Bitmap);
end;

procedure TMainForm.N66DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
  ACanvas.Brush.Color:=clBtnFace;
  ACanvas.FillRect(ARect);
  ACanvas.Draw(ARect.Left+4,ARect.Top,ViewAsChartImage.Picture.Bitmap);
end;

procedure TMainForm.N69DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; Selected: Boolean);
begin
  ACanvas.Brush.Color:=clBtnFace;
  ACanvas.FillRect(ARect);
  ACanvas.Draw(ARect.Left+4,ARect.Top,ViewAsChartImage.Picture.Bitmap);
end;

procedure TMainForm.SetMenuOwnerDraw;
begin
  ChartItem.OnMeasureItem:=ChartItemMeasureItem;
  ChartItem.OnDrawItem:=ChartItemDrawItem;
  DesItem.OnMeasureItem:=DesItemMeasureItem;
  DesItem.OnAdvancedDrawItem:=DesItemAdvancedDrawItem;
  AsChart2.OnMeasureItem:=ChartItemMeasureItem;
  AsChart3.OnMeasureItem:=ChartItemMeasureItem;
  AsChart4.OnMeasureItem:=ChartItemMeasureItem;
  AsChart5.OnMeasureItem:=ChartItemMeasureItem;
  AsChart2.OnDrawItem:=ChartItemDrawItem;
  AsChart3.OnDrawItem:=ChartItemDrawItem;
  AsChart4.OnDrawItem:=ChartItemDrawItem;
  AsChart5.OnDrawItem:=ChartItemDrawItem;
  DesItem2.OnMeasureItem:=DesItem.OnMeasureItem;
  DesItem3.OnMeasureItem:=DesItem.OnMeasureItem;
  DesItem4.OnMeasureItem:=DesItem.OnMeasureItem;
  DesItem5.OnMeasureItem:=DesItem.OnMeasureItem;
  AnimItem.OnAdvancedDrawItem:=AnimItemAdvancedDrawItem;
  AnimItem.OnMeasureItem:=AnimItemMeasureItem;
end;

procedure TMainForm.AnimTimerTimer(Sender: TObject);
begin
  Inc(AnimIndex1);
  if AnimIndex1=(StatAnim1.Count+1) then
    AnimIndex1:=1;
  AnimItem.Enabled:=True;
  AnimItem.Enabled:=False;
end;

procedure TMainForm.AnimItemAdvancedDrawItem(Sender: TObject;
  ACanvas: TCanvas; ARect: TRect; State: TOwnerDrawState);
var Bitmap:TBitmap;
begin
  ACanvas.Brush.Color:=clWhite;
  ACanvas.FillRect(ARect);
  Bitmap:=TBitmap.Create;
  StatAnim1.GetBitmap(AnimIndex1-1,Bitmap);
  ACanvas.Draw(ARect.Left+5,ARect.Top,Bitmap);
  Bitmap.Free;
end;

procedure TMainForm.AutosizeTabelColumn;
var i,MaxW:Integer;
begin
  MaxW:=TableGrid.Canvas.TextWidth(TableGrid.Cells[0,0]);
  for i:=2 to (TableGrid.RowCount-1) do
    if TableGrid.Canvas.TextWidth(TableGrid.Cells[0,i])>MaxW then
      MaxW:=TableGrid.Canvas.TextWidth(TableGrid.Cells[0,i]);
  TableGrid.ColWidths[0]:=MaxW+5;    
end;

procedure TMainForm.CloseAnaReportClick(Sender: TObject);
begin
  ReportPopup.Close;
end;

procedure TMainForm.AnimItemMeasureItem(Sender: TObject; ACanvas: TCanvas;
  var Width, Height: Integer);
begin
  Width:=StatAnim1.Width;
  Height:=StatAnim1.Height;
end;

procedure TMainForm.TTIconDblClick(Sender: TObject);
begin
  if IsMinimized then
    OpenStatMan1.Click;        
end;

procedure TMainForm.TableGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  if (ACol=0) or (ARow=0) then
    TableGrid.Options:=TableGrid.Options+[goEditing]
  else
    TableGrid.Options:=TableGrid.Options-[goEditing];
end;

procedure TMainForm.FreqExportBtnDropDownClick(Sender: TObject);
var i,j:Integer;
begin
  TempAdvGrid.ColCount:=TableGrid.ColCount;
  TempAdvGrid.RowCount:=TableGrid.RowCount;
  for i:=0 to (TableGrid.RowCount-1) do
    for j:=0 to (TableGrid.ColCount-1) do
      TempAdvGrid.Cells[j,i]:=TableGrid.Cells[j,i];
end;

procedure TMainForm.ToolButton18Click(Sender: TObject);
var i,j:Integer;
begin
  TempAdvGrid.ColCount:=TableGrid.ColCount;
  TempAdvGrid.RowCount:=TableGrid.RowCount;
  for i:=0 to (TableGrid.RowCount-1) do
    for j:=0 to (TableGrid.ColCount-1) do
      TempAdvGrid.Cells[j,i]:=TableGrid.Cells[j,i];
  if TempPrintSettingsD.Execute then
    if PSD1.Execute then
      TempAdvGrid.Print;
end;

procedure TMainForm.MicrosoftExcelSpreadSheet1Click(Sender: TObject);
var S:String;
begin
  if GetXLSFileName(S) then
  begin
    try
      TempAdvGrid.SaveToXLS(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.WordDocument2Click(Sender: TObject);
var S:String;
begin
  if GetWordDOCFileName(S) then
  begin
    try
      TempAdvGrid.SaveToDOC(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.HTMLFile2Click(Sender: TObject);
var S:String;
begin
  if GetHTMLFileName(S) then
  begin
    try
      TempAdvGrid.SaveToHTML(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.extFile1Click(Sender: TObject);
var S:String;
    R:Integer;
begin
  SDExt.Filter:='Text Files(*.txt)|*.txt|All Files(*.*)|*.*';
  SDExt.FilterIndex:=0;
  if SDExt.Execute then
  begin
    S:=SDExt.FileName;
    if FileExists(S) then
    begin
      R:=MessageDlg('A file named "'+S+'" already exists in this location.'+Chr(13)+'Do you want to replace it?',mtConfirmation,[mbYes,mbNo],0);
      if R=mrYes then
      begin
        if not(DeleteFile(S)) then
        begin
          ShowMessage('Can not write in the specified file name.');
          Exit;
        end;
      end
      else
        Exit;
    end;
    if SDExt.FilterIndex=0 then
      if LowerCase(RightStr(S,4))<>'.txt' then
        S:=S+'.txt';
    try
      TempAdvGrid.SaveToASCII(S);
    except
      on E:Exception do
        ShowMessage(E.Message);
    end;
  end;
end;

procedure TMainForm.FrequencyTableToolbar1Click(Sender: TObject);
begin
  FreqToolbar.Visible:=(Sender as TMenuItem).Checked;
end;

procedure TMainForm.DettachFreqToolbarClick(Sender: TObject);
begin
  if (FreqToolbar.Parent=ControlBar1) or
     (FreqToolbar.Parent=StatusControlBar) then
  begin
    FreqToolbar.Visible:=False;
    FreqToolbar.Left:=11;
    FreqToolbar.Top:=2;
    FreqToolbar.Parent:=FreqControlBar;
    DettachFreqToolbar.Hint:='Dettach Toolbar';
    FreqToolbar.Visible:=True;
  end
  else
  begin
    FreqToolbar.Visible:=False;
    FreqToolbar.Parent:=ControlBar1;
    DettachFreqToolbar.Hint:='Attach Toolbar';
    FreqToolbar.Visible:=True;
  end;
end;

procedure TMainForm.AutoSizeGridRows(SGrid: TStringGrid);
var i,j,MaxH:Integer;
begin
  MaxH:=0;
  for i:=0 to (SGrid.RowCount-2) do
    for j:=0 to (SGrid.ColCount-1) do
      if Length(SGrid.Cells[j,i])>0 then
        if SGrid.Canvas.TextHeight(SGrid.Cells[j,i])>MaxH then
          MaxH:=SGrid.Canvas.TextHeight(SGrid.Cells[j,i]);
  MaxH:=MaxH+11;
  for i:=1 to (SGrid.RowCount-2) do
    SGrid.RowHeights[i]:=MaxH;
end;

procedure TMainForm.English2Click(Sender: TObject);
begin
   Application.HelpFile:=ExeDir+'STATMAN.HLP';
   Application.HelpContext(HelpTopic);
end;

procedure TMainForm.Persian2Click(Sender: TObject);
begin
  Application.HelpFile:=ExeDir+'STATMANPERSIAN.HLP';
  Application.HelpContext(HelpTopic);
end;

procedure TMainForm.Del1Enter(Sender: TObject);
begin
  if (List1.ItemIndex<0) or (List1.Items.Count=0) then
    Del1.Enabled:=False;
  Del1.Repaint;  
  HelpLabel.Caption:='Delete Button:'+Chr(13)+Chr(13)+'   Removes the selected value from the value list';
end;

procedure TMainForm.EditFmlEnter(Sender: TObject);
begin
  HelpLabel.Caption:='Edit Formula Button:'+Chr(13)+Chr(13)+'   Click to change the filter formula of the selected data sheet filter';
end;

procedure TMainForm.Add1Enter(Sender: TObject);
begin
  HelpLabel.Caption:='Add Button:'+Chr(13)+Chr(13)+'   Adds the new value to the value list';
end;

procedure TMainForm.OpenWorkProject(const FName: String);

  procedure LoadGridFromList(SGrid:TAdvStringGrid;AList:TListBox;
    Stpos,Count:Integer);
  var i,Col,Row:Integer;
  begin
    ClearGrid(SGrid);
    if Count=0 then Exit;
    i:=0;
    for Row:=1 to (SGrid.RowCount-2) do
      for Col:=1 to (SGrid.ColCount-1) do
      begin
        SGrid.Cells[Col,Row]:=AList.Items.Strings[StPos+i];
        Inc(i);
        if i=Count then Exit;
      end;
  end;

  procedure LoadInSheet(Sheet:SheetSettings;IntList,
    DecList,StrList:TListBox);
  var L:TListBox;
      i:Integer;
  begin
    List1.Items.Clear;
    VEdit1.Text:='';
    ForceValue.Checked:=Sheet.ForceValue;
    if Sheet.ForceValue then
    begin
      case Sheet.TypeIndex of
        1: L:=IntList;
        2: L:=DecList;
        3: L:=StrList;
        4: begin
             for i:=0 to (SpanList.Items.Count-1) do
               if StrToInt(Copy(SpanList.Items.Strings[i],1,1))=Sheet.VListIndex then
                 List1.Items.Append(RightStr(SpanList.Items.Strings[i],Length(SpanList.Items.Strings[i])-1));
           end;
      end;
      if Sheet.TypeIndex<4 then
        List1.Items.AddStrings(L.Items);
    end;
    UseFilter.Checked:=Sheet.Filter;
    if Sheet.Filter then
      FxLabel.Caption:=StrTemp.Items.Strings[Sheet.FilterIndex];
    if Sheet.TypeIndex=3 then
    begin
      StrLimitsTable.Enabled:=True;
      StrLimitsTable.Cells[1,1]:=IntToStr(Sheet.MaxStringLength);
      case Sheet.CharStyle of
        csLowercase: StrLimitsTable.Cells[1,2]:='Lowercase';
        csNone: StrLimitsTable.Cells[1,2]:='None';
        csUppercase: StrLimitsTable.Cells[1,2]:='Uppercase';
      end;
      CharStyleCombo.ItemIndex:=CharStyleCombo.Items.IndexOf(StrLimitsTable.Cells[1,2]);
    end
    else
      StrLimitsTable.Enabled:=False;
  end;

var WF:WorkFile;
    Rec:array of WorkSettings;
    R:WorkSettings;
    i,j,ListPos:Integer;
    List:TListBox;
begin
  SetLength(Rec,5);
  AssignFile(WF,FName+'.wrk');
  Reset(WF);
  for i:=0 to 4 do
  begin
    Read(WF,R);
    Rec[i]:=R;
  end;
  CloseFile(WF);

  ResetProgram;

  MSettingSection.Click;
  SettingGrid.Col:=1;
  if Rec[0].Sheet.Used then
  begin
    S1Check.Checked:=True;
    SettingGrid.Row:=1;
  end
  else if Rec[1].Sheet.Used then
  begin
    S2Check.Checked:=True;
    SettingGrid.Row:=2;
  end
  else if Rec[2].Sheet.Used then
  begin
    S3Check.Checked:=True;
    SettingGrid.Row:=3;
  end
  else if Rec[3].Sheet.Used then
  begin
    S4Check.Checked:=True;
    SettingGrid.Row:=4;
  end
  else if Rec[4].Sheet.Used then
  begin
    S5Check.Checked:=True;
    SettingGrid.Row:=5;
  end;

  Sheet1:=Rec[0].Sheet;
  Sheet2:=Rec[1].Sheet;
  Sheet3:=Rec[2].Sheet;
  Sheet4:=Rec[3].Sheet;
  Sheet5:=Rec[4].Sheet;

  FileList.Items.Clear;
  FileList.Items.LoadFromFile(FName+'.dat');
  StrTemp.Items.Clear;
  ListPos:=0;
  for ListPos:=0 to (Rec[0].StrTempCount-1) do
    StrTemp.Items.Append(FileList.Items.Strings[ListPos]);
  Dec(ListPos);
  SpanList.Items.Clear;
  for i:=1 to Rec[0].SpanCount do
  begin
    Inc(ListPos);
    SpanList.Items.Append(FileList.Items.Strings[ListPos]);
  end;
  for j:=0 to 4 do
  begin
    case j of
      0: List:=IntList1;
      1: List:=IntList2;
      2: List:=IntList3;
      3: List:=IntList4;
      4: List:=IntList5;
    end;
    List.Items.Clear;
    for i:=1 to Rec[j].IntCount do
    begin
      Inc(ListPos);
      List.Items.Append(FileList.Items.Strings[ListPos]);
    end;
    case j of
      0: List:=DecList1;
      1: List:=DecList2;
      2: List:=DecList3;
      3: List:=DecList4;
      4: List:=DecList5;
    end;
    List.Items.Clear;
    for i:=1 to Rec[j].DecCount do
    begin
      Inc(ListPos);
      List.Items.Append(FileList.Items.Strings[ListPos]);
    end;
    case j of
      0: List:=StrList1;
      1: List:=StrList2;
      2: List:=StrList3;
      3: List:=StrList4;
      4: List:=StrList5;
    end;
    List.Items.Clear;
    for i:=1 to Rec[j].StrCount do
    begin
      Inc(ListPos);
      List.Items.Append(FileList.Items.Strings[ListPos]);
    end;
  end;
  LoadGridFromList(SGrid1,FileList,ListPos+1,Rec[0].ValueCount);
  SGrid1.Tag:=10; //Means after load
  ListPos:=ListPos+Rec[0].ValueCount;
  LoadGridFromList(SGrid2,FileList,ListPos+1,Rec[1].ValueCount);
  SGrid2.Tag:=10; //Means after load
  ListPos:=ListPos+Rec[1].ValueCount;
  LoadGridFromList(SGrid3,FileList,ListPos+1,Rec[2].ValueCount);
  SGrid3.Tag:=10; //Means after load
  ListPos:=ListPos+Rec[2].ValueCount;
  LoadGridFromList(SGrid4,FileList,ListPos+1,Rec[3].ValueCount);
  SGrid4.Tag:=10; //Means after load
  ListPos:=ListPos+Rec[3].ValueCount;
  LoadGridFromList(SGrid5,FileList,ListPos+1,Rec[4].ValueCount);
  SGrid5.Tag:=10; //Means after load
  ListPos:=ListPos+Rec[4].ValueCount;

  SetLength(Rec,0);

  ApplySheetSettings(1,Sheet1);
  ApplySheetSettings(2,Sheet2);
  ApplySheetSettings(3,Sheet3);
  ApplySheetSettings(4,Sheet4);
  ApplySheetSettings(5,Sheet5);

  if S1Check.Checked then
    SettingGrid.Cells[0,1]:='      Data Sheet 1'
  else
    SettingGrid.Cells[0,1]:='      [Not Used]';
  if S2Check.Checked then
    SettingGrid.Cells[0,2]:='      Data Sheet 2'
  else
    SettingGrid.Cells[0,2]:='      [Not Used]';
  if S3Check.Checked then
    SettingGrid.Cells[0,3]:='      Data Sheet 3'
  else
    SettingGrid.Cells[0,3]:='      [Not Used]';
  if S4Check.Checked then
    SettingGrid.Cells[0,4]:='      Data Sheet 4'
  else
    SettingGrid.Cells[0,4]:='      [Not Used]';
  if S5Check.Checked then
    SettingGrid.Cells[0,5]:='      Data Sheet 5'
  else
    SettingGrid.Cells[0,5]:='      [Not Used]';

  if S1Check.Checked then
    LoadInSheet(Sheet1,IntList1,DecList1,StrList1)
  else if S2Check.Checked then
    LoadInSheet(Sheet2,IntList2,DecList2,StrList2)
  else if S3Check.Checked then
    LoadInSheet(Sheet3,IntList3,DecList3,StrList3)
  else if S4Check.Checked then
    LoadInSheet(Sheet4,IntList4,DecList4,StrList4)
  else if S5Check.Checked then
    LoadInSheet(Sheet5,IntList5,DecList5,StrList5);
end;

procedure TMainForm.SGrid4ClickCell(Sender: TObject; ARow, ACol: Integer);
begin
//    ShowMessage('CanSelect=True'+Chr(13)+'Row='+IntToStr(ARow)+Chr(13)+'Col='+IntToStr(ACol))
  if S4Combo.Visible then
    S4Combo.Visible:=False;
  if SGrid4.Row<>ARow then
    SGrid4.Tag:=5 //Don't do anything
  else
    SGrid4.Tag:=0;
  SGrid4.Col:=ACol;
  if SGrid4.Tag<>0 then
    SGrid4.Tag:=0;
  SGrid4.Row:=ARow;
end;

procedure TMainForm.SGrid1ClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if S1Combo.Visible then
    S1Combo.Visible:=False;
  if SGrid1.Row<>ARow then
    SGrid1.Tag:=5 //Don't do anything
  else
    SGrid1.Tag:=0;
  SGrid1.Col:=ACol;
  if SGrid1.Tag<>0 then
    SGrid1.Tag:=0;
  SGrid1.Row:=ARow;
end;

procedure TMainForm.SGrid3ClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if S3Combo.Visible then
    S3Combo.Visible:=False;
  if SGrid3.Row<>ARow then
    SGrid3.Tag:=5 //Don't do anything
  else
    SGrid3.Tag:=0;
  SGrid3.Col:=ACol;
  if SGrid3.Tag<>0 then
    SGrid3.Tag:=0;
  SGrid3.Row:=ARow;
end;

procedure TMainForm.SGrid5ClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if S5Combo.Visible then
    S5Combo.Visible:=False;
  if SGrid5.Row<>ARow then
    SGrid5.Tag:=5 //Don't do anything
  else
    SGrid5.Tag:=0;
  SGrid5.Col:=ACol;
  if SGrid5.Tag<>0 then
    SGrid5.Tag:=0;
  SGrid5.Row:=ARow;
end;

procedure TMainForm.SGrid2ClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if S2Combo.Visible then
    S2Combo.Visible:=False;
  if SGrid2.Row<>ARow then
    SGrid2.Tag:=5 //Don't do anything
  else
    SGrid2.Tag:=0;
  SGrid2.Col:=ACol;
  if SGrid2.Tag<>0 then
    SGrid2.Tag:=0;
  SGrid2.Row:=ARow;
end;

procedure TMainForm.SGrid1TopLeftChanged(Sender: TObject);
begin
  if S1Combo.Visible then
    S1Combo.Visible:=False;
end;

procedure TMainForm.ResetProgram;
var i:Integer;
begin
  ClearGrid(SGrid1);
  ClearGrid(SGrid2);
  ClearGrid(SGrid3);
  ClearGrid(SGrid4);
  ClearGrid(SGrid5);
  SetSheetDefaults;
  with Sheet1 do
  begin
    Used:=True;
    FieldName:='Field 1';
    TypeIndex:=1;
    DValue1:=0;
    StrTemp.Items.Strings[DesIndex]:='My field 1';
  end;
  ActiveField:=1;
  ResetChart(Chart1);
  ResetChart(ChartForm.Chart1);
  ClearGrid(TableGrid);
  for i:=1 to (TableGrid.RowCount-2) do
    TableGrid.Cells[0,i]:='';
  S1Check.Checked:=True;
  S2Check.Checked:=False;
  S3Check.Checked:=False;
  S4Check.Checked:=False;
  S5Check.Checked:=False;
  //****** S E T V A R I A B L E S D E F A U L T ******
  VListCount:=0;
  CurrentType:=1;
  ModeListVisible:=False;
  CanCreateAChart:=False;
  ActiveSheet:=0;
  SettingPanel.BringToFront;
  SettingSection.Down:=True;
  IconSettings.Down:=True;
  ToolSettings.Down:=True;
  MSettingSection.Click;
end;

procedure TMainForm.PromptAsMenuClick(Sender: TObject);
begin
  PromptAsList.Checked:=False;
  PromptAsMenu.Checked:=True;
end;

procedure TMainForm.PromptAsListClick(Sender: TObject);
begin
  PromptAsMenu.Checked:=False;
  PromptAsList.Checked:=True;
end;

procedure TMainForm.ToolbarsClick(Sender: TObject);
begin
  FastCalc1.Checked:=FastCalcToolbar.Visible;
  GlobalToolbar1.Checked:=GlobalToolbar.Visible;
  FrequencyTableToolbar1.Checked:=FreqToolbar.Visible;
end;

procedure TMainForm.GlobalToolbarResize(Sender: TObject);
begin
  GlobalToolbar.Width:=236;
  GlobalToolbar.Height:=22;
end;

procedure TMainForm.FreqToolbarResize(Sender: TObject);
begin
  FreqToolbar.Width:=208;
  FreqToolbar.Height:=27;
end;

procedure TMainForm.FreqControlBarResize(Sender: TObject);
begin
  CreateTable2.Top:=FreqControlBar.Top+FreqControlBar.Height+4;
end;

procedure TMainForm.S4EditFastKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    S4AddFastbtn.Click;
  end;
end;

procedure TMainForm.S4AddFastbtnClick(Sender: TObject);
begin
  AddFastInput(SGrid4,4,S4EditFast.Text);
  CheckGridCell(SGrid4,S4TypeLabel);
  S4EditFast.SelectAll;
  S4EditFast.SetFocus;
end;

procedure TMainForm.S4EditFastChange(Sender: TObject);
begin
  if Length(S4EditFast.Text)>0 then
    S4AddFastbtn.Enabled:=True
  else
    S4AddFastbtn.Enabled:=False;
end;

procedure TMainForm.AddFastInput(SGrid: TAdvStringGrid;
  SheetID:Byte;const Input: String);
var i,j:Integer;
    HaveLockedCell:Boolean;
begin
  HaveLockedCell:=False;
  for i:=1 to (SGrid.RowCount-2) do
    for j:=1 to (SGrid.ColCount-1) do
    begin
      if (Length(SGrid.Cells[j,i])=0) then
      begin
        if not(HaveLockedCell) then
          HaveLockedCell:=not(CanEditCell(j,i,SheetID));
        if CanEditCell(j,i,SheetID) then
        begin
          SGrid.Cells[j,i]:=Input;
          Exit;
        end;
      end;
    end;
  if HaveLockedCell then
    ShowMessage('There is no empty cell in the data sheet to add this value to it.'+
      'You can unlock a locked empty cell to add this new item in it.')
  else
    ShowMessage('There is no empty cell in the data sheet to add this value to it.');
end;

procedure TMainForm.S1EditFastChange(Sender: TObject);
begin
  if Length(S1EditFast.Text)>0 then
    S1AddFastbtn.Enabled:=True
  else
    S1AddFastbtn.Enabled:=False;
end;

procedure TMainForm.S1EditFastKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    S1AddFastbtn.Click;
  end;
end;

procedure TMainForm.S1AddFastbtnClick(Sender: TObject);
begin
  AddFastInput(SGrid1,1,S1EditFast.Text);
  CheckGridCell(SGrid1,S1TypeLabel);
  S1EditFast.SelectAll;
  S1EditFast.SetFocus;
end;

procedure TMainForm.S3EditFastChange(Sender: TObject);
begin
  if Length(S3EditFast.Text)>0 then
    S3AddFastbtn.Enabled:=True
  else
    S3AddFastbtn.Enabled:=False;
end;

procedure TMainForm.S3EditFastKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    S3AddFastbtn.Click;
  end;
end;

procedure TMainForm.S3AddFastbtnClick(Sender: TObject);
begin
  AddFastInput(SGrid3,3,S3EditFast.Text);
  CheckGridCell(SGrid3,S3TypeLabel);
  S3EditFast.SelectAll;
  S3EditFast.SetFocus;
end;

procedure TMainForm.S5EditFastChange(Sender: TObject);
begin
  if Length(S5EditFast.Text)>0 then
    S5AddFastbtn.Enabled:=True
  else
    S5AddFastbtn.Enabled:=False;
end;

procedure TMainForm.S5EditFastKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    S5AddFastbtn.Click;
  end;
end;

procedure TMainForm.S5AddFastbtnClick(Sender: TObject);
begin
  AddFastInput(SGrid5,5,S5EditFast.Text);
  CheckGridCell(SGrid5,S5TypeLabel);
  S5EditFast.SelectAll;
  S5EditFast.SetFocus;
end;

procedure TMainForm.S2EditFastChange(Sender: TObject);
begin
  if Length(S2EditFast.Text)>0 then
    S2AddFastbtn.Enabled:=True
  else
    S2AddFastbtn.Enabled:=False;
end;

procedure TMainForm.S2EditFastKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=Chr(13) then
  begin
    Key:=Chr(0);
    S2AddFastbtn.Click;
  end;
end;

procedure TMainForm.S2AddFastbtnClick(Sender: TObject);
begin
  AddFastInput(SGrid2,2,S2EditFast.Text);
  CheckGridCell(SGrid2,S2TypeLabel);
  S2EditFast.SelectAll;
  S2EditFast.SetFocus;
end;

end.
