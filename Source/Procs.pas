unit Procs;

interface
uses
  MainUnit,Chart, TeEngine, TeCanvas, Graphics, TeeProcs;
  procedure LoadChart(Chart:TChart;const ChartRec:TChartSettingsRec);
  procedure LoadAxisSettings(ChartAxis:TChartAxis;const AxisSettings:TChartAxisSettings);
  procedure SaveChart(Chart:TChart;var ChartRec:TChartSettingsRec);

implementation

uses SysUtils, Classes;

procedure LoadChartPenSettings(Pen:TChartPen;PenSettings:TChartPenSettings);
begin
  Pen.Color:=PenSettings.Color;
  Pen.Mode:=PenSettings.Mode;
  Pen.SmallDots:=PenSettings.SmallDots;
  Pen.Style:=PenSettings.Style;
  Pen.Visible:=PenSettings.Visible;
  Pen.Width:=PenSettings.Width;
end;

procedure LoadFontSettings(Font:TFont;FontSettings:TFontSettings);
begin
  Font.Charset:=FontSettings.Charset;
  Font.Color:=FontSettings.Color;
  Font.Height:=FontSettings.Height;
  Font.Pitch:=FontSettings.Pitch;
  Font.Size:=FontSettings.Size;
  Font.Style:=FontSettings.Style;
end;

procedure LoadAxisSettings(ChartAxis:TChartAxis;const AxisSettings:TChartAxisSettings);
begin
    ChartAxis.Automatic:=AxisSettings.Automatic;
    ChartAxis.AutomaticMaximum:=AxisSettings.AutomaticMaximum;
    ChartAxis.AutomaticMinimum:=AxisSettings.AutomaticMinimum;
    ChartAxis.EndPosition:=AxisSettings.EndPosition;
    ChartAxis.GridCentered:=AxisSettings.GridCentered;
    ChartAxis.Increment:=AxisSettings.Increment;
    ChartAxis.Inverted:=AxisSettings.Inverted;
    ChartAxis.Labels:=AxisSettings.Labels;
    ChartAxis.LabelsAngle:=AxisSettings.LabelsAngle;
    ChartAxis.LabelsMultiLine:=AxisSettings.LabelsMultiLine;
    ChartAxis.LabelsOnAxis:=AxisSettings.LabelsOnAxis;
    ChartAxis.LabelsSeparation:=AxisSettings.LabelsSeparation;
    ChartAxis.LabelsSize:=AxisSettings.LabelsSize;
    ChartAxis.LabelStyle:=AxisSettings.LabelStyle;
    ChartAxis.Maximum:=AxisSettings.Maximum;
    ChartAxis.Minimum:=AxisSettings.Minimum;
    ChartAxis.MinorTickCount:=AxisSettings.MinorTickCount;
    ChartAxis.MinorTickLength:=AxisSettings.MinorTickLength;
    ChartAxis.PositionPercent:=AxisSettings.PositionPercent;
    ChartAxis.RoundFirstLabel:=AxisSettings.RoundFirstLabel;
    ChartAxis.StartPosition:=AxisSettings.StartPosition;
    ChartAxis.TickInnerLength:=AxisSettings.TickInnerLength;
    ChartAxis.TickLength:=AxisSettings.TickLength;
    ChartAxis.TickOnLabelsOnly:=AxisSettings.TickOnLabelsOnly;
    ChartAxis.TitleSize:=AxisSettings.TitleSize;
    ChartAxis.Visible:=AxisSettings.Visible;
    ChartAxis.Axis.Color:=AxisSettings.Axis.Color;
    LoadChartPenSettings(ChartAxis.Axis,AxisSettings.Axis);
    LoadChartPenSettings(ChartAxis.Grid,AxisSettings.Grid);
    LoadChartPenSettings(ChartAxis.MinorGrid,AxisSettings.MinorGrid);
    LoadChartPenSettings(ChartAxis.MinorTicks,AxisSettings.MinorTicks);
    LoadChartPenSettings(ChartAxis.Ticks,AxisSettings.Ticks);
    LoadChartPenSettings(ChartAxis.TicksInner,AxisSettings.TicksInner);
    ChartAxis.Title.Angle:=AxisSettings.Title.Angle;
    ChartAxis.Title.Caption:=AxisSettings.Title.Caption;
    LoadFontSettings(ChartAxis.Title.Font,AxisSettings.Title.Font);
    LoadFontSettings(ChartAxis.LabelsFont,AxisSettings.LabelsFont);
end;

procedure LoadChartBrushSettings(Brush:TChartBrush;BrushSettings:TChartBrushSettings);
begin
  Brush.Color:=BrushSettings.Color;
  Brush.Style:=BrushSettings.Style;
end;

procedure LoadChartWallSettings(ChartWall:TChartWall;const WallSettings:TChartWallSettings);
begin
  ChartWall.Color:=WallSettings.Color;
  ChartWall.Dark3D:=WallSettings.Dark3D;
  ChartWall.Size:=WallSettings.Size;
  LoadChartPenSettings(ChartWall.Pen,WallSettings.Pen);
  LoadChartBrushSettings(ChartWall.Brush,WallSettings.Brush);
end;
procedure LoadChartTitleSettings(Title:TChartTitle;TitleSettings:TChartTitleSettings);
var NewS:String;
begin
  Title.AdjustFrame:=TitleSettings.AdjustFrame;
  Title.Alignment:=TitleSettings.Alignment;
  LoadChartBrushSettings(Title.Brush,TitleSettings.Brush);
  Title.Color:=TitleSettings.Color;
  LoadFontSettings(Title.Font,TitleSettings.Font);
  LoadChartPenSettings(title.Frame,TitleSettings.Frame);
  Title.Text.Clear;
  NewS:=TitleSettings.Text;
  while Length(NewS)>0 do
  begin
    if Pos(Chr(13),NewS)>0 then
    begin
      Title.Text.Append(Copy(NewS,1,Pos(Chr(13),NewS)-1));
      NewS:=Copy(NewS,Pos(Chr(13),NewS)+1,Length(NewS)-Length(Copy(NewS,1,Pos(Chr(13),NewS)-1))-1);
    end
    else if Length(NewS)>0 then
    begin
      Title.Text.Append(NewS);
      NewS:='';
    end;
  end;
  Title.Visible:=TitleSettings.Visible;
end;

procedure LoadChartGradientSettings(Gradient:TChartGradient;GSettings:TChartGradientSettings);
begin
  Gradient.Direction:=GSettings.Direction;
  Gradient.EndColor:=GSettings.EndColor;
  Gradient.StartColor:=GSettings.StartColor;
  Gradient.Visible:=GSettings.Visible;
end;

procedure LoadChartLegendSettings(Legend:TChartLegend;LSettings:TChartLegendSettings);
begin
  Legend.Alignment:=LSettings.Alignment;
  LoadChartBrushSettings(Legend.Brush,LSettings.Brush);
  Legend.Color:=LSettings.Color;
  Legend.ColorWidth:=LSettings.ColorWidth;
  LoadChartPenSettings(Legend.DividingLines,LSettings.DividingLines);
  Legend.FirstValue:=LSettings.FirstValue;
  LoadFontSettings(Legend.Font,LSettings.Font);
  LoadChartPenSettings(Legend.Frame,LSettings.Frame);
  Legend.HorizMargin:=LSettings.HorizMargin;
  Legend.Inverted:=LSettings.Inverted;
  Legend.LegendStyle:=LSettings.LegendStyle;
  Legend.MaxNumRows:=LSettings.MaxNumRows;
  Legend.ResizeChart:=LSettings.ResizeChart;
  Legend.ShadowColor:=LSettings.ShadowColor;
  Legend.ShadowSize:=LSettings.ShadowSize;
  Legend.TextStyle:=LSettings.TextStyle;
  Legend.TopPos:=LSettings.TopPos;
  Legend.VertMargin:=LSettings.VertMargin;
  Legend.Visible:=LSettings.Visible;
end;

procedure LoadChart3DOptions(Options:TView3DOptions;Settings:TChart3DOptionsSettings);
begin
  Options.Elevation:=Settings.Elevation;
  Options.HorizOffset:=Settings.HorizOffset;
  Options.Orthogonal:=Settings.Orthogonal;
  Options.Perspective:=Settings.Perspective;
  Options.Rotation:=Settings.Rotation;
  Options.Tilt:=Settings.Tilt;
  Options.VertOffset:=Settings.VertOffset;
  Options.Zoom:=Settings.Zoom;
  Options.ZoomText:=Settings.ZoomText;
end;

procedure LoadChart(Chart:TChart;const ChartRec:TChartSettingsRec);
begin
  with Chart do
  begin
    AxisVisible:=ChartRec.AxisVisible;
    BackColor:=ChartRec.BackColor;
    if FileExists(ChartRec.BackImage) then
      BackImage.LoadFromFile(ChartRec.BackImage)
    else
      BackImage.Assign(nil);
    BackImageInside:=ChartRec.BackImageInside;
    BackImageMode:=ChartRec.BackImageMode;
    Chart3DPercent:=ChartRec.Chart3dPercent;
    Color:=ChartRec.Color;
    Monochrome:=ChartRec.Monochrome;
    PrintProportional:=ChartRec.PrintProportional;
    ScaleLastPage:=ChartRec.ScaleLastPage;
    View3D:=ChartRec.View3D;
    View3DWalls:=ChartRec.View3DWalls;
    LoadAxisSettings(LeftAxis,ChartRec.LeftAxis);
    LoadAxisSettings(BottomAxis,ChartRec.BottomAxis);
    LoadChartWallSettings(Backwall,ChartRec.BackWall);
    LoadChartWallSettings(BottomWall,ChartRec.BottomWall);
    LoadChartWallSettings(LeftWall,ChartRec.LeftWall);
    LoadChartTitleSettings(Foot,ChartRec.Foot);
    LoadChartTitleSettings(Title,ChartRec.Title);
    LoadChartPenSettings(Frame,ChartRec.Frame);
    LoadChartGradientSettings(Gradient,ChartRec.Gradient);
    LoadChartLegendSettings(Legend,ChartRec.Legend);
    LoadChart3DOptions(View3DOptions,ChartRec.View3DOptions);
  end;
end;

{*******************************************************
 *******************************************************}

procedure GetChartGradientSettings(ChartGradient:TChartGradient;var Gradint:TChartGradientSettings);
begin
  Gradint.Direction:=ChartGradient.Direction;
  Gradint.EndColor:=ChartGradient.EndColor;
  Gradint.StartColor:=ChartGradient.StartColor;
  Gradint.Visible:=ChartGradient.Visible;
end;

procedure GetChartBrushSettings(ChartBrush:TChartBrush;var Brush:TChartBrushSettings);
begin
  Brush.Color:=ChartBrush.Color;
  Brush.Style:=ChartBrush.Style;
end;

Procedure GetChart3DOptions(ChartOptions:TView3DOptions;var Options:TChart3DOptionsSettings);
begin
  Options.Elevation:=ChartOptions.Elevation;
  Options.HorizOffset:=ChartOptions.HorizOffset;
  Options.Orthogonal:=ChartOptions.Orthogonal;
  Options.Perspective:=ChartOptions.Perspective;
  Options.Rotation:=ChartOptions.Rotation;
  Options.Tilt:=ChartOptions.Tilt;
  Options.VertOffset:=ChartOptions.VertOffset;
  Options.Zoom:=ChartOptions.Zoom;
  Options.ZoomText:=ChartOptions.ZoomText;
end;

procedure GetChartPenSettings(ChartPen:TChartPen;var Pen:TChartPenSettings);
begin
  Pen.Color:=ChartPen.Color;
  Pen.Mode:=ChartPen.Mode;
  Pen.SmallDots:=ChartPen.SmallDots;
  Pen.Style:=ChartPen.Style;
  Pen.Visible:=ChartPen.Visible;
  Pen.Width:=ChartPen.Width;
end;

procedure GetFontSettings(AFont:TFont;var FontSettins:TFontSettings);
begin
  FontSettins.Charset:=AFont.Charset;
  FontSettins.Color:=AFont.Color;
  FontSettins.Height:=AFont.Height;
  FontSettins.Pitch:=AFont.Pitch;
  FontSettins.Size:=AFont.Size;
  FontSettins.Style:=AFont.Style;
end;

procedure GetChartTitleSettings(ChartTitle:TChartTitle;var ATitle:TChartTitleSettings);
var i:Integer;
begin
  ATitle.AdjustFrame:=ChartTitle.AdjustFrame;
  ATitle.Alignment:=ChartTitle.Alignment;
  GetChartBrushSettings(ChartTitle.Brush,ATitle.Brush);
  ATitle.Color:=ChartTitle.Color;
  GetFontSettings(ChartTitle.Font,ATitle.Font);
  GetChartPenSettings(ChartTitle.Frame,ATitle.Frame);
  ATitle.Text:='';
  if ChartTitle.Text.Count>0 then
  begin
    ATitle.Text:=ChartTitle.Text.Strings[0];
    for i:=1 to (ChartTitle.Text.Count-1) do
      ATitle.Text:=ATitle.Text+Chr(13)+ChartTitle.Text.Strings[i];
  end;
  ATitle.Visible:=ChartTitle.Visible;
end;

procedure GetChartWallSettings(ChartWall:TChartWall;var Wall:TChartWallSettings);
begin
  Wall.Color:=ChartWall.Color;
  Wall.Dark3D:=ChartWall.Dark3D;
  Wall.Size:=ChartWall.Size;
  GetChartPenSettings(ChartWall.Pen,Wall.Pen);
  GetChartBrushSettings(ChartWall.Brush,Wall.Brush);
end;

procedure GetChartLegendSettings(ChartLegend:TChartLegend;var Legend:TChartLegendSettings);
begin
  Legend.Alignment:=ChartLegend.Alignment;
  GetChartBrushSettings(ChartLegend.Brush,Legend.Brush);
  Legend.Color:=ChartLegend.Color;
  Legend.ColorWidth:=ChartLegend.ColorWidth;
  GetChartPenSettings(ChartLegend.DividingLines,Legend.DividingLines);
  Legend.FirstValue:=ChartLegend.FirstValue;
  GetFontSettings(ChartLegend.Font,Legend.Font);
  GetChartPenSettings(ChartLegend.Frame,Legend.Frame);
  Legend.HorizMargin:=ChartLegend.HorizMargin;
  Legend.Inverted:=ChartLegend.Inverted;
  Legend.LegendStyle:=ChartLegend.LegendStyle;
  Legend.MaxNumRows:=ChartLegend.MaxNumRows;
  Legend.ResizeChart:=ChartLegend.ResizeChart;
  Legend.ShadowColor:=ChartLegend.ShadowColor;
  Legend.ShadowSize:=ChartLegend.ShadowSize;
  Legend.TextStyle:=ChartLegend.TextStyle;
  Legend.TopPos:=ChartLegend.TopPos;
  Legend.VertMargin:=ChartLegend.VertMargin;
  Legend.Visible:=ChartLegend.Visible;
end;

procedure GetChartAxisSettings(ChartAxis:TChartAxis;var Axis:TChartAxisSettings);
begin
  Axis.Automatic:=ChartAxis.Automatic;
  Axis.AutomaticMaximum:=ChartAxis.AutomaticMaximum;
  Axis.AutomaticMinimum:=ChartAxis.AutomaticMinimum;
  Axis.EndPosition:=ChartAxis.EndPosition;
  Axis.GridCentered:=ChartAxis.GridCentered;
  Axis.Increment:=ChartAxis.Increment;
  Axis.Inverted:=ChartAxis.Inverted;
  Axis.Labels:=ChartAxis.Labels;
  Axis.LabelsAngle:=ChartAxis.LabelsAngle;
  Axis.LabelsMultiLine:=ChartAxis.LabelsMultiLine;
  Axis.LabelsOnAxis:=ChartAxis.LabelsOnAxis;
  Axis.LabelsSeparation:=ChartAxis.LabelsSeparation;
  Axis.LabelsSize:=ChartAxis.LabelsSize;
  Axis.LabelStyle:=ChartAxis.LabelStyle;
  Axis.Maximum:=ChartAxis.Maximum;
  Axis.Minimum:=ChartAxis.Minimum;
  Axis.MinorTickCount:=ChartAxis.MinorTickCount;
  Axis.MinorTickLength:=ChartAxis.MinorTickLength;
  Axis.PositionPercent:=ChartAxis.PositionPercent;
  Axis.RoundFirstLabel:=ChartAxis.RoundFirstLabel;
  Axis.StartPosition:=ChartAxis.StartPosition;
  Axis.TickInnerLength:=ChartAxis.TickInnerLength;
  Axis.TickLength:=ChartAxis.TickLength;
  Axis.TickOnLabelsOnly:=ChartAxis.TickOnLabelsOnly;
  Axis.TitleSize:=ChartAxis.TitleSize;
  Axis.Visible:=ChartAxis.Visible;
  Axis.Title.Angle:=ChartAxis.Title.Angle;
  Axis.Title.Caption:=ChartAxis.Title.Caption;
  GetFontSettings(ChartAxis.Title.Font,Axis.Title.Font);
  GetFontSettings(ChartAxis.LabelsFont,Axis.LabelsFont);
  GetChartPenSettings(ChartAxis.Axis,Axis.Axis);
  GetChartPenSettings(ChartAxis.Grid,Axis.Grid);
  GetChartPenSettings(ChartAxis.MinorGrid,Axis.MinorGrid);
  GetChartPenSettings(ChartAxis.MinorTicks,Axis.MinorTicks);
  GetChartPenSettings(ChartAxis.Ticks,Axis.Ticks);
  GetChartPenSettings(ChartAxis.TicksInner,Axis.TicksInner);
end;

procedure SaveChart(Chart:TChart;var ChartRec:TChartSettingsRec);
begin
  ChartRec.AxisVisible:=Chart.AxisVisible;
  ChartRec.BackColor:=Chart.BackColor;
//  ChartRec.BackImage:=Chart
  ChartRec.BackImageInside:=Chart.BackImageInside;
  ChartRec.BackImageMode:=Chart.BackImageMode;
  ChartRec.Chart3dPercent:=Chart.Chart3DPercent;
  ChartRec.Color:=Chart.Color;
  ChartRec.Monochrome:=Chart.Monochrome;
  ChartRec.PrintProportional:=Chart.PrintProportional;
  ChartRec.ScaleLastPage:=Chart.ScaleLastPage;
  ChartRec.View3D:=Chart.View3D;
  ChartRec.View3DWalls:=Chart.View3DWalls;
  GetChartAxisSettings(Chart.LeftAxis,ChartRec.LeftAxis);
  GetChartAxisSettings(Chart.BottomAxis,ChartRec.BottomAxis);
  GetChartWallSettings(Chart.BackWall,ChartRec.BackWall);
  GetChartWallSettings(Chart.BottomWall,ChartRec.BottomWall);
  GetChartWallSettings(Chart.LeftWall,ChartRec.LeftWall);
  GetChartTitleSettings(Chart.Foot,ChartRec.Foot);
  GetChartPenSettings(Chart.Frame,ChartRec.Frame);
  GetChartGradientSettings(Chart.Gradient,ChartRec.Gradient);
  GetChartTitleSettings(Chart.Title,ChartRec.Title);
  GetChartLegendSettings(Chart.Legend,ChartRec.Legend);
  GetChart3DOptions(Chart.View3DOptions,ChartRec.View3DOptions);
end;

{*******************************************************
 *******************************************************}

procedure SetColorsAndAppearanceTab(Chart:TChart);
begin
  with MainForm do
  begin
    UseColors.Checked:=Chart.Monochrome;
    BackColor.Color:=Chart.BackColor;
    if Chart.BackColor=clTeeColor then
      bc.Checked:=False
    else
      bc.Checked:=True;
    FrameColor.Color:=Chart.Frame.Color;
    if Chart.Frame.Color=clTeeColor then
      bf.Checked:=False
    else
      bf.Checked:=True;
    StartColor.Selected:=Chart.Gradient.StartColor;
    EndColor.Selected:=Chart.Gradient.EndColor;
    cg.Checked:=Chart.Gradient.Visible;
    PutInside.Checked:=Chart.BackImageInside;
    case Chart.BackImageMode of
      pbmTile: TileRadio.Checked:=True;
      pbmStretch: StretchRadio.Checked:=True;
      pbmCenter: CenterRadio.Checked:=True;
    end;
    if Chart.BackImage<>nil then
      UseBackImage.Checked:=True
    else
      UseBackImage.Checked:=False;
  end;
end;

procedure SetChartWallsTab(Chart:TChart);
begin
  with MainForm do
  begin
    View3DCheck.Checked:=Chart.View3DWalls;
    ColorSelector6.Color:=Chart.LeftWall.Color;
    SpinEdit5.Value:=Chart.LeftWall.Size;
    xpCheckBox13.Checked:=Chart.LeftWall.Pen.Visible;
    ColorSelector7.Color:=Chart.LeftWall.Pen.Color;
    penwidthcombo3.ItemIndex:=Chart.LeftWall.Pen.Width-1;
    penstylecombo2.ItemIndex:=Ord(Chart.LeftWall.Pen.Style);
    ColorSelector9.Color:=Chart.BackWall.Color;
    SpinEdit6.Value:=Chart.BackWall.Size;
    xpCheckBox14.Checked:=Chart.BackWall.Pen.Visible;
    ColorSelector8.Color:=Chart.BackWall.Pen.Color;
    penwidthcombo4.ItemIndex:=Chart.BackWall.Pen.Width-1;
    penstylecombo3.ItemIndex:=Ord(Chart.BackWall.Pen.Style);
  end;
end;

procedure SetChartAxisTab(Axis:TChartAxis);
begin
  with MainForm do
  begin
    xpCheckBox22.Checked:=Axis.Visible;
    RadioButton8.Checked:=Axis.Automatic;
    RadioButton9.Checked:=not(Axis.Automatic);
    minspin.Value:=Round(Axis.Minimum);
    MaxSpin.Value:=Round(Axis.Maximum);
    Increment.Value:=Round(Axis.Increment);
    SpinEdit10.Value:=Round(Axis.PositionPercent);
    SpinEdit11.Value:=Round(Axis.EndPosition);
    SpinEdit12.Value:=Round(Axis.StartPosition);
    xpCheckBox21.Checked:=Axis.Inverted;
  end;
end;

procedure SetChartLegendTab(Legend:TChartLegend);
begin
  with MainForm do
  begin
    xpCheckBox7.Checked:=Legend.Visible;
    LegendPos.ItemIndex:=Ord(Legend.Alignment);
    ColorSelector3.Color:=Legend.Color;
    LegendWidth.Value:=Legend.ColorWidth;
    Legendfont.Font:=Legend.Font;
    if Legend.TextStyle=ltsLeftValue then
      LegendStyle.ItemIndex:=0
    else if Legend.TextStyle=ltsLeftPercent then
      LegendStyle.ItemIndex:=1;
    LegendStyle.Hint:=LegendStyle.Items.Strings[LegendStyle.ItemIndex];
    xpCheckBox8.Checked:=Legend.Inverted;
    xpCheckBox9.Checked:=Legend.ResizeChart;

    xpCheckBox10.Checked:=Legend.Frame.Visible;
    lcolor.Color:=Legend.Frame.Color;
    iwidth2.ItemIndex:=Legend.Frame.Width-1;
    istyle2.ItemIndex:=Ord(Legend.Frame.Style);
    HorizMarg.Value:=Legend.HorizMargin;
    VertMarg.Value:=Legend.VertMargin;
    ShadowSize.Value:=Legend.ShadowSize;
    ShadowColor.Color:=Legend.ShadowColor;
    if Legend.ShadowSize=0 then
    begin
      ShadowSize.Value:=3;
      xpCheckBox11.Checked:=False;
    end
    else
      xpCheckBox11.Checked:=True;
  end;
end;

procedure SetChart3DViewTab(Chart:TChart);
begin
  with MainForm do
  begin
    xpCheckBox6.Checked:=Chart.View3D;
    ZoomTrack.Position:=Chart.View3DOptions.Zoom;
    Chart3DTrack.Position:=Chart.Chart3DPercent;
    NormalView.Checked:=Chart.View3DOptions.Orthogonal;
    CustomizedView.Checked:=not(Chart.View3DOptions.Orthogonal);
    XR.Position:=Chart.View3DOptions.Elevation;
    YR.Position:=Chart.View3DOptions.Rotation;
    ZR.Position:=Chart.View3DOptions.Tilt;
    PerspectiveTrack.Position:=Chart.View3DOptions.Perspective;    
  end;
end;

procedure SetTitleAndFooterTab(Chart:TChart);
begin
  with MainForm do
  begin
    {*************  TITLE  *************}
    xpCheckBox15.Checked:=Chart.Title.Visible;
    ResizeT.Checked:=Chart.Title.AdjustFrame;
    ColorT.Color:=Chart.Title.Color;
    xpCheckBox16.Checked:=Chart.Title.Frame.Visible;
    bColor.Color:=Chart.Title.Frame.Color;
    bWidth.ItemIndex:=Chart.Title.Frame.Width-1;
    bStyle.ItemIndex:=Ord(Chart.Title.Frame.Style);
    ChartTitleEdit.Lines:=Chart.Title.Text;
    ChartTitleEdit.Font:=Chart.Title.Font;
    if (fsBold in Chart.Title.Font.Style) then
      BoldBtn1.Down:=True
    else
      BoldBtn1.Down:=False;
    if (fsItalic in Chart.Title.Font.Style) then
      ItalicBtn1.Down:=True
    else
      ItalicBtn1.Down:=False;
    if (fsUnderline in Chart.Title.Font.Style) then
      UnderBtn1.Down:=True
    else
      UnderBtn1.Down:=False;
    case Chart.Title.Alignment of
      taLeftJustify: LeftAl.Down:=True;
      taCenter: CenterAl.Down:=True;
      taRightJustify: RightAl.Down:=True;
    end;
    SizeCombo.ItemIndex:=SizeCombo.Items.IndexOf(IntToStr(Chart.Title.Font.Size));
    ColorBtn1.SelectedColor:=Chart.Title.Font.Color;
    {*************  FOOTER  *************}
    xpCheckBox17.Checked:=Chart.Foot.Visible;
    ResizeF.Checked:=Chart.Foot.AdjustFrame;
    ColorF.Color:=Chart.Foot.Color;
    xpCheckBox20.Checked:=Chart.Foot.Frame.Visible;
    bColorf.Color:=Chart.Foot.Frame.Color;
    bWidthf.ItemIndex:=Chart.Foot.Frame.Width-1;
    bStylef.ItemIndex:=Ord(Chart.Foot.Frame.Style);
    ChartFooterEdit.Lines:=Chart.Foot.Text;
    ChartFooterEdit.Font:=Chart.Foot.Font;
    if fsBold in Chart.Foot.Font.Style then
      BoldBtn2.Down:=True
    else
      BoldBtn2.Down:=False;
    if fsItalic in Chart.Foot.Font.Style then
      ItalicBtn2.Down:=True
    else
      ItalicBtn2.Down:=False;
    if fsUnderline in Chart.Foot.Font.Style then
      UnderBtn2.Down:=True
    else
      UnderBtn2.Down:=False;
    case Chart.Foot.Alignment of
      taLeftJustify: LeftAl2.Down:=True;
      taCenter: CenterAl2.Down:=True;
      taRightJustify: RightAl2.Down:=True;
    end;
    SizeCombo2.ItemIndex:=SizeCombo2.Items.IndexOf(IntToStr(Chart.Foot.Font.Size));
    ColorBtn2.SelectedColor:=Chart.Foot.Font.Color;
  end;
end;

procedure SetChartControllerSettings(Chart:TChart);
begin
  SetColorsAndAppearanceTab(Chart);
  SetChartWallsTab(Chart);
  SetChartAxisTab(Chart.LeftAxis);
  SetChartLegendTab(Chart.Legend);
  SetChart3DViewTab(Chart);
  SetTitleAndFooterTab(Chart);
end;

end.
