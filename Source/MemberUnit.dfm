object MemberForm: TMemberForm
  Left = 364
  Top = 156
  AutoSize = True
  BorderStyle = bsToolWindow
  Caption = 'Member Pane'
  ClientHeight = 457
  ClientWidth = 222
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  OnClose = FormClose
  OnHide = FormHide
  PixelsPerInch = 96
  TextHeight = 13
  object DettachPopup: TPopupMenu
    OwnerDraw = True
    Left = 160
    Top = 80
    object StayOnTop1: TMenuItem
      AutoCheck = True
      Caption = 'Stay On Top'
      Checked = True
      OnClick = StayOnTop1Click
    end
  end
  object XPMenu1: TXPMenu
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMenuText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Color = clBtnFace
    IconBackColor = clBtnFace
    MenuBarColor = clBtnFace
    SelectColor = clHighlight
    SelectBorderColor = clHighlight
    SelectFontColor = clMenuText
    DisabledColor = clInactiveCaption
    SeparatorColor = clBtnFace
    CheckedColor = clHighlight
    IconWidth = 24
    DrawSelect = True
    UseSystemColors = True
    OverrideOwnerDraw = False
    Gradient = False
    FlatMenu = False
    AutoDetect = True
    Active = True
    Left = 72
    Top = 88
  end
end
