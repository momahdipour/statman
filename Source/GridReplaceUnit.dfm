object GridReplaceForm: TGridReplaceForm
  Left = 272
  Top = 122
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'GridReplaceForm'
  ClientHeight = 120
  ClientWidth = 356
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 10
    Width = 56
    Height = 13
    Caption = 'Text to find:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 8
    Top = 42
    Width = 74
    Height = 13
    Caption = 'Text to replace:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object TextCombo: TComboBox
    Left = 85
    Top = 6
    Width = 265
    Height = 21
    ItemHeight = 13
    TabOrder = 0
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 63
    Width = 161
    Height = 51
    Caption = 'Search Options'
    TabOrder = 1
    object CaseSensitive: TxpCheckBox
      Left = 7
      Top = 12
      Width = 97
      Height = 17
      Caption = 'Case sensitive'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      TabOrder = 0
      TabStop = True
      Color = clBtnFace
      Alignment = cbaRight
    end
    object WholeWords: TxpCheckBox
      Left = 7
      Top = 30
      Width = 125
      Height = 17
      Caption = 'Whole words only'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      TabOrder = 1
      TabStop = True
      Color = clBtnFace
      Alignment = cbaRight
    end
  end
  object Findbtn: TxpButton
    Left = 183
    Top = 81
    Width = 75
    Height = 25
    Caption = 'Replace'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 2
    TabStop = True
    OnClick = FindbtnClick
  end
  object Closebtn: TxpButton
    Left = 272
    Top = 81
    Width = 75
    Height = 25
    Caption = 'Cancel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 3
    TabStop = True
    OnClick = ClosebtnClick
  end
  object ReplaceCombo: TComboBox
    Left = 85
    Top = 38
    Width = 265
    Height = 21
    ItemHeight = 13
    TabOrder = 4
  end
end
