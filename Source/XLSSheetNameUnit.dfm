object XLSSheetName: TXLSSheetName
  Left = 192
  Top = 114
  BorderStyle = bsDialog
  Caption = 'XLS Sheet Name'
  ClientHeight = 71
  ClientWidth = 393
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 16
    Width = 62
    Height = 13
    Caption = 'Sheet Name:'
  end
  object OKBtn: TxpButton
    Left = 311
    Top = 40
    Width = 75
    Height = 25
    Caption = 'OKBtn'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 3
    TabStop = True
    ModalResult = 1
  end
  object xpButton1: TxpButton
    Left = 223
    Top = 40
    Width = 75
    Height = 25
    Caption = 'Cancel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 2
    TabStop = True
    Cancel = True
    ModalResult = 2
  end
  object SheetName: TEdit
    Left = 88
    Top = 13
    Width = 297
    Height = 21
    TabOrder = 0
    Text = 'String Grid'
    OnChange = SheetNameChange
  end
  object NoSheetName: TCheckBox
    Left = 17
    Top = 43
    Width = 97
    Height = 17
    Caption = 'No sheet name'
    TabOrder = 1
    Visible = False
    OnClick = NoSheetNameClick
  end
end
