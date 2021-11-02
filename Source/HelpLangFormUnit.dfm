object HelpLangForm: THelpLangForm
  Left = 377
  Top = 173
  BorderIcons = []
  BorderStyle = bsNone
  ClientHeight = 80
  ClientWidth = 232
  Color = 14020827
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnDeactivate = FormDeactivate
  OnHide = FormHide
  OnKeyUp = FormKeyUp
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object psvRTFLabel1: TpsvRTFLabel
    Left = 6
    Top = 2
    Width = 121
    Height = 17
    Color = 14020827
    ParentColor = False
    Text.Strings = (
      
        '{\rtf1\fbidis\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0' +
        ' MS Sans Serif;}}'
      '\viewkind4\uc1\pard\ltrpar\lang14337\b\f0\fs20 Select Language:'
      '\par }')
    Transparent = False
    WordWrap = False
  end
  object LangList: TListBox
    Left = 5
    Top = 23
    Width = 154
    Height = 33
    ItemHeight = 13
    Items.Strings = (
      'English'
      'Persian')
    TabOrder = 0
    OnDblClick = LangListDblClick
    OnKeyDown = LangListKeyDown
  end
  object PromptCheck: TxpCheckBox
    Left = 7
    Top = 60
    Width = 145
    Height = 16
    Caption = 'Do not prompt again'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 1
    TabStop = True
    Color = 14020827
    Alignment = cbaRight
    OnClick = PromptCheckClick
  end
  object Display: TxpButton
    Left = 164
    Top = 10
    Width = 60
    Height = 23
    Caption = '&Display'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 2
    TabStop = True
    Default = True
    OnClick = DisplayClick
  end
  object Cancelbtn: TxpButton
    Left = 164
    Top = 45
    Width = 60
    Height = 23
    Caption = '&Cancel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    TabOrder = 3
    TabStop = True
    Cancel = True
    ModalResult = 2
    OnClick = CancelbtnClick
  end
end
