object Form1: TForm1
  Left = 218
  Top = 128
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Bigstruct'
  ClientHeight = 607
  ClientWidth = 808
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Icon.Data = {
    0000010001002020100000000000E80200001600000028000000200000004000
    0000010004000000000080020000000000000000000000000000000000000000
    0000000080000080000000808000800000008000800080800000C0C0C0008080
    80000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00CCC0
    000CCCC0000000000CCCC7777CCCCCCC0000CCCC00000000CCCC7777CCCCCCCC
    C0000CCCCCCCCCCCCCC7777CCCCC0CCCCC0000CCCCCCCCCCCC7777CCCCC700CC
    C00CCCC0000000000CCCC77CCC77000C0000CCCC00000000CCCC7777C7770000
    00000CCCC000000CCCC777777777C000C00000CCCC0000CCCC77777C777CCC00
    CC00000CCCCCCCCCC77777CC77CCCCC0CCC000CCCCC00CCCCC777CCC7CCCCCCC
    CCCC0CCCCCCCCCCCCCC7CCCCCCCCCCCC0CCCCCCCCCCCCCCCCCCCCCC7CCC70CCC
    00CCCCCCCC0CC0CCCCCCCC77CC7700CC000CCCCCC000000CCCCCC777CC7700CC
    0000CCCC00000000CCCC7777CC7700CC0000C0CCC000000CCC7C7777CC7700CC
    0000C0CCC000000CCC7C7777CC7700CC0000CCCC00000000CCCC7777CC7700CC
    000CCCCCC000000CCCCCC777CC7700CC00CCCCCCCC0CC0CCCCCCCC77CC770CCC
    0CCCCCCCCCCCCCCCCCCCCCC7CCC7CCCCCCCC0CCCCCCCCCCCCCC7CCCCCCCCCCC0
    CCC000CCCCC00CCCCC777CCC7CCCCC00CC00000CCCCCCCCCC77777CC77CCC000
    C00000CCCC0000CCCC77777C777C000000000CCCC000000CCCC777777777000C
    0000CCCC00000000CCCC7777C77700CCC00CCCC0000000000CCCC77CCC770CCC
    CC0000CCCCCCCCCCCC7777CCCCC7CCCCC0000CCCCCCCCCCCCCC7777CCCCCCCCC
    0000CCCC00000000CCCC7777CCCCCCC0000CCCC0000000000CCCC7777CCC0000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  Menu = MainMenu1
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object Bevel1: TBevel
    Left = 8
    Top = 8
    Width = 793
    Height = 561
  end
  object Label1: TLabel
    Left = 24
    Top = 64
    Width = 45
    Height = 16
    Caption = #1060'.'#1048'.'#1054'.:'
  end
  object Label2: TLabel
    Left = 24
    Top = 96
    Width = 200
    Height = 16
    Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1086#1088#1075#1072#1085#1080#1079#1072#1094#1080#1080':'
  end
  object Label3: TLabel
    Left = 24
    Top = 128
    Width = 44
    Height = 16
    Caption = #1040#1076#1088#1077#1089':'
  end
  object Label4: TLabel
    Left = 24
    Top = 160
    Width = 65
    Height = 16
    Caption = #1058#1077#1083#1077#1092#1086#1085':'
  end
  object Label5: TLabel
    Left = 24
    Top = 192
    Width = 113
    Height = 16
    Caption = #1044#1072#1090#1072' ('#1076#1076'/'#1084#1084'/'#1075#1075'):'
  end
  object Label6: TLabel
    Left = 24
    Top = 248
    Width = 182
    Height = 16
    Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1087#1088#1086#1076#1091#1082#1094#1080#1080':'
  end
  object Label7: TLabel
    Left = 24
    Top = 336
    Width = 90
    Height = 16
    Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077':'
  end
  object Label8: TLabel
    Left = 24
    Top = 432
    Width = 81
    Height = 16
    Caption = #1056#1077#1096#1077#1085#1080#1077' '#8470':'
  end
  object Label9: TLabel
    Left = 24
    Top = 496
    Width = 200
    Height = 16
    Caption = #1053#1072#1087#1088#1072#1074#1080#1090#1100' '#1076#1072#1085#1085#1086#1077' '#1088#1077#1096#1077#1085#1080#1077'...'
  end
  object Label10: TLabel
    Left = 24
    Top = 464
    Width = 64
    Height = 16
    Caption = #1056#1077#1096#1077#1085#1080#1077':'
  end
  object Label11: TLabel
    Left = 24
    Top = 32
    Width = 99
    Height = 16
    Caption = #1053#1086#1084#1077#1088' '#1079#1072#1103#1074#1082#1080':'
  end
  object Gauge1: TGauge
    Left = 24
    Top = 528
    Width = 761
    Height = 25
    ForeColor = clNavy
    Progress = 0
  end
  object Edit1: TEdit
    Left = 240
    Top = 24
    Width = 545
    Height = 24
    TabOrder = 0
    OnKeyPress = Edit1KeyPress
  end
  object Edit2: TEdit
    Left = 240
    Top = 56
    Width = 545
    Height = 24
    TabOrder = 1
  end
  object Edit3: TEdit
    Left = 240
    Top = 120
    Width = 545
    Height = 24
    TabOrder = 3
  end
  object Memo1: TMemo
    Left = 240
    Top = 248
    Width = 545
    Height = 73
    TabOrder = 6
  end
  object Memo2: TMemo
    Left = 240
    Top = 336
    Width = 545
    Height = 73
    TabOrder = 7
  end
  object Edit4: TEdit
    Left = 240
    Top = 152
    Width = 545
    Height = 24
    TabOrder = 4
  end
  object ComboBox1: TComboBox
    Left = 240
    Top = 88
    Width = 545
    Height = 24
    ItemHeight = 16
    TabOrder = 2
    OnSelect = ComboBox1Select
  end
  object Edit5: TEdit
    Left = 120
    Top = 424
    Width = 665
    Height = 24
    TabOrder = 8
    OnKeyPress = Edit5KeyPress
  end
  object Edit6: TEdit
    Left = 104
    Top = 456
    Width = 681
    Height = 24
    TabOrder = 9
  end
  object Button1: TButton
    Left = 8
    Top = 576
    Width = 793
    Height = 25
    Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' Word'
    TabOrder = 11
    OnClick = Button1Click
  end
  object Edit7: TEdit
    Left = 240
    Top = 488
    Width = 545
    Height = 24
    TabOrder = 10
  end
  object ME1: TMaskEdit
    Left = 240
    Top = 184
    Width = 56
    Height = 24
    EditMask = '!99/99/00;1; '
    MaxLength = 8
    TabOrder = 5
    Text = '00.00.02'
  end
  object MainMenu1: TMainMenu
    Left = 96
    Top = 56
    object N1: TMenuItem
      Caption = #1060#1072#1081#1083
      object N2: TMenuItem
        Caption = #1053#1086#1074#1099#1081
        OnClick = N2Click
      end
      object N6: TMenuItem
        Caption = #1054#1090#1082#1088#1099#1090#1100'/'#1059#1076#1072#1083#1080#1090#1100' '#1079#1072#1087#1080#1089#1100
        OnClick = N6Click
      end
      object N4: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1082#1072#1082' ...'
        OnClick = N4Click
      end
      object N5: TMenuItem
        Caption = '-'
      end
      object N7: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = N7Click
      end
    end
  end
  object DataSource1: TDataSource
    DataSet = Table1
    Left = 192
    Top = 56
  end
  object Table1: TTable
    TableType = ttParadox
    Left = 128
    Top = 56
  end
  object Table2: TTable
    Left = 160
    Top = 56
  end
  object SaveDialog1: TSaveDialog
    Filter = #1044#1086#1082#1091#1084#1077#1085#1090#1099' Word|*.doc'
    Left = 128
    Top = 128
  end
  object WordApplication1: TWordApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 160
    Top = 128
  end
  object WordDocument1: TWordDocument
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 192
    Top = 128
  end
end
