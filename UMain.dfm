object FMain: TFMain
  Left = 0
  Top = 0
  Caption = 'FMain'
  ClientHeight = 676
  ClientWidth = 1110
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 185
    Height = 676
    Align = alLeft
    TabOrder = 0
    object RadioGroup1: TRadioGroup
      Left = 2
      Top = 0
      Width = 177
      Height = 81
      Caption = #1058#1080#1087' '#1096#1072#1073#1083#1086#1085#1072' '#1057#1059#1054#1057
      ItemIndex = 0
      Items.Strings = (
        #1041#1072#1082#1072#1083#1072#1074#1088#1080#1072#1090
        #1052#1072#1075#1080#1089#1090#1088#1072#1090#1091#1088#1072
        #1057#1087#1077#1094#1080#1072#1083#1080#1090#1077#1090)
      TabOrder = 0
    end
    object BtLoadSuos: TButton
      Left = 10
      Top = 87
      Width = 169
      Height = 33
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1057#1059#1054#1057
      TabOrder = 1
      OnClick = BtLoadSuosClick
    end
    object Button1: TButton
      Left = 10
      Top = 160
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 2
      OnClick = Button1Click
    end
    object BtSprFile: TButton
      Left = 10
      Top = 128
      Width = 167
      Height = 25
      Caption = #1048#1089#1087#1088#1072#1074#1080#1090#1100' '#1092#1072#1081#1083
      TabOrder = 3
      OnClick = BtSprFileClick
    end
  end
  object Panel3: TPanel
    Left = 185
    Top = 0
    Width = 925
    Height = 676
    Align = alClient
    TabOrder = 1
    object Panel2: TPanel
      Left = 1
      Top = 440
      Width = 923
      Height = 235
      Align = alBottom
      TabOrder = 0
      object MeProt: TMemo
        Left = 1
        Top = 1
        Width = 921
        Height = 233
        Align = alClient
        ScrollBars = ssBoth
        TabOrder = 0
      end
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 440
      Height = 439
      Align = alLeft
      TabOrder = 1
      object Panel6: TPanel
        Left = 1
        Top = 1
        Width = 438
        Height = 79
        Align = alTop
        TabOrder = 0
        object LaNameShablon: TLabel
          Left = 4
          Top = 19
          Width = 76
          Height = 13
          Caption = 'LaNameShablon'
        end
        object LaNameFGOS: TLabel
          Left = 4
          Top = 57
          Width = 65
          Height = 13
          Caption = 'LaNameFGOS'
        end
        object Label1: TLabel
          Left = 4
          Top = 0
          Width = 40
          Height = 13
          Caption = #1064#1072#1073#1083#1086#1085
        end
        object Label2: TLabel
          Left = 4
          Top = 38
          Width = 29
          Height = 13
          Caption = #1060#1043#1054#1057
        end
      end
      object MeFileShab: TMemo
        Left = 1
        Top = 80
        Width = 438
        Height = 358
        Align = alClient
        ScrollBars = ssBoth
        TabOrder = 1
      end
    end
    object Panel5: TPanel
      Left = 441
      Top = 1
      Width = 483
      Height = 439
      Align = alClient
      TabOrder = 2
      object Panel7: TPanel
        Left = 1
        Top = 1
        Width = 481
        Height = 79
        Align = alTop
        TabOrder = 0
        object LaNameFile: TLabel
          Left = 4
          Top = 19
          Width = 54
          Height = 13
          Caption = 'LaNameFile'
        end
        object Label3: TLabel
          Left = 5
          Top = 0
          Width = 101
          Height = 13
          Caption = #1055#1088#1086#1074#1077#1088#1103#1077#1084#1099#1081' '#1057#1059#1054#1057
        end
        object Label4: TLabel
          Left = 4
          Top = 38
          Width = 40
          Height = 13
          Caption = #1054#1096#1080#1073#1086#1082
        end
        object LaColOsh: TLabel
          Left = 50
          Top = 38
          Width = 45
          Height = 13
          Caption = 'LaColOsh'
        end
        object Label5: TLabel
          Left = 101
          Top = 60
          Width = 30
          Height = 13
          Caption = #1043#1083#1072#1074#1072
        end
        object LaGlava: TLabel
          Left = 144
          Top = 60
          Width = 45
          Height = 13
          Caption = 'LaColOsh'
        end
        object Label6: TLabel
          Left = 144
          Top = 40
          Width = 29
          Height = 13
          Caption = #1057#1059#1054#1057
        end
        object Label7: TLabel
          Left = 192
          Top = 40
          Width = 40
          Height = 13
          Caption = #1064#1072#1073#1083#1086#1085
        end
        object Label8: TLabel
          Left = 248
          Top = 40
          Width = 29
          Height = 13
          Caption = #1060#1043#1054#1057
        end
        object LaGlavaSYOS: TLabel
          Left = 192
          Top = 59
          Width = 64
          Height = 13
          Caption = 'LaGlavaSYOS'
        end
        object LaGlavaFGOS: TLabel
          Left = 248
          Top = 59
          Width = 65
          Height = 13
          Caption = 'LaGlavaFGOS'
        end
      end
      object MeFileSYOS: TMemo
        Left = 1
        Top = 80
        Width = 481
        Height = 358
        Align = alClient
        ScrollBars = ssBoth
        TabOrder = 1
      end
    end
  end
  object OpenDialog: TOpenDialog
    Left = 128
  end
end
