object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'MainForm'
  ClientHeight = 577
  ClientWidth = 812
  Color = clBtnFace
  CustomTitleBar.CaptionAlignment = taCenter
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object BitBtn1: TBitBtn
    Left = 24
    Top = 16
    Width = 75
    Height = 25
    Caption = 'Formatage'
    TabOrder = 0
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 24
    Top = 47
    Width = 75
    Height = 25
    Caption = 'Importation'
    TabOrder = 1
    OnClick = BitBtn2Click
  end
  object ExcelApplication1: TExcelApplication
    AutoConnect = True
    ConnectKind = ckRunningOrNew
    AutoQuit = True
    Left = 96
    Top = 160
  end
  object ExcelWorkbook1: TExcelWorkbook
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 160
    Top = 216
  end
  object ExcelWorksheet1: TExcelWorksheet
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    Left = 272
    Top = 232
  end
end
