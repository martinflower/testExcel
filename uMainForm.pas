unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,  Excel_tlb, Vcl.StdCtrls, Vcl.Buttons,system.win.ComObj ;

type
  TMainForm = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}
type
    TItem = record
      FieldName : string;
      Position:Integer;
      Wide:Integer;
    end;
    TTableau=record
      Mois:Integer;
      Etat:Char;
    end;
Const
  Entete_colonne: array[1..6] of string=('Col1','Col2','Col3','Col4','Col5','Cle');
  tab: array[1..3] of TItem=((FieldName : 'Annee'; Position : 1;Wide:4),
                             (FieldName : 'Mois'; Position : 5;Wide:2),
                             (FieldName : 'Compte'; Position : 40;Wide:80));

function GetAppPath:String;
begin
  result:=ExtractFilePath(application.exename);
end;

procedure TMainForm.BitBtn1Click(Sender: TObject);
VAR Excel, {application}
wbk,  {workbook}
Wsh, {worksheet}
range: Olevariant;
i:integer;
LastDataExcel:integer;

Nom_Classeur:string;


Procedure SelectOnglet(aSheet: string);
begin
  wbk.sheets[aSheet].Activate; //connexion à un onglet
end;

procedure FormaterRange;
const
  //Alignement
  xlCenter=-4108;
  //Bordures
  //xlEdgeLeft=7;
  //xlEdgeTop=8;
  //xlEdgeBottom=9;
  //xlEdgeRight=10;
  //xlInsideVertical=11;
  //xlInsideHorizontal=12;
  //Largeur de ligne
  xlContinuous=1;
  //xlThin=2;
  //xlMedium=-4138;
  //Couleur
  xlAutomatic=-4105;
  //xlCellValue=1;
  //xlEqual=3;
var
  oActiveWindow, oActiveWorkBook, oSelection, oActiveSheet, oCells,oRows, oColumns :OleVariant;
  i,j :Cardinal;
  //R:olevariant;
  FC: olevariant; //objet de condition formaté

  procedure importation;
  var lst :Tstrings;
  cpt:integer;
  enregistrement: TTableau;
  begin
    lst:=TStringList.Create;
    try
      lst.LoadFromFile(GetAppPath+'data.txt');
      for cpt := 0 to lst.count-1 do
      begin
        enregistrement.Mois:=strtoint(copy(lst[cpt],5,2));
        enregistrement.Etat:=lst[cpt][7];    // le 8eme caractere
      oActiveSheet.Range['B'+trim(inttostr(LastDataExcel+cpt))].value:=enregistrement.Mois;
      oActiveSheet.Range['C'+trim(inttostr(LastDataExcel+cpt))].value:=enregistrement.etat;
      oActiveSheet.Range['I'+trim(inttostr(LastDataExcel+cpt))].value:='=B'+trim(inttostr(LastDataExcel+cpt))+'&C'+trim(inttostr(LastDataExcel+cpt));
      inc(LastDataExcel);
      end;
    finally
      lst.free;
    end;
  end;

begin
  //Variables objets
  //oExcel:=OleContainer1.OleObject.Application;
  oActiveWindow:=Excel.ActiveWindow;
  oActiveWorkBook:=Excel.ActiveWorkBook;
  oActiveSheet:=Wbk.Worksheets.Item[2];
//  oActiveSheet.Name:='Test2'; affecter un nom au
  SelectOnglet('Test2');
  oActiveSheet.activate;
  oRows:=oActiveSheet.Rows;
  oColumns:=oActiveSheet.Columns;
  oCells:=oActiveSheet.Cells;
  oSelection:=Excel.Selection;

  (* B. Tableau de données *)
  //En-têtes de colonnes
  for i:=2 to 8 do
    oCells.Item[2,i].Value:= 'En-tête col. ' + IntToStr(i);
  //Données
  for i:=3 to 8 do
    for j:=2 to 8 do
        oCells.Item[i,j].Value:=' Cells(' + IntToStr(i) + ', ' + IntToStr(j) + ') ';

  (* C. Mise en forme *)
  //Sélection des en-têtes
  oSelection:=oActiveSheet.Range[oCells.Item[2,2], oCells.Item[2,8]];
  //Propriétés de la police des en-têtes
  oSelection.Font.Bold:=1;
  oSelection.Font.ColorIndex := 9;
  //Couleur de fond de la sélection
  oSelection.Interior.ColorIndex := 15;
  //Centrer les en-têtes
  oSelection.HorizontalAlignment:= xlCenter;
  //Bordures
  for i:=7 to 11 do
    oSelection.Borders[i].LineStyle:=xlContinuous;
  //Sélection    Ligne -------------------
  oActiveSheet.Rows[3].Select;
  //Figer les volets
  oActiveWindow.FreezePanes:=1;
  //Ajustement automatique des colonnes
  for i:=2 to 8 do
    oColumns.Item[i].AutoFit;
  //Largeur de la colonne 1
  oColumns.Item[1].ColumnWidth:=1.71;
  oSelection:=oActiveSheet.Range[oCells.Item[1,13], oCells.Item[2,15]];
  oSelection.Font.Bold:=1;
  oSelection.Font.ColorIndex := 9;
  //Couleur de fond de la sélection
  oSelection.Interior.ColorIndex := 37;
  oActiveSheet.Range['I2'].value :='Clé';
  oActiveSheet.Range['N1'].value :='NB';
  oActiveSheet.Range['M2'].value :='Mois';
  oActiveSheet.Range['N2'].value :='Accepté';
  oActiveSheet.Range['O2'].value :='Rejeté';

  i:= 3;
  while not (oActiveSheet.Range['B'+trim(inttostr(i))].value='') do
  begin
    oActiveSheet.Range['I'+trim(inttostr(i))].value:='=B'+trim(inttostr(i))+'&C'+trim(inttostr(i));
    inc(i);
  end;
  LastDataExcel:=i;
  importation;
  for I := 1 to 12 do
  begin
    oActiveSheet.Range['M'+trim(inttostr(i+2))].value:=(inttostr(i));
    oActiveSheet.Range['N'+trim(inttostr(i+2))].formula:='=COUNTIF($I:$I,$M'+trim(inttostr(i+2))+'&LEFT(N$2,1))';
    oActiveSheet.Range['o'+trim(inttostr(i+2))].formula:='=COUNTIF($I:$I,$M'+trim(inttostr(i+2))+'&LEFT(O$2,1))';

//    oActiveSheet.Range['N'+trim(inttostr(i+2))].value:='=nb.si($I:$I;$M3&gauche(N$2;1))';
//    oActiveSheet.Range['O'+trim(inttostr(i+2))].value:='=nb.si($I:$I;$M3&gauche(N$2;1))';
  end;
  oSelection:=oActiveSheet.Range[oCells.Item[1,13], oCells.Item[1,15]];
  for i:=7 to 10 do
  begin
    oSelection.Borders[i].LineStyle:=xlContinuous;
    oSelection.Borders[i].Weight :=xlThick;
  end;
  oSelection:=oActiveSheet.Range[oCells.Item[2,13], oCells.Item[14,15]];
  for i:=7 to 12 do
  begin
    oSelection.Borders[i].LineStyle:=xlContinuous;
  end;
//  R := oactivesheet.Range['B1', 'B1'].EntireColumn;
//  FC := R.FormatConditions.Add(xlCellValue,
//         xlGreater, WS.Range['A1', 'A1'], EmptyParam);
//  FC.Interior.Color := clRed;
  fc:=oSelection.FormatConditions.add(xlCellValue,xlEqual, '0');
  fc.Interior.colorindex  :=rgbBlack;

  //Sélection finale
  oCells.Item[3, 2].Select;
end;

begin
  try
    Excel := CreateOleObject('Excel.Application');
    Nom_Classeur:=GetAppPath+'new.xlsx';
    // ajoute un nouveau classeur
    //Wbk:=Excel.Workbooks.Add;
    //Wbk.sheets.add;
    // ouvre un classeur
    Wbk:=Excel.Workbooks.Open(Nom_Classeur);
    Wsh:=Wbk.Worksheets.Item[1];

    Wsh.visible:=true;
    Wsh.Name:='Test1';
    wsh.activate;
    range:=Wsh.range['A1:I1'];
    range.Font.Bold:=true;
    range.Font.Underline:=true;

   for i:=1 to 6 do
        wsh.cells[1,i].value:=Entete_colonne[i];

   FormaterRange;
  finally
// Fermer Excel
//   wbk.saveas(GetAppPath+'new');
    wbk.save;
    Excel.Quit;
    Excel := unassigned;
  end;
End;


end.
