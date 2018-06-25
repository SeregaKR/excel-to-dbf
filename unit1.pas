unit Unit1;

{$mode objfpc}{$H+}
{$codepage cp1251}

interface

uses
  Classes, SysUtils, dbf, db, Forms, Controls, Graphics, Dialogs, StdCtrls,
  EditBtn, DBGrids, ExtCtrls,  LazUTF8, lconvencoding, dbf_dbffile,
  fpspreadsheet, fpstypes, fpspreadsheetgrid, xlsBiff8, fpsexport;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    Button1: TButton;
    DataSource1: TDataSource;
    Dbf1: TDbf;
    DBGrid1: TDBGrid;
    FileNameEdit1: TFileNameEdit;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure Button1Click(Sender: TObject);
    procedure DbfGetTextHandler(Sender: TField; var AText: string; DisplayText: Boolean);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
  private
    procedure ExportToDBF(AWorksheet: TsWorksheet; AFileName: String);

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}


{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin
  sWorksheetGrid1.LoadFromSpreadsheetFile(FilenameEdit1.Filename, sfExcel8);
  ExportToDBF(sWorksheetGrid1.Worksheet, ChangeFileExt(FileNameEdit1.Filename, '.dbf'));
end;

procedure TForm1.DbfGetTextHandler(Sender: TField; var AText: string; DisplayText: Boolean);
begin
  if DisplayText then
    AText := CP1251ToUTF8(Sender.AsString);
end;

procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  Dbf1.Close;
end;

procedure TForm1.ExportToDBF(AWorksheet: TsWorksheet; AFileName: String);
var
  i: Integer;
  f: TField;
  r, c: Cardinal;
  cell: PCell;
begin
  DbfGlobals.DefaultCreateCodePage := 1251; //default encoding of the dbf
  DbfGlobals.DefaultOpenCodePage := 1251; //default encoding of the dbf
  if Dbf1.Active then Dbf1.Close;
  if FileExists(AFileName) then DeleteFile(AFileName);

  Dbf1.FilePathFull := ExtractFilePath(AFileName);
  Dbf1.TableName := ExtractFileName(AFileName);
  Dbf1.TableLevel := 25;  // DBase IV: 4 - most widely used; or 25 = FoxPro supports nfCurrency
  Dbf1.LanguageID := $C9; //russian language by default
  Dbf1.FieldDefs.Clear;
  //below are the fields in excel
  Dbf1.FieldDefs.Add('fam', ftString);
  Dbf1.FieldDefs.Add('im', ftString);
  Dbf1.FieldDefs.Add('ot', ftString);
  Dbf1.FieldDefs.Add('dt', ftDateTime);
  Dbf1.FieldDefs.Add('raion', ftString);
  Dbf1.FieldDefs.Add('adres', ftString);
  Dbf1.FieldDefs.Add('tel', ftString);
  Dbf1.FieldDefs.Add('per_pm', ftInteger);
  Dbf1.FieldDefs.Add('priznak_do', ftInteger);
  Dbf1.FieldDefs.Add('status_pm', ftInteger);
  Dbf1.FieldDefs.Add('per_pm_ut', ftInteger);
  Dbf1.FieldDefs.Add('fname', ftString);

  Dbf1.CreateTable;
  Dbf1.Open;

  for f in Dbf1.Fields do
    f.OnGetText := @DbfGetTextHandler;

  // Skip row 0 which contains the headers
  for r := 1 to AWorksheet.GetLastRowIndex do begin
    Dbf1.Append;
    for c := 0 to Dbf1.FieldDefs.Count-1 do begin
      f := Dbf1.Fields[c];
      cell := AWorksheet.FindCell(r, c);
      if cell = nil then
        f.Value := NULL
      else
        case cell^.ContentType of
          cctUTF8String: f.AsString := UTF8ToCP1251(cell^.UTF8StringValue);
          cctNumber: f.AsFloat := cell^.NumberValue;
          cctDateTime: f.AsDateTime := cell^.DateTimeValue;
          else f.AsString := UTF8ToCP1251(AWorksheet.ReadAsText(cell));
        end;
    end;
    Dbf1.Post;
  end;
end;

end.

