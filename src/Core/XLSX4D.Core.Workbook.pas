unit XLSX4D.Core.Workbook;

interface

uses
  System.Generics.Collections,
  XLSX4D.Intf.Workbook,
  XLSX4D.Intf.Worksheet;

type
  TXLSX4DWorkbook = class(TInterfacedObject, IXLSX4DWorkbook)
  private
    FWorksheets: TList<IXLSX4DWorksheet>;
    FActiveWorksheet: IXLSX4DWorksheet;
    FFilePath: string;
    FModified: Boolean;

    procedure CheckWorksheetIndex(AIndex: Integer);
    function GenerateWorksheetName: string;
  protected
    function GetWorksheetCount: Integer;
    function GetWorksheets: TList<IXLSX4DWorksheet>;
    function GetActiveWorksheet: IXLSX4DWorksheet;
    function GetFilePath: string;

    procedure SetActiveWorksheet(const AWorksheet: IXLSX4DWorksheet);
  public
    constructor Create; overload;
    constructor Create(const AFilePath: string); overload;
    destructor Destroy; override;

    function AddWorksheet(const AName: string = ''): IXLSX4DWorksheet;
    function GetWorksheet(AIndex: Integer): IXLSX4DWorksheet; overload;
    function GetWorksheet(const AName: string = ''): IXLSX4DWorksheet; overload;
    procedure DeleteWorksheet(AIndex: Integer); overload;
    procedure DeleteWorksheet(const AName: string = ''); overload;

    procedure SaveAs(const AFileName: string);
    procedure Save;
    procedure Close;

    property WorksheetCount: Integer read GetWorksheetCount;
    property Worksheets: TList<IXLSX4DWorksheet> read GetWorksheets;
    property ActiveWorksheet: IXLSX4DWorksheet read GetActiveWorksheet write SetActiveWorksheet;
    property FilePath: string read GetFilePath;
    property Modified: Boolean read FModified write FModified;
  end;

implementation

uses
  System.SysUtils,
  System.Classes,
  XLSX4D.Core.Worksheet;

{ TXLSX4DWorkbook }

function TXLSX4DWorkbook.AddWorksheet(const AName: string): IXLSX4DWorksheet;
var
  WorksheetName: string;
  LSheet: IXLSX4DWorksheet;
  NewWorksheet: IXLSX4DWorksheet;
begin
  if Trim(AName) = '' then
    WorksheetName := GenerateWorksheetName
  else
  begin
    for LSheet in FWorksheets do
    begin
      if SameText(LSheet.Name, AName) then
        raise Exception.CreateFmt('Worksheet with name "%s" already exists', [AName]);
    end;
    WorksheetName := AName;
  end;

  NewWorksheet := TXLSX4DWorksheet.Create(WorksheetName, Self);
  FWorksheets.Add(NewWorksheet);

  if FWorksheets.Count = 1 then
    FActiveWorksheet := NewWorksheet;

  FModified := True;
  Result := NewWorksheet;
end;

procedure TXLSX4DWorkbook.CheckWorksheetIndex(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex >= FWorksheets.Count) then
    raise Exception.CreateFmt('Worksheet index %d is out of bounds (0..%d)',
      [AIndex, FWorksheets.Count - 1]);
end;

procedure TXLSX4DWorkbook.Close;
begin
  FWorksheets.Clear;
  FActiveWorksheet := nil;
  FModified := False;
end;

constructor TXLSX4DWorkbook.Create(const AFilePath: string);
begin
  Create;
  FFilePath := AFilePath;
end;

constructor TXLSX4DWorkbook.Create;
begin
  inherited Create;
  FWorksheets := TList<IXLSX4DWorksheet>.Create;
  FFilePath := '';
  FModified := False;

  AddWorksheet('Sheet1');
end;

procedure TXLSX4DWorkbook.DeleteWorksheet(const AName: string);
var
  Index: Integer;
  I: Integer;
begin
  Index := -1;

  for I := 0 to FWorksheets.Count - 1 do
  begin
    if SameText(FWorksheets[I].Name, AName) then
    begin
      Index := I;
      Break;
    end;
  end;

  if Index = -1 then
    raise Exception.CreateFmt('Worksheet "%s" not found', [AName]);

  DeleteWorksheet(Index);
end;

procedure TXLSX4DWorkbook.DeleteWorksheet(AIndex: Integer);
var
  Worksheet: IXLSX4DWorksheet;
  I: Integer;
begin
  if FWorksheets.Count <= 1 then
    raise Exception.Create('Cannot delete the last worksheet');

  CheckWorksheetIndex(AIndex);

  Worksheet := FWorksheets[AIndex];

  if Worksheet = FActiveWorksheet then
  begin
    if AIndex > 0 then
      FActiveWorksheet := FWorksheets[AIndex - 1]
    else if FWorksheets.Count > 1 then
      FActiveWorksheet := FWorksheets[1]
    else
      FActiveWorksheet := nil;
  end;

  FWorksheets.Delete(AIndex);

  FModified := True;
end;

destructor TXLSX4DWorkbook.Destroy;
begin
  Close;
  FWorksheets.Free;
  inherited;
end;

function TXLSX4DWorkbook.GenerateWorksheetName: string;
var
  I: Integer;
  Name: string;
  Exists: Boolean;
  LSheet: IXLSX4DWorksheet;
begin
  I := FWorksheets.Count + 1;
  repeat
    Name := Format('Sheet%d', [I]);
    Exists := False;

    for LSheet in FWorksheets do
    begin
      if SameText(LSheet.Name, Name) then
      begin
        Exists := True;
        Break;
      end;
    end;

    if not Exists then
      Break;

    Inc(I);
  until False;

  Result := Name;
end;

function TXLSX4DWorkbook.GetActiveWorksheet: IXLSX4DWorksheet;
begin
  if not Assigned(FActiveWorksheet) and (FWorksheets.Count > 0) then
    FActiveWorksheet := FWorksheets[0];

  Result := FActiveWorksheet;
end;

function TXLSX4DWorkbook.GetFilePath: string;
begin
  Result := FFilePath;
end;

function TXLSX4DWorkbook.GetWorksheet(AIndex: Integer): IXLSX4DWorksheet;
begin
  CheckWorksheetIndex(AIndex);
  Result := FWorksheets[AIndex];
end;

function TXLSX4DWorkbook.GetWorksheet(const AName: string): IXLSX4DWorksheet;
var
  LSheet: IXLSX4DWorksheet;
begin
  Result := nil;

  for LSheet in FWorksheets do
  begin
    if SameText(LSheet.Name, AName) then
    begin
      Result := LSheet;
      Break;
    end;
  end;

  if not Assigned(Result) then
    raise Exception.CreateFmt('Worksheet "%s" not found', [AName]);
end;

function TXLSX4DWorkbook.GetWorksheetCount: Integer;
begin
  Result := FWorksheets.Count;
end;

function TXLSX4DWorkbook.GetWorksheets: TList<IXLSX4DWorksheet>;
begin
  Result := FWorksheets;
end;

procedure TXLSX4DWorkbook.Save;
begin
  if FFilePath = '' then
    raise Exception.Create('File pathnot set. Use SaveAs to specify a file name');

  SaveAs(FFilePath);
end;

procedure TXLSX4DWorkbook.SaveAs(const AFileName: string);
begin
  raise Exception.Create('Not implemented yet');
end;

procedure TXLSX4DWorkbook.SetActiveWorksheet(
  const AWorksheet: IXLSX4DWorksheet);
begin
  if not Assigned(AWorksheet) then
    raise Exception.Create('Cannot set nil as active worksheet');

  if FWorksheets.IndexOf(AWorksheet) = -1 then
    raise Exception.Create('Worksheet does not belong to this workbook');

  FActiveWorksheet := AWorksheet;
end;

end.
