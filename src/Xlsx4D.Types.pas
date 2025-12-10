unit Xlsx4D.Types;

interface

uses
  System.Generics.Collections,
  System.SysUtils,
  System.Variants;

type
  TCellType = (ctEmpty, ctString, ctNumber, ctBoolean, ctDate, ctFormula, ctError);

  TCell = record
    Row: Integer;
    Col: Integer;
    Value: Variant;
    CellType: TCellType;
    FormattedValue: string;
    Formula: string;

    class function Empty(ARow, ACol: Integer): TCell; static;
    class function Create(ARow, ACol: Integer; const AValue: Variant; ACellType: TCellType = ctString): TCell; static;
    class function CreateWithFormula(ARow, ACol: Integer; const AValue: Variant; const AFormula: string): TCell; static;
    function IsEmpty: Boolean;
    function AsString: string;
    function AsInteger: Integer;
    function AsFloat: Double;
    function AsBoolean: Boolean;
    function AsDateTime: TDateTime;
    function HasFormula: Boolean;
  end;

  TRowCells = TList<TCell>;

  TWorksheet = class
  private
    FName: string;
    FRows: TObjectList<TRowCells>;
    FMaxRow: Integer;
    FMaxCol: Integer;

    function GetCell(ARow, ACol: Integer): TCell;
    procedure SetCell(ARow, ACol: Integer; const Value: TCell);
    function GetCellByRef(ARow: Integer; const AColRef: string): TCell;
    procedure SetCellByRef(ARow: Integer; const AColRef: string; const Value: TCell);
    function GetRowCount: Integer;
    function GetColCount: Integer;
    class function ColRefToNumber(const AColRef: string): Integer; static;
  public
    constructor Create(const AName: string);
    destructor Destroy; override;

    procedure Clear;
    procedure AddCell(const ACell: TCell);
    function GetRow(ARow: Integer): TRowCells;

    property Name: string read FName write FName;
    property Cells[ARow, ACol: Integer]: TCell read GetCell write SetCell; default;
    property CellsByRef[ARow: Integer; const AColRef: string]: TCell read GetCellByRef write SetCellByRef;
    property RowCount: Integer read GetRowCount;
    property ColCount: Integer read GetColCount;
  end;

  TWorksheets = class(TObjectList<TWorksheet>)
  public
    function FindByName(const AName: string): TWorksheet;
  end;

  EXlsx4DException = class(Exception);

implementation

{ TCell }

function TCell.AsBoolean: Boolean;
begin
  if IsEmpty then
    Result := False
  else
  try
    Result := VarAsType(Value, varBoolean)
  except
    Result := False;
  end;
end;

function TCell.AsDateTime: TDateTime;
begin
  if IsEmpty then
    Result := 0
  else
  try
    if VarIsNumeric(Value) then
      Result := Double(Value) + EncodeDate(1899, 12, 30)
    else
      Result := StrToDateTime(VarToStr(Value));
  except
    Result := 0;
  end;
end;

function TCell.AsFloat: Double;
begin
  if IsEmpty then
    Result := 0.00
  else
  try
    Result := VarAsType(Value, varDouble);
  except
    Result := 0.00;
  end;
end;

function TCell.AsInteger: Integer;
begin
  if IsEmpty then
    Result := 0
  else
    Result := StrToIntDef(VarToStrDef(Value, '0'), 0);
end;

function TCell.AsString: string;
begin
  if IsEmpty then
    Result := ''
  else
    Result := VarToStrDef(Value, '');
end;

class function TCell.Create(ARow, ACol: Integer; const AValue: Variant; ACellType: TCellType): TCell;
begin
  Result.Row := ARow;
  Result.Col := ACol;
  Result.Value := AValue;
  Result.CellType := ACellType;
  Result.FormattedValue := VarToStrDef(AValue, '');
  Result.Formula := '';
end;

class function TCell.CreateWithFormula(ARow, ACol: Integer;
  const AValue: Variant; const AFormula: string): TCell;
begin
  Result.Row := ARow;
  Result.Col := ACol;
  Result.Value := AValue;
  Result.CellType := ctFormula;
  Result.FormattedValue := VarToStrDef(AValue, '');
  Result.Formula := AFormula;
end;

class function TCell.Empty(ARow, ACol: Integer): TCell;
begin
  Result := TCell.Create(ARow, ACol, Null, ctEmpty);
end;

function TCell.HasFormula: Boolean;
begin
  Result := (Formula <> '') and (CellType <> ctFormula);
end;

function TCell.IsEmpty: Boolean;
begin
  Result := (CellType = ctEmpty) or VarIsNull(Value) or VarIsEmpty(Value);
end;


{ TWorksheet }

procedure TWorksheet.AddCell(const ACell: TCell);
var
  RowList: TRowCells;
  I: Integer;
begin
  while FRows.Count < ACell.Row do
    FRows.Add(TRowCells.Create);

  RowList := FRows[ACell.Row - 1];

  for I := 0 to RowList.Count - 1 do
  begin
    if RowList[I].Col = ACell.Col then
    begin
      RowList[I] := ACell;
      Exit;
    end;
  end;

  RowList.Add(ACell);

  if ACell.Row > FMaxRow then
    FMaxRow := ACell.Row;
  if ACell.Col > FMaxCol then
    FMaxCol := ACell.Col;
end;

procedure TWorksheet.Clear;
begin
  FRows.Clear;
  FMaxRow := 0;
  FMaxCol := 0;
end;

constructor TWorksheet.Create(const AName: string);
begin
  inherited Create;
  FName := AName;
  FRows := TObjectList<TRowCells>.Create(True);
  FMaxRow := 0;
  FMaxCol := 0;
end;

destructor TWorksheet.Destroy;
begin
  FRows.Free;
  inherited;
end;

function TWorksheet.GetCell(ARow, ACol: Integer): TCell;
var
  RowList: TRowCells;
  I: Integer;
begin
  Result := TCell.Empty(ARow, ACol);

  if (ARow < 1) or (ARow > FRows.Count) then
    Exit;

  RowList := FRows[ARow - 1];
  for I := 0 to RowList.Count - 1 do
  begin
    if RowList[I].Col = ACol then
    begin
      Result := RowList[I];
      Exit;
    end;
  end;
end;

function TWorksheet.GetColCount: Integer;
begin
  Result := FMaxCol;
end;

function TWorksheet.GetRow(ARow: Integer): TRowCells;
begin
  if (ARow < 1) or (ARow > FRows.Count) then
    Result := nil
  else
    Result := FRows[ARow - 1];
end;

function TWorksheet.GetRowCount: Integer;
begin
  Result := FMaxRow;
end;

procedure TWorksheet.SetCell(ARow, ACol: Integer; const Value: TCell);
begin
  AddCell(Value);
end;

class function TWorksheet.ColRefToNumber(const AColRef: string): Integer;
var
  I: Integer;
  C: Char;
begin
  Result := 0;
  for I := 1 to Length(AColRef) do
  begin
    C := UpCase(AColRef[I]);
    if not (C in ['A'..'Z']) then
      raise EXlsx4DException.CreateFmt('Invalid column reference: %s', [AColRef]);
    Result := Result * 26 + (Ord(C) - Ord('A') + 1);
  end;
end;

function TWorksheet.GetCellByRef(ARow: Integer; const AColRef: string): TCell;
begin
  Result := GetCell(ARow, ColRefToNumber(AColRef));
end;

procedure TWorksheet.SetCellByRef(ARow: Integer; const AColRef: string; const Value: TCell);
begin
  SetCell(ARow, ColRefToNumber(AColRef), Value);
end;

{ TWorksheets }

function TWorksheets.FindByName(const AName: string): TWorksheet;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to Count - 1 do
  begin
    if SameText(Items[I].Name, AName) then
    begin
      Result := Items[I];
      Exit;
    end;
  end;
end;

end.

