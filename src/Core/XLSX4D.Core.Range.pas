unit XLSX4D.Core.Range;

interface

uses
  System.Generics.Collections,
  XLSX4D.Intf.Range,
  XLSX4D.Intf.Cell;

type
  TXLSX4DRange = class(TInterfacedObject, IXLSX4DRange)
  private
    FStartRow: Integer;
    FStartColumn: Integer;
    FEndRow: Integer;
    FEndColumn: Integer;
    FAddress: string;
    FCells: TList<IXLSX4DCell>;

    procedure ValidateCoordinates;
    function CellAddressToString(ARow, AColumn: Integer): string;
    function ColumnNumberToLetter(AColumn: Integer): string;
  protected
    function GetStartRow: Integer;
    function GetStartColumn: Integer;
    function GetEndRow: Integer;
    function GetEndColumn: Integer;
    function GetAddress: string;
    function GetCells: TList<IXLSX4DCell>;
  public
    constructor Create(AStartRow, AStartColumn, AEndRow, AEndColumn: Integer);
    destructor Destroy; override;

    property StartRow: Integer read GetStartRow;
    property StartColumn: Integer read GetStartColumn;
    property EndRow: Integer read GetEndRow;
    property EndColumn: Integer read GetEndColumn;
    property Address: string read GetAddress;
    property Cells: TList<IXLSX4DCell> read GetCells;

    function GetCell(ARow, AColumn: Integer): IXLSX4DCell;
    procedure SetValue(const AValue: Variant);
    procedure SetFormula(const AFormula: string);
    procedure Clear;
    function GetValues: TArray<TArray<Variant>>;
    procedure SetValues(const AValues: TArray<TArray<Variant>>);
  end;

implementation

uses
  System.SysUtils,
  System.Variants,
  System.Math;

{ TXLSX4DRange }

function TXLSX4DRange.CellAddressToString(ARow, AColumn: Integer): string;
begin
  Result := ColumnNumberToLetter(AColumn) + IntToStr(ARow);
end;

procedure TXLSX4DRange.Clear;
var
  LCell: IXLSX4DCell;
begin
  for LCell in FCells do
    LCell.Clear;
end;

function TXLSX4DRange.ColumnNumberToLetter(AColumn: Integer): string;
begin
  Result := '';
  while AColumn > 0 do
  begin
    Dec(AColumn);
    Result := Chr(Ord('A') + (AColumn mod 26)) + Result;
    AColumn := AColumn div 26;
  end;
end;

constructor TXLSX4DRange.Create(AStartRow, AStartColumn, AEndRow, AEndColumn: Integer);
begin
  inherited Create;

  FStartRow := AStartRow;
  FStartColumn := AStartColumn;
  FEndRow := AEndRow;
  FEndColumn := AEndColumn;

  ValidateCoordinates;

  FCells := TList<IXLSX4DCell>.Create;
end;

destructor TXLSX4DRange.Destroy;
begin
  FCells.Free;
  inherited;
end;

function TXLSX4DRange.GetAddress: string;
begin
  if (FStartRow = FEndRow) and (FStartColumn = FEndColumn) then
    Result := CellAddressToString(FStartRow, FStartColumn)
  else
    Result := CellAddressToString(FStartRow, FStartColumn) + ':' +
              CellAddressToString(FEndRow, FEndColumn);
end;

function TXLSX4DRange.GetCell(ARow, AColumn: Integer): IXLSX4DCell;
var
  LCell: IXLSX4DCell;
begin
  Result := nil;

  if (ARow < FStartRow) or (ARow > FEndRow) or
     (AColumn < FStartColumn) or (AColumn > FEndColumn) then
    raise Exception.CreateFmt('Cell (%d,%d) is out of range', [ARow, AColumn]);

  for LCell in FCells do
  begin
    if (LCell.Row = ARow) and (LCell.Column = AColumn) then
    begin
      Result := LCell;
      Break;
    end;
  end;

  if Result = nil then
    raise Exception.CreateFmt('Cell (%d,%d) not found in range', [ARow, AColumn]);
end;

function TXLSX4DRange.GetCells: TList<IXLSX4DCell>;
begin
  Result := FCells;
end;

function TXLSX4DRange.GetEndColumn: Integer;
begin
  Result := FEndColumn;
end;

function TXLSX4DRange.GetEndRow: Integer;
begin
  Result := FEndRow;
end;

function TXLSX4DRange.GetStartColumn: Integer;
begin
  Result := FStartColumn;
end;

function TXLSX4DRange.GetStartRow: Integer;
begin
  Result := FStartRow;
end;

function TXLSX4DRange.GetValues: TArray<TArray<Variant>>;
var
  LRowCount, LColCount: Integer;
  LRow, LCol: Integer;
  LCell: IXLSX4DCell;
begin
  LRowCount := FEndRow - FStartRow + 1;
  LColCount := FEndColumn - FStartColumn + 1;

  SetLength(Result, LRowCount);
  for LRow := 0 to LRowCount - 1 do
    SetLength(Result[LRow], LColCount);

  for LRow := FStartRow to FEndRow do
    for LCol := FStartColumn to FEndColumn do
    begin
      try
        LCell := GetCell(LRow, LCol);
        if Assigned(LCell) then
          Result[LRow - FStartRow][LCol - FStartColumn] := LCell.Value
        else
          Result[LRow - FStartRow][LCol - FStartColumn] := Null;
      except
        Result[LRow - FStartRow][LCol - FStartColumn] := Null;
      end;
    end;
end;

procedure TXLSX4DRange.SetFormula(const AFormula: string);
var
  LCell: IXLSX4DCell;
begin
  for LCell in FCells do
    LCell.Formula := AFormula;
end;

procedure TXLSX4DRange.SetValue(const AValue: Variant);
var
  LCell: IXLSX4DCell;
begin
  for LCell in FCells do
    LCell.Value := AValue;
end;

procedure TXLSX4DRange.SetValues(const AValues: TArray<TArray<Variant>>);
var
  LRow, LCol: Integer;
  LCell: IXLSX4DCell;
  LMaxRow, LMaxCol: Integer;
begin
  if Length(AValues) = 0 then
    Exit;

  LMaxRow := Min(Length(AValues), FEndRow - FStartRow + 1);

  for LRow := 0 to LMaxRow - 1 do
  begin
    if Length(AValues[LRow]) > 0 then
    begin
      LMaxCol := Min(Length(AValues[LRow]), FEndColumn - FStartColumn + 1);

      for LCol := 0 to LMaxCol - 1 do
      begin
        try
          LCell := GetCell(FStartRow + LRow, FStartColumn + LCol);
          if Assigned(LCell) then
            LCell.Value := AValues[LRow][LCol];
        except
          // ignore not found cell
        end;
      end;
    end;
  end;
end;

procedure TXLSX4DRange.ValidateCoordinates;
begin
  if (FStartRow < 1) or (FStartColumn < 1) then
    raise Exception.Create('Start coordinates must be >= 1');

  if (FEndRow < FStartRow) or (FEndColumn < FStartColumn) then
    raise Exception.Create('End coordinates must be >= start coordinates');
end;

end.

