unit XLSX4D.Core.Worksheet;

interface

uses
  System.Generics.Collections,
  XLSX4D.Intf.Worksheet,
  XLSX4D.Intf.Range,
  XLSX4D.Intf.Cell,
  XLSX4D.Core.Range,
  XLSX4D.Core.Cell;

type
  TXLSX4DWorksheet = class(TInterfacedObject, IXLSX4DWorksheet)
  private
    FName: string;
    FIndex: Integer;
    FCells: TObjectDictionary<string, TXLSX4DCell>;
    FRowCount: Integer;
    FColumnCount: Integer;

    function CellAddressToCoordinates(const AAddress: string; out ARow, AColumn: Integer): Boolean;
    function CoordinatesToCellAddress(ARow, AColumn: Integer): string;
    procedure UpdateDimensions(ARow, AColumn: Integer);
    function GetOrCreateCell(ARow, AColumn: Integer): TXLSX4DCell;
  protected
    function GetName: string;
    function GetIndex: Integer;
    function GetRowCount: Integer;
    function GetColumnCount: Integer;
    function GetUsedRange: IXLSX4DRange;

    procedure SetName(const AValue: string);
  public
    constructor Create(AName: string; AIndex: Integer);
    destructor Destroy; override;

    property Name: string read GetName write SetName;
    property Index: Integer read GetIndex;
    property RowCount: Integer read GetRowCount;
    property ColumnCount: Integer read GetColumnCount;
    property UsedRange: IXLSX4DRange read GetUsedRange;

    function GetCell(ARow, AColumn: Integer): IXLSX4DCell; overload;
    function GetCell(const AAddress: string): IXLSX4DCell; overload;
    function GetRange(const AAddress: string): IXLSX4DRange; overload;
    function GetRange(AStartRow, AStartCol, AEndRow, AEndCol: Integer): IXLSX4DRange; overload;

    procedure InsertRow(ARow: Integer; ACount: Integer = 1);
    procedure DeleteRow(ARow: Integer; ACount: Integer = 1);
    procedure InsertColumn(AColumn: Integer; ACount: Integer = 1);
    procedure DeleteColumn(AColumn: Integer; ACount: Integer = 1);

    procedure Clear;
    function Find(const AValue: Variant): IXLSX4DCell;
    function FindAll(const AValue: Variant): TList<IXLSX4DCell>;
  end;

implementation

uses
  System.SysUtils,
  System.RegularExpressions,
  System.Math,
  System.Variants;

{ TXLSX4DWorksheet }

function TXLSX4DWorksheet.CellAddressToCoordinates(const AAddress: string;
  out ARow, AColumn: Integer): Boolean;
var
  LRegEx: TRegEx;
  LMatch: TMatch;
  LColumnStr: string;
  LRowStr: string;
  I: Integer;
begin
  Result := False;
  ARow := 0;
  AColumn := 0;

  // regex to get cell address (eg: A1, AB123)
  LRegEx := TRegEx.Create('^([A-Z]+)(\d+)$', [roIgnoreCase]);
  LMatch := LRegEx.Match(AAddress);

  if LMatch.Success then
  begin
    LColumnStr := UpperCase(LMatch.Groups[1].Value);
    LRowStr := LMatch.Groups[2].Value;

    // convert column letter to number (eg: A=1, B=2, ..., Z=26, AA=27, etc)
    for I := 1 to Length(LColumnStr) do
    begin
      AColumn := AColumn * 26 + (Ord(LColumnStr[I]) - Ord('A') + 1);
    end;

    // convert line
    ARow := StrToIntDef(LRowStr, 0);

    Result := (ARow > 0) and (AColumn > 0);
  end;
end;

procedure TXLSX4DWorksheet.Clear;
begin
  FCells.Clear;
  FRowCount := 0;
  FColumnCount := 0;
end;

function TXLSX4DWorksheet.CoordinatesToCellAddress(ARow,
  AColumn: Integer): string;
var
  LColumn: Integer;
begin
  Result := '';
  LColumn := AColumn;

  while LColumn > 0 do
  begin
    Dec(LColumn);
    Result := Chr(Ord('A') + (LColumn mod 26)) + Result;
    LColumn := LColumn div 26;
  end;

  Result := Result + IntToStr(ARow);
end;

constructor TXLSX4DWorksheet.Create(AName: string; AIndex: Integer);
begin
  inherited Create;

  FName := AName;
  FIndex := AIndex;
  FCells := TObjectDictionary<string, TXLSX4DCell>.Create([doOwnsValues]);
  FRowCount := 0;
  FColumnCount := 0;
end;

procedure TXLSX4DWorksheet.DeleteColumn(AColumn, ACount: Integer);
begin
  raise Exception.Create('Not implemented yet');
end;

procedure TXLSX4DWorksheet.DeleteRow(ARow, ACount: Integer);
begin
  raise Exception.Create('Not implemented yet');
end;

destructor TXLSX4DWorksheet.Destroy;
begin
  FCells.Free;
  inherited;
end;

function TXLSX4DWorksheet.Find(const AValue: Variant): IXLSX4DCell; 
var 
  LCell: TXLSX4DCell;
begin
  Result := nil;

  for LCell in FCells.Values do
  begin
    if (not LCell.IsEmpty) and VarSameValue(LCell.Value, AValue) then
    begin
      Result := LCell;
      Break;
    end;
  end;
end;

function TXLSX4DWorksheet.FindAll(const AValue: Variant): TList<IXLSX4DCell>;
var 
  LCell: TXLSX4DCell;
begin                
  Result := TList<IXLSX4DCell>.Create;

  for LCell in FCells.Values do
  begin
    if (not LCell.IsEmpty) and VarSameValue(LCell.Value, AValue) then
      Result.Add(LCell)
  end;                 
end;

function TXLSX4DWorksheet.GetCell(ARow, AColumn: Integer): IXLSX4DCell;
begin
  if (ARow < 1) or (AColumn < 1) then
    raise Exception.CreateFmt('Invalid coordinates: Row=%d, Column=%d', [ARow, AColumn]);

  Result := GetOrCreateCell(ARow, AColumn);
end;

function TXLSX4DWorksheet.GetCell(const AAddress: string): IXLSX4DCell;
var
  LRow, LColumn: Integer;
begin
  if not CellAddressToCoordinates(AAddress, LRow, LColumn) then
    raise Exception.CreateFmt('Invalid cell address: %s', [AAddress]);

  Result := GetCell(LRow, LColumn);
end;

function TXLSX4DWorksheet.GetColumnCount: Integer;
begin
  Result := FColumnCount;
end;

function TXLSX4DWorksheet.GetIndex: Integer;
begin
  Result := FIndex;
end;

function TXLSX4DWorksheet.GetName: string;
begin
  Result := FName;
end;

function TXLSX4DWorksheet.GetOrCreateCell(ARow, AColumn: Integer): TXLSX4DCell;
var
  LAddress: string;
  LCell: TXLSX4DCell;
begin
  LAddress := CoordinatesToCellAddress(ARow, AColumn);

  if not FCells.TryGetValue(LAddress, LCell) then
  begin
    LCell := TXLSX4DCell.Create(ARow, AColumn);
    FCells.Add(LAddress, LCell);
    UpdateDimensions(ARow, AColumn);
  end;

  Result := LCell;
end;

function TXLSX4DWorksheet.GetRange(const AAddress: string): IXLSX4DRange;
var
  LParts: TArray<string>;
  LStartRow, LStartCol, LEndRow, LEndCol: Integer;
begin
  // supports "A1:B2" or "A1"
  LParts := AAddress.Split([':']);

  if Length(LParts) = 1 then
  begin
    // unique cell
    if not CellAddressToCoordinates(LParts[0], LStartRow, LStartCol) then
      raise Exception.CreateFmt('Invalid cell address: %s', [LParts[0]]);

    LEndRow := LStartRow;
    LEndCol := LStartCol;
  end
  else if Length(LParts) = 2 then
  begin
    // cell range
    if not CellAddressToCoordinates(LParts[0], LStartRow, LStartCol) then
      raise Exception.CreateFmt('Invalid starting cell address: %s', [LParts[0]]);

    if not CellAddressToCoordinates(LParts[1], LEndRow, LEndCol) then
      raise Exception.CreateFmt('Invalid ending cell address: %s', [LParts[1]]);
  end
  else
    raise Exception.CreateFmt('Invalid range forma: %s', [AAddress]);

  Result := GetRange(LStartRow, LStartCol, LEndRow, LEndCol);
end;

function TXLSX4DWorksheet.GetRange(AStartRow, AStartCol, AEndRow,
  AEndCol: Integer): IXLSX4DRange;
begin
  if (AStartRow < 1) or (AStartCol < 1) or (AEndRow < 1) or (AEndCol < 1) then
    raise Exception.Create('Coordinates must be greater than zero');

  if AStartRow > AEndRow then
    raise Exception.Create('The starting line cannot be greater than ending line');
  if AStartCol > AEndCol then
    raise Exception.Create('The starting column cannot be greater than ending column');

  Result := TXLSX4DRange.Create(AStartRow, AStartCol, AEndRow, AEndCol);
end;

function TXLSX4DWorksheet.GetRowCount: Integer;
begin
  Result := FRowCount;
end;

function TXLSX4DWorksheet.GetUsedRange: IXLSX4DRange;
var
  LMinRow, LMaxRow, LMinCol, LMaxCol: Integer;
  LCell: TXLSX4DCell;
begin
  LMinRow := MaxInt;
  LMaxRow := 0;
  LMinCol := MaxInt;
  LMaxCol := 0;

  for LCell in FCells.Values do
  begin
    if not LCell.IsEmpty then
    begin
      LMinRow := Min(LMinRow, LCell.Row);
      LMaxRow := Max(LMaxRow, LCell.Row);
      LMinCol := Min(LMinCol, LCell.Column);
      LMaxCol := Max(LMaxCol, LCell.Column);
    end;
  end;

  if LMinRow = MaxInt then
  begin
    LMinRow := 1;
    LMaxRow := 1;
    LMinCol := 1;
    LMaxCol := 1;
  end;

  Result := GetRange(LMinRow, LMinCol, LMaxRow, LMaxCol);
end;

procedure TXLSX4DWorksheet.InsertColumn(AColumn, ACount: Integer);
begin
  raise Exception.Create('Not implemented yet');
end;

procedure TXLSX4DWorksheet.InsertRow(ARow, ACount: Integer);
var
  LCellsToMove: TList<TPair<string, TXLSX4DCell>>;
  LPair: TPair<string, TXLSX4DCell>;
  LNewAddress: string;
  LCell: TXLSX4DCell;
begin
  if ARow < 1 then
    raise Exception.Create('Line number must be greater than zero');
  if ACount < 1 then
    raise Exception.Create('Lines count must be greater than zero');

  LCellsToMove := TList<TPair<string, TXLSX4DCell>>.Create;
  try
    for LPair in FCells do
    begin
      if LPair.Value.Row >= ARow then
        LCellsToMove.Add(LPair);
    end;

    for LPair in LCellsToMove do
    begin
      FCells.Remove(LPair.Key);
      LCell := LPair.Value;
      LCell.UpdatePosition(LCell.Row + ACount, LCell.Column); 
      LNewAddress := CoordinatesToCellAddress(LCell.Row, LCell.Column);
      FCells.Add(LNewAddress, LCell);
    end;

    Inc(FRowCount, ACount);
  finally
    LCellsToMove.Free;
  end;
end;

procedure TXLSX4DWorksheet.SetName(const AValue: string);
begin
  if Trim(AValue) = '' then
    raise Exception.Create('Sheet name cannot be empty');
  FName := AValue;
end;

procedure TXLSX4DWorksheet.UpdateDimensions(ARow, AColumn: Integer);
begin
  if ARow > FRowCount then
    FRowCount := ARow;
  if AColumn > FColumnCount then
    FColumnCount := AColumn;
end;

end.

