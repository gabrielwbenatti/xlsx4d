unit XLSX4D.Core.Cell;

interface

uses
  XLSX4D.Intf.Cell,
  XLSX4D.Types;

type
  TXLSX4DCell = class(TInterfacedObject, IXLSX4DCell)
  private
    FRow: Integer;
    FColumn: Integer;
    FValue: Variant;
    FFormula: string;
    FFormat: string;
    FDataType: TXLSX4DCellDataType;

    function DeterminateDataType(const AValue: Variant): TXLSX4DCellDataType;
    function ColumnNumberToLetter(AColumn: Integer): string;
    procedure UpdateDataType;
  protected
    function GetRow: Integer;
    function GetColumn: Integer;
    function GetAddress: string;
    function GetValue: Variant;
    function GetFormula: string;
    function GetDataType: TXLSX4DCellDataType;
    function GetFormat: string;

    procedure SetValue(const AValue: Variant);
    procedure SetFormula(const AValue: string);
    procedure SetFormat(const AValue: string);
  public
    constructor Create(ARow, AColumn: Integer);

    property Row: Integer read GetRow;
    property Column: Integer read GetColumn;
    property Address: string read GetAddress;
    property Value: Variant read GetValue write SetValue;
    property Formula: string read GetFormula write SetFormula;
    property DataType: TXLSX4DCellDataType read GetDataType;
    property Format: string read GetFormat write SetFormat;

    procedure Clear;
    function IsEmpty: Boolean;
    function HasFormula: Boolean;
  end;

implementation

uses
  System.SysUtils,
  System.Variants;

{ TXLSX4DCell }

procedure TXLSX4DCell.Clear;
begin
  FValue := Null;
  FFormula := '';
  FFormat := '';
  FDataType := ctString;
end;

function TXLSX4DCell.ColumnNumberToLetter(AColumn: Integer): string;
var
  Temp: Integer;
begin
  Result := '';
  Temp := AColumn;

  while Temp > 0 do
  begin
    Dec(Temp);
    Result := Chr(Ord('A') + (Temp mod 26)) + Result;
    Temp := Temp div 26;
  end;
end;

constructor TXLSX4DCell.Create(ARow, AColumn: Integer);
begin
  inherited Create;
  FRow := ARow;
  FColumn := AColumn;
  FValue := Null;
  FFormula := '';
  FFormat := '';
  FDataType := ctString;
end;

function TXLSX4DCell.DeterminateDataType(
  const AValue: Variant): TXLSX4DCellDataType;
var
  LVarType: Integer;
begin
  if VarIsNull(AValue) or VarIsEmpty(AValue) then
  begin
    result := ctString;
    Exit;
  end;

  LVarType := VarType(AValue);

  case LVarType of
    varSmallint, varInteger, varSingle, varDouble, varCurrency, varInt64:
      Result := ctNumber;
    varDate:
      Result := ctDate;
    varBoolean:
      Result := ctBoolean;
    varString, varUString, varOleStr:
      begin
        if (VarToStr(AValue) = '#NULL!') or
           (VarToStr(AValue) = '#DIV/0!') or
           (VarToStr(AValue) = '#VALUE!') or
           (VarToStr(AValue) = '#REF!') or
           (VarToStr(AValue) = '#NAME?') or
           (VarToStr(AValue) = '#NUM!') or
           (VarToStr(AValue) = '#N/A') then
          Result := ctError
        else
          Result := ctString;
      end;
  else
    Result := ctString;
  end;
end;

function TXLSX4DCell.GetAddress: string;
begin
  Result := ColumnNumberToLetter(FColumn) + IntToStr(FRow);
end;

function TXLSX4DCell.GetColumn: Integer;
begin
  Result := FColumn;
end;

function TXLSX4DCell.GetDataType: TXLSX4DCellDataType;
begin
  Result := FDataType;
end;

function TXLSX4DCell.GetFormat: string;
begin
  Result := FFormat;
end;

function TXLSX4DCell.GetFormula: string;
begin
  Result := FFormula;
end;

function TXLSX4DCell.GetRow: Integer;
begin
  Result := FRow;
end;

function TXLSX4DCell.GetValue: Variant;
begin
  Result := FValue;
end;

function TXLSX4DCell.HasFormula: Boolean;
begin
  Result := FFormula <> '';
end;

function TXLSX4DCell.IsEmpty: Boolean;
begin
  Result := VarIsNull(FValue) and (FFormula = '');
end;

procedure TXLSX4DCell.SetFormat(const AValue: string);
begin
  FFormat := AValue;
end;

procedure TXLSX4DCell.SetFormula(const AValue: string);
begin
  FFormula := AValue;
  if AValue <> '' then
  begin
    FDataType := ctFormula;
    FValue := Null;
  end;
end;

procedure TXLSX4DCell.SetValue(const AValue: Variant);
begin
  FValue := AValue;
  FFormula := '';
  UpdateDataType;
end;

procedure TXLSX4DCell.UpdateDataType;
begin
  if FFormula <> '' then
    FDataType := ctFormula
  else
    FDataType := DeterminateDataType(FValue);
end;

end.

