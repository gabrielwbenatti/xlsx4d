unit XLSX4D.Intf.Cell;

interface

uses
  XLSX4D.Types;

type
  IXLSX4DCell = interface
    ['{F3DA9CB7-D274-4A9A-9ECB-4DDD89FB1E55}']

    // Getters
    function GetRow: Integer;
    function GetColumn: Integer;
    function GetAddress: string;
    function GetValue: Variant;
    function GetFormula: string;
    function GetDataType: TXLSX4DCellDataType;
    function GetFormat: string;

    // Setters
    procedure SetValue(const AValue: Variant);
    procedure SetFormula(const AValue: string);
    procedure SetFormat(const AValue: string);

    // Properties
    property Row: Integer read GetRow;
    property Column: Integer read GetColumn;
    property Address: string read GetAddress;
    property Value: Variant read GetValue write SetValue;
    property Formula: string read GetFormula write SetFormula;
    property DataType: TXLSX4DCellDataType read GetDataType;
    property Format: string read GetFormat write SetFormat;

    // Methods
    procedure Clear;
    function IsEmpty: Boolean;
    function HasFormula: Boolean;
  end;

implementation

end.

