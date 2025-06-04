unit XLSX4D.Intf.Range;

interface

uses
  System.Generics.Collections,
  XLSX4D.Intf.Cell;

type
  IXLSX4DRange = interface
    ['{AF83B0FB-120B-436F-8F00-D9D8A4A460BC}']

    // Getters
    function GetStartRow: Integer;
    function GetStartColumn: Integer;
    function GetEndRow: Integer;
    function GetEndColumn: Integer;
    function GetAddress: string;
    function GetCells: TList<IXLSX4DCell>;

    // Properties
    property StartRow: Integer read GetStartRow;
    property StartColumn: Integer read GetStartColumn;
    property EndRow: Integer read GetEndRow;
    property EndColumn: Integer read GetEndColumn;
    property Address: string read GetAddress;
    property Cells: TList<IXLSX4DCell> read GetCells;

    // Methods
    function GetCell(ARow, AColumn: Integer): IXLSX4DCell;
    procedure SetValue(const AValue: Variant);
    procedure SetFormula(const AFormula: string);
    procedure Clear;
    function GetValues: TArray<TArray<Variant>>;
    procedure SetValues(const AValues: TArray<TArray<Variant>>);
  end;

implementation

end.

