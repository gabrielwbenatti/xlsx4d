unit XLSX4D.Intf.Worksheet;

interface

uses
  System.Generics.Collections,
  XLSX4D.Intf.Cell,
  XLSX4D.Intf.Range;

type
  IXLSX4DWorksheet = interface
    ['{7361A08D-1C56-4C84-ADEF-C89DA83C7477}']

    // Getters
    function GetName: string;
    function GetIndex: Integer;
    function GetRowCount: Integer;
    function GetColumnCount: Integer;
    function GetUsedRange: IXLSX4DRange;

    // Setters
    procedure SetName(const AValue: string);

    // Properties
    property Name: string read GetName write SetName;
    property Index: Integer read GetIndex;
    property RowCount: Integer read GetRowCount;
    property ColumnCount: Integer read GetColumnCount;
    property UsedRange: IXLSX4DRange read GetUsedRange;

    // Cell operations
    function GetCell(ARow, AColumn: Integer): IXLSX4DCell; overload;
    function GetCell(const AAddress: string): IXLSX4DCell; overload;
    function GetRange(const AAddress: string): IXLSX4DRange; overload;
    function GetRange(AStartRow, AStartCol, AEndRow, AEndCol: Integer): IXLSX4DRange; overload;

    // Row/Column operations
    procedure InsertRow(ARow: Integer; ACount: Integer = 1);
    procedure DeleteRow(ARow: Integer; ACount: Integer = 1);
    procedure InsertColumn(AColumn: Integer; ACount: Integer = 1);
    procedure DeleteColumn(AColumn: Integer; ACount: Integer = 1);

    // Data operations
    procedure Clear;
    function Find(const AValue: Variant): IXLSX4DCell;
    function FindAll(const AValue: Variant): TList<IXLSX4DCell>;
  end;

implementation

end.

