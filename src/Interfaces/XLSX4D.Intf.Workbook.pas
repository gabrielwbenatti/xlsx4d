unit XLSX4D.Intf.Workbook;

interface

uses
  System.Generics.Collections,
  XLSX4D.Intf.Worksheet,
  XLSX4D.Types;

type
  IXLSX4DWorkbook = interface
    ['{E0C8E8A3-71C7-4988-AC81-E9E100AD2371}']

    // Getters
    function GetWorksheetCount: Integer;
    function GetWorksheets: TList<IXLSX4DWorksheet>;
    function GetActiveWorksheet: IXLSX4DWorksheet;
    function GetFilePath: string;

    // Setters
    procedure SetActiveWorksheet(const AWorksheet: IXLSX4DWorksheet);

    // Properties
    property WorksheetCount: Integer read GetWorksheetCount;
    property Worksheets: TList<IXLSX4DWorksheet> read GetWorksheets;
    property ActiveWorksheet: IXLSX4DWorksheet read GetActiveWorksheet write SetActiveWorksheet;
    property FilePath: string read GetFilePath;

    // Worksheet operations
    function AddWorksheet(const AName: string = ''): IXLSX4DWorksheet;
    function GetWorksheet(AIndex: Integer): IXLSX4DWorksheet; overload;
    function GetWorksheet(const AName: string = ''): IXLSX4DWorksheet; overload;
    procedure DeleteWorksheet(AIndex: Integer); overload;
    procedure DeleteWorksheet(const AName: string = ''); overload;

    // File operations
    procedure SaveAs(const AFileName: string);
    procedure Save;
    procedure Close;
  end;

implementation

end.

