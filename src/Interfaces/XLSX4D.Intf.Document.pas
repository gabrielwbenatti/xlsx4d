unit XLSX4D.Intf.Document;

interface

uses
  XLSX4D.Intf.Workbook;

type
  IXLSX4DDocument = interface
    ['{D96E517A-D506-4E25-B43B-3B7BC0D1D65C}']

    // Getters
    function GetWorkbook: IXLSX4DWorkbook;

    // Properties
    property Workbook: IXLSX4DWorkbook read GetWorkbook;

    // Methods
    function CreateNew: IXLSX4DWorkbook;
    function Open(const AFileName: string): IXLSX4DWorkbook;
    procedure Close;
  end;

implementation

end.

