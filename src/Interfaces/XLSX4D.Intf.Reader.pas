unit XLSX4D.Intf.Reader;

interface

uses
  System.Classes,
  XLSX4D.Intf.Workbook;

type
  IXLSX4DReader = interface
    ['{EFE38B66-1AAA-4AD0-BCCD-3593ECF0EC96}']

    // FIle operations
    function ReadFile(const AFileName: string): IXLSX4DWorkbook;
    function ReadStream(AStream: TStream): IXLSX4DWorkbook;

    // Getters
    function GetReadFormulas: Boolean;
    function GetReadFormats: Boolean;
    function GetReadComments: Boolean;

    // Setters
    procedure SetReadFormulas(AValue: Boolean);
    procedure SetReadFormats(AValue: Boolean);
    procedure SetReadComments(AValue: Boolean);

    // Properties
    property ReadFormulas: Boolean read GetReadFormulas write SetReadFormulas;
    property ReadFormats: Boolean read GetReadFormats write SetReadFormats;
    property ReadComments: Boolean read GetReadComments write SetReadComments;
  end;

implementation

end.

