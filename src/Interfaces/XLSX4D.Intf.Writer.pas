unit XLSX4D.Intf.Writer;

interface

uses
  System.Classes,
  XLSX4D.Intf.Workbook,
  XLSX4D.Types;

type
  IXLSX4DWriter = interface
    ['{E7AAC853-B917-4BB9-876D-678A1C581812}']

    // File operations
    procedure WriteFile(const AWorkbook: IXLSX4DWorkbook; const AFileName: string);
    procedure WriteStream(const AWorkbook: IXLSX4DWorkbook; AStream: TStream);

    // Getters
    function GetReadFormulas: Boolean;
    function GetReadFormats: Boolean;
    function GetReadComments: Boolean;
    function GetCompressionLevel: TXLSX4DCompressionLevel;

    // Setters
    procedure SetReadFormulas(AValue: Boolean);
    procedure SetReadFormats(AValue: Boolean);
    procedure SetReadComments(AValue: Boolean);
    procedure SetCompressionLevel(AValue: TXLSX4DCompressionLevel);

    // Properties
    property ReadFormulas: Boolean read GetReadFormulas write SetReadFormulas;
    property ReadFormats: Boolean read GetReadFormats write SetReadFormats;
    property ReadComments: Boolean read GetReadComments write SetReadComments;
    property CompressionLevel: TXLSX4DCompressionLevel read GetCompressionLevel write SetCompressionLevel;

  end;

implementation

end.

