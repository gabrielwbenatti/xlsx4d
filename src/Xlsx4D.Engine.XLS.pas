unit Xlsx4D.Engine.XLS;

interface

uses
  System.Classes,
  Xlsx4D.Types;

type
  TXLSEngine = class
  private
    FStream: TFileStream;
    FWorksheets: TWorksheets;
    FCurrentWorksheet: TWorksheet;
    FSSTStrings: TStringList;

    // BIFF8 record type constants
    const
      BIFF_BOF = $0809;
      BIFF_EOF = $000A;
      BIFF_BOUNDSHEET = $0085;
      BIFF_LABEL = $0204;
      BIFF_LABELSST = $00FD;
      BIFF_NUMBER = $0203;
      BIFF_RK = $027E;
      BIFF_MULRK = $00BD;
      BIFF_SST = $00FC;

    function ReadWord: Word;
    function ReadDWord: Cardinal;
    function ReadDouble: Double;
    procedure ReadSST;
    procedure ReadBoundSheets;
    procedure ReadWorksheetData;
    function DecoreRK(const ARK: Cardinal): Double;
  public
    constructor Create;
    destructor Destroy; override;

    function LoadFromFile(const AFileName: string): TWorksheets;
  end;

implementation

{ TXLSEngine }

constructor TXLSEngine.Create;
begin

end;

function TXLSEngine.DecoreRK(const ARK: Cardinal): Double;
begin

end;

destructor TXLSEngine.Destroy;
begin

  inherited;
end;

function TXLSEngine.LoadFromFile(const AFileName: string): TWorksheets;
begin
  raise EXlsx4DException.Create('Not implemented yet');
end;

procedure TXLSEngine.ReadBoundSheets;
begin

end;

function TXLSEngine.ReadDouble: Double;
begin

end;

function TXLSEngine.ReadDWord: Cardinal;
begin

end;

procedure TXLSEngine.ReadSST;
begin

end;

function TXLSEngine.ReadWord: Word;
begin

end;

procedure TXLSEngine.ReadWorksheetData;
begin

end;

end.
