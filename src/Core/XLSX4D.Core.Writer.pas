unit XLSX4D.Core.Writer;

interface

uses
  System.Classes,
  XLSX4D.Intf.Writer,
  XLSX4D.Types,
  XLSX4D.Intf.Workbook;

type
  TXLSX4DWriter = class(TInterfacedObject, IXLSX4DWriter)
  private
    FReadFormulas: Boolean;
    FReadFormats: Boolean;
    FReadComments: Boolean;
    FCompressionLevel: TXLSX4DCompressionLevel;

    procedure ProcessWorkbook(const AWorkbook: IXLSX4DWorkbook; AOutput: TStream);
    procedure WriteSharedStrings(AOutput: TStream);
    procedure WriteStyles(AOutput: TStream);
    procedure WriteWorksheets(const AWorkbook: IXLSX4DWorkbook; AOutput: TStream);
    procedure WriteWorkbookXML(const AWorkbook: IXLSX4DWorkbook; AOutput: TStream);
  protected
    procedure WriteFile(const AWorkbook: IXLSX4DWorkbook; const AFileName: string);
    procedure WriteStream(const AWorkbook: IXLSX4DWorkbook; AStream: TStream);

    function GetReadFormulas: Boolean;
    function GetReadFormats: Boolean;
    function GetReadComments: Boolean;
    function GetCompressionLevel: TXLSX4DCompressionLevel;

    procedure SetReadFormulas(AValue: Boolean);
    procedure SetReadFormats(AValue: Boolean);
    procedure SetReadComments(AValue: Boolean);
    procedure SetCompressionLevel(AValue: TXLSX4DCompressionLevel);
  public
    constructor Create;
    destructor Destroy; override;

    property ReadFormulas: Boolean read GetReadFormulas write SetReadFormulas;
    property ReadFormats: Boolean read GetReadFormats write SetReadFormats;
    property ReadComments: Boolean read GetReadComments write SetReadComments;
    property CompressionLevel: TXLSX4DCompressionLevel read GetCompressionLevel write SetCompressionLevel;

  end;

implementation

uses
  System.SysUtils,
  System.Zip;

{ TXLSX4DWriter }

constructor TXLSX4DWriter.Create;
begin
  inherited Create;

  FReadFormulas := False;
  FReadFormats := False;
  FReadComments := False;
  FCompressionLevel := clDefault;
end;

destructor TXLSX4DWriter.Destroy;
begin
  inherited;
end;

function TXLSX4DWriter.GetCompressionLevel: TXLSX4DCompressionLevel;
begin
  Result := FCompressionLevel;
end;

function TXLSX4DWriter.GetReadComments: Boolean;
begin
  Result := FReadComments;
end;

function TXLSX4DWriter.GetReadFormats: Boolean;
begin
  Result := FReadFormats;
end;

function TXLSX4DWriter.GetReadFormulas: Boolean;
begin
  Result := FReadFormulas;
end;

procedure TXLSX4DWriter.ProcessWorkbook(const AWorkbook: IXLSX4DWorkbook;
  AOutput: TStream);
var
  ZipFile: TZipFile;
  TempStream: TMemoryStream;
begin
  ZipFile := TZipFile.Create;
  TempStream := TMemoryStream.Create;
  try
    ZipFile.Open(AOutput, zmWrite);

    TempStream.Clear;
    // TODO: Implement content types generation
    ZipFile.Add(TempStream, '[Content_Types].xml');

    TempStream.Clear;
    // TODO: Implement relationships generation
    ZipFile.Add(TempStream, '_rels/.rels');

    TempStream.Clear;
    // TODO: Implement app properties generation
    ZipFile.Add(TempStream, 'docProps/app.xml');;

    TempStream.Clear;
    // TODO: Implement core properties generation
    ZipFile.Add(TempStream, 'docProps/core.xml');;

    TempStream.Clear;
    WriteWorkbookXML(AWorkbook, TempStream);
    ZipFile.Add(TempStream, 'xl/workbook.xml');;

    TempStream.Clear;
    // TODO: Implement workbook relationships generation
    ZipFile.Add(TempStream, 'xl/_rels/workbook.xml.rels');

    if FReadFormats then
    begin
      TempStream.Clear;
      WriteSharedStrings(TempStream);
      ZipFile.Add(TempStream, 'xl/sharedStrings.xml');
    end;

    if FReadFormats then
    begin
      TempStream.Clear;
      WriteSharedStrings(TempStream);
      ZipFile.Add(TempStream, 'xl/styles.xml');
    end;

    WriteWorksheets(AWorkbook, AOutput);
  finally
    ZipFile.Free;
    TempStream.Free;
  end;
end;

procedure TXLSX4DWriter.SetCompressionLevel(AValue: TXLSX4DCompressionLevel);
begin
  FCompressionLevel := AValue;
end;

procedure TXLSX4DWriter.SetReadComments(AValue: Boolean);
begin
  FReadComments := AValue;
end;

procedure TXLSX4DWriter.SetReadFormats(AValue: Boolean);
begin
  FReadFormats := AValue;
end;

procedure TXLSX4DWriter.SetReadFormulas(AValue: Boolean);
begin
  FReadFormulas := AValue;
end;

procedure TXLSX4DWriter.WriteFile(const AWorkbook: IXLSX4DWorkbook;
  const AFileName: string);
var
  FileStream: TFileStream;
begin
  if not Assigned(AWorkbook) then
    raise Exception.Create('Workbook cannot be nil');

  if Trim(AFileName) = '' then
    raise Exception.Create('Filename cannot be empty');

  FileStream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteStream(AWorkbook, FileStream);
  finally
    FileStream.Free;
  end;
end;

procedure TXLSX4DWriter.WriteSharedStrings(AOutput: TStream);
var
  XMLContent: string;
  StringStream: TStringStream;
begin
  XMLContent := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + sLineBreak +
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + sLineBreak +
                '</sst>';

  StringStream := TStringStream.Create(XMLContent, TEncoding.UTF8);
  try
    AOutput.CopyFrom(StringStream, 0)
  finally
    StringStream.Free;
  end;
end;

procedure TXLSX4DWriter.WriteStream(const AWorkbook: IXLSX4DWorkbook;
  AStream: TStream);
begin
  if not Assigned(AWorkbook) then
    raise Exception.Create('Workbook cannot be nil');

  if not Assigned(AStream) then
    raise Exception.Create('Stream cannot be nil');

  try
    ProcessWorkbook(AWorkbook, AStream);
  except
    on E: Exception do
      raise Exception.CreateFmt('Error writing workbook to stream: %s', [E.Message]);
  end;
end;

procedure TXLSX4DWriter.WriteStyles(AOutput: TStream);
begin

end;

procedure TXLSX4DWriter.WriteWorkbookXML(const AWorkbook: IXLSX4DWorkbook;
  AOutput: TStream);
begin

end;

procedure TXLSX4DWriter.WriteWorksheets(const AWorkbook: IXLSX4DWorkbook;
  AOutput: TStream);
begin

end;

end.

