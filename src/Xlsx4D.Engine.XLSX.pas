unit Xlsx4D.Engine.XLSX;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Zip,
  System.Generics.Collections,
  Xlsx4D.Types,
  Xlsx4D.XML.Parser;

type
  TXLSXEngine = class
  private
    FWorksheets: TWorksheets;
    FSharedStrings: TStringList;
    FZipFile: TZipFile;
    FTempPath: string;
    
    procedure ExtractXLSXContents(const AFileName: string);
    procedure LoadSharedStrings;
    procedure LoadWorksheets;
    procedure LoadWorksheetData(const ASheetName: string; AWorksheet: TWorksheet; ASheetIndex: Integer);
    function GetColumnIndex(const ACell: string): Integer;
    function ParseCellReference(const ACellRef: string; out ARow, ACol: Integer): Boolean;
    function GetSharedString(AIndex: Integer): string;
    procedure Cleanup;
  public
    constructor Create;
    destructor Destroy; override;
    
    function LoadFromFile(const AFileName: string): TWorksheets;
  end;

implementation

uses
  System.IOUtils,
  System.Variants,
  System.StrUtils;

{ TXLSXEngine }

constructor TXLSXEngine.Create;
begin
  inherited Create;
  FSharedStrings := TStringList.Create;
  FWorksheets := nil;
  FZipFile := TZipFile.Create;
end;

destructor TXLSXEngine.Destroy;
begin
  Cleanup;
  FSharedStrings.Free;
  if FWorksheets <> nil then
    FWorksheets.Free;
  FZipFile.Free;
  inherited;
end;

procedure TXLSXEngine.Cleanup;
begin
  if FTempPath <> '' then
  begin
    if TDirectory.Exists(FTempPath) then
      TDirectory.Delete(FTempPath, True);
    FTempPath := '';
  end;
end;

procedure TXLSXEngine.ExtractXLSXContents(const AFileName: string);
begin
  Cleanup;
  
  FTempPath := TPath.Combine(TPath.GetTempPath, 'xlsx4d_' + TGuid.NewGuid.ToString);
  TDirectory.CreateDirectory(FTempPath);
  
  FZipFile.Open(AFileName, zmRead);
  try
    FZipFile.ExtractAll(FTempPath);
  finally
    FZipFile.Close;
  end;
end;

procedure TXLSXEngine.LoadSharedStrings;
var
  SharedStringsPath: string;
  XMLContent: string;
  Parser: TXMLParser;
  Root: TXMLNode;
  SINodes: TObjectList<TXMLNode>;
  I: Integer;
  TNode: TXMLNode;
begin
  FSharedStrings.Clear;
  
  SharedStringsPath := TPath.Combine(FTempPath, 'xl\sharedStrings.xml');
  
  if not TFile.Exists(SharedStringsPath) then
    Exit;

  Parser := TXMLParser.Create;
  try
    try
      XMLContent := TFile.ReadAllText(SharedStringsPath, TEncoding.UTF8);

      Root := Parser.Parse(XMLContent);
      try
        SINodes := TXMLHelper.FindNodesByName(Root, 'si');
        try
          for I := 0 to SINodes.Count - 1 do
          begin
            // try frind <t> within <si>
            TNode := SINodes[I].FindNode('t');
            if TNode <> nil then
              FSharedStrings.Add(TNode.Value)
            else
              FSharedStrings.Add(''); // empty str
          end;
        finally
          SINodes.Free;
        end;
      finally
        Root.Free;
      end;
    except
      on E: Exception do
        raise EXlsx4DException.CreateFmt(
          'Failed to load shared strings: %s',
          [E.Message]
        );
    end;
  finally
    Parser.Free;
  end;
end;

procedure TXLSXEngine.LoadWorksheets;
var
  WorkbookPath: string;
  XMLContent: string;
  Parser: TXMLParser;
  Root: TXMLNode;
  SheetNodes: TObjectList<TXMLNode>;
  I: Integer;
  SheetName: string;
  Worksheet: TWorksheet;
begin
  WorkbookPath := TPath.Combine(FTempPath, 'xl\workbook.xml');
  
  if not TFile.Exists(WorkbookPath) then
    raise EXlsx4DException.Create('workbook.xml not found in XLSX package.');

  Parser := TXMLParser.Create;
  try
    try
      XMLContent := TFile.ReadAllText(WorkbookPath, TEncoding.UTF8);

      Root := Parser.Parse(XMLContent);
      try
        SheetNodes := TXMLHelper.FindNodesByName(Root, 'sheet');
        try
          for I := 0 to SheetNodes.Count - 1 do
          begin
            SheetName := SheetNodes[I].Attribute['name'];

            if SheetName <> '' then
            begin
              Worksheet := TWorksheet.Create(SheetName);
              FWorksheets.Add(Worksheet);

              LoadWorksheetData(SheetName, Worksheet, I + 1);
            end;
          end;
        finally
          SheetNodes.Free;
        end;
      finally
        Root.Free;
      end;
    except
      on E: Exception do
        raise EXlsx4DException.CreateFmt(
          'Failed to load worksheets: %s',
          [E.Message]
        );
    end;
  finally
    Parser.Free;
  end;
end;

procedure TXLSXEngine.LoadWorksheetData(const ASheetName: string; 
  AWorksheet: TWorksheet; ASheetIndex: Integer);
var
  SheetDataPath: string;
  XMLContent: string;
  Parser: TXMLParser;
  Root: TXMLNode;
  RowNodes, CellNodes: TObjectList<TXMLNode>;
  I, J: Integer;
  CellRef, CellType, CellValue: string;
  Row, Col: Integer;
  Cell: TCell;
  NumValue: Double;
  CellNode, ValueNode: TXMLNode;
  FormatSettings: TFormatSettings;
begin
  SheetDataPath := TPath.Combine(FTempPath, Format('xl\worksheets\sheet%d.xml', [ASheetIndex]));
  
  if not TFile.Exists(SheetDataPath) then
    Exit;

  FormatSettings := TFormatSettings.Create('en-US');
  FormatSettings.DecimalSeparator := '.';

  Parser := TXMLParser.Create;
  try
    try
      XMLContent := TFile.ReadAllText(SheetDataPath, TEncoding.UTF8);

      Root := Parser.Parse(XMLContent);
      try
        RowNodes := TXMLHelper.FindNodesByName(Root, 'row');
        try
          for I := 0 to RowNodes.Count - 1 do
          begin
            CellNodes := TObjectList<TXMLNode>.Create(False);
            try
              RowNodes[I].FindNodes('c', CellNodes);

              for J := 0 to CellNodes.Count - 1 do
              begin
                CellNode := CellNodes[J];

                CellRef := CellNode.Attribute['r'];
                CellType := CellNode.Attribute['t'];

                // find <v> node for value
                ValueNode := CellNode.FindNode('v');
                if ValueNode <> nil then
                  CellValue := ValueNode.Value
                else
                  CellValue := '';

                if not ParseCellReference(CellRef, Row, Col) then
                  Continue;

                if CellType = 's' then
                begin
                  // shared strings
                  Cell := TCell.Create(Row, Col, GetSharedString(StrToIntDef(CellValue, 0)), ctString);
                end
                else if CellType = 'str' then
                begin
                  // inline string
                  Cell := TCell.Create(Row, Col, CellValue, ctString);
                end
                else if CellType = 'inlineStr' then
                begin
                  // inline string (check <is><t> structure
                  var ISNode := CellNode.FindNode('is');
                  if ISNode <> nil then
                  begin
                    var TNode := ISNode.FindNode('t');
                    if TNode <> nil then
                      Cell := TCell.Create(Row, Col, TNode.Value, ctString)
                    else
                      Cell := TCell.Create(Row, Col, CellValue, ctString);
                  end
                  else
                    Cell := TCell.Create(Row, Col, CellValue, ctString);
                end
                else if CellType = 'b' then
                begin
                  // boolean
                  Cell := TCell.Create(Row, Col, CellValue = '1', ctBoolean);
                end
                else if CellType = 'e' then
                begin
                  // error
                  Cell := TCell.Create(Row, Col, CellValue, ctError);
                end
                else if CellValue <> '' then
                begin
                  // number or date
                  if TryStrToFloat(CellValue, NumValue, FormatSettings) then
                    Cell := TCell.Create(Row, Col, NumValue, ctNumber)
                  else
                    Cell := TCell.Create(Row, Col, CellValue, ctString);
                end
                else
                begin
                  // empty cell
                  Cell := TCell.Empty(Row, Col);
                end;

                AWorksheet.AddCell(Cell);
              end;
            finally
              CellNodes.Free;
            end;
          end;
        finally
          RowNodes.Free;
        end;
      finally
        Root.Free;
      end;
    except
      on E: Exception do
        raise EXlsx4DException.CreateFmt(
          'Failed to load worksheet data for sheet "%s": %s',
          [ASheetName, E.Message]
        );
    end;
  finally
    Parser.Free;
  end;
end;

function TXLSXEngine.ParseCellReference(const ACellRef: string; 
  out ARow, ACol: Integer): Boolean;
var
  I: Integer;
  ColPart, RowPart: string;
begin
  Result := False;
  ColPart := '';
  RowPart := '';
  
  for I := 1 to Length(ACellRef) do
  begin
    if CharInSet(ACellRef[I], ['A'..'Z', 'a'..'z']) then
      ColPart := ColPart + UpCase(ACellRef[I])
    else if CharInSet(ACellRef[I], ['0'..'9']) then
      RowPart := RowPart + ACellRef[I];
  end;
  
  if (ColPart = '') or (RowPart = '') then
    Exit;
    
  ACol := GetColumnIndex(ColPart);
  ARow := StrToIntDef(RowPart, 0);
  
  Result := (ARow > 0) and (ACol > 0);
end;

function TXLSXEngine.GetColumnIndex(const ACell: string): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 1 to Length(ACell) do
  begin
    Result := Result * 26 + (Ord(ACell[I]) - Ord('A') + 1);
  end;
end;

function TXLSXEngine.GetSharedString(AIndex: Integer): string;
begin
  if (AIndex >= 0) and (AIndex < FSharedStrings.Count) then
    Result := FSharedStrings[AIndex]
  else
    Result := '';
end;

function TXLSXEngine.LoadFromFile(const AFileName: string): TWorksheets;
begin
  if not TFile.Exists(AFileName) then
    raise EXlsx4DException.CreateFmt(
      'Unable to load file "%s". File does not exist or is not accessible.',
      [AFileName]
    );
    
  FWorksheets := TWorksheets.Create(True);
  
  try
    ExtractXLSXContents(AFileName);
    LoadSharedStrings;
    LoadWorksheets;

    Result := FWorksheets;
    FWorksheets := nil;
  except
    on E: Exception do
    begin
      Cleanup;
      raise EXlsx4DException.CreateFmt(
        'Failed to load XLSX file "%s": %s',
        [AFileName, E.Message]
      );
    end;
  end;
end;

end.
