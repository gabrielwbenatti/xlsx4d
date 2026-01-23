unit Xlsx4D.Engine.XLSX;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Zip,
  System.Generics.Collections,
  Xlsx4D.Types;

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
    
    // Helpers para parsing XML
    function ParseXMLAttribute(const AXMLLine, AAttrName: string): string;
    function ParseXMLValue(const AXMLContent, ATagName: string): string;
    function ExtractAllBetweenTags(const AXMLContent, AStartTag, AEndTag: string): TStringList;
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

function TXLSXEngine.ParseXMLAttribute(const AXMLLine, AAttrName: string): string;
var
  StartPos, EndPos: Integer;
  SearchStr: string;
begin
  Result := '';
   
  SearchStr := AAttrName + '="';
  StartPos := Pos(SearchStr, AXMLLine);
  
  if StartPos > 0 then
  begin
    StartPos := StartPos + Length(SearchStr);
    EndPos := PosEx('"', AXMLLine, StartPos);
    if EndPos > StartPos then
      Result := Copy(AXMLLine, StartPos, EndPos - StartPos);
  end;
end;

function TXLSXEngine.ParseXMLValue(const AXMLContent, ATagName: string): string;
var
  StartTag, EndTag: string;
  StartPos, EndPos, TagEndPos: Integer;
begin
  Result := '';

  // Look for the opening tag (it may have attributes).
  StartTag := '<' + ATagName;
  EndTag := '</' + ATagName + '>';

  StartPos := Pos(StartTag, AXMLContent);
  if StartPos > 0 then
  begin
    // Find the end of the opening tag (look for '>')
    TagEndPos := PosEx('>', AXMLContent, StartPos);

    if TagEndPos > 0 then
    begin
      StartPos := TagEndPos + 1; // Position after the '>'
      EndPos := PosEx(EndTag, AXMLContent, StartPos);
      if EndPos > StartPos then
        Result := Copy(AXMLContent, StartPos, EndPos - StartPos);
    end;
  end;
end;

function TXLSXEngine.ExtractAllBetweenTags(const AXMLContent, AStartTag, AEndTag: string): TStringList;
var
  CurrentPos: Integer;
  StartPos, OpenEndPos, EndPos: Integer;
  TagContent: string;
begin
  Result := TStringList.Create;
  CurrentPos := 1;

  while True do
  begin
    StartPos := PosEx(AStartTag, AXMLContent, CurrentPos);
    if StartPos = 0 then
      Break;

    // Find the end of the opening tag (eg.: <c ...>)
    OpenEndPos := PosEx('>', AXMLContent, StartPos);
    if OpenEndPos = 0 then
      Break;

    // Check if it's self-closing (eg.: <c .../>)
    if (AXMLContent[OpenEndPos - 1] = '/') then
    begin
      // (ex.: <c .../>)
      TagContent := Copy(
        AXMLContent,
        StartPos,
        OpenEndPos - StartPos + 1
      );
      CurrentPos := OpenEndPos + 1;
    end
    else
    begin
      // (ex.: <c ...> ... </c>)
      EndPos := PosEx(AEndTag, AXMLContent, OpenEndPos);
      if EndPos = 0 then
        Break;

      TagContent := Copy(
        AXMLContent,
        StartPos,
        (EndPos + Length(AEndTag)) - StartPos
      );
      CurrentPos := EndPos + Length(AEndTag);
    end;

    Result.Add(TagContent);
  end;
end;

procedure TXLSXEngine.LoadSharedStrings;
var
  SharedStringsPath: string;
  XMLContent: string;
  SITags: TStringList;
  I: Integer;
  TValue: string;
begin
  FSharedStrings.Clear;
  
  SharedStringsPath := TPath.Combine(FTempPath, 'xl\sharedStrings.xml');
  
  if not TFile.Exists(SharedStringsPath) then
    Exit;
  
  try
    XMLContent := TFile.ReadAllText(SharedStringsPath, TEncoding.UTF8);
    
    SITags := ExtractAllBetweenTags(XMLContent, '<si>', '</si>');
    try
      for I := 0 to SITags.Count - 1 do
      begin
        TValue := ParseXMLValue(SITags[I], 't');
        FSharedStrings.Add(TValue);
      end;
    finally
      SITags.Free;
    end;
  except
    on E: Exception do
      raise EXlsx4DException.CreateFmt(
        'Failed to load shared strings: %s',
        [E.Message]
      );
  end;
end;

procedure TXLSXEngine.LoadWorksheets;
var
  WorkbookPath: string;
  XMLContent: string;
  SheetTags: TStringList;
  I: Integer;
  SheetName: string;
  Worksheet: TWorksheet;
begin
  WorkbookPath := TPath.Combine(FTempPath, 'xl\workbook.xml');
  
  if not TFile.Exists(WorkbookPath) then
    raise EXlsx4DException.Create('workbook.xml not found in XLSX package.');
  
  try
    XMLContent := TFile.ReadAllText(WorkbookPath, TEncoding.UTF8);
    
    SheetTags := TStringList.Create;
    try
      var CurrentPos := 1;
      while True do
      begin
        var StartPos := PosEx('<sheet ', XMLContent, CurrentPos);
        if StartPos = 0 then
          Break;
        
        var EndPos := PosEx('/>', XMLContent, StartPos);
        if EndPos = 0 then
          EndPos := PosEx('>', XMLContent, StartPos);
        
        if EndPos = 0 then
          Break;
        
        var SheetTag := Copy(XMLContent, StartPos, (EndPos + 2) - StartPos);
        SheetTags.Add(SheetTag);
        
        CurrentPos := EndPos + 2;
      end;
      
      for I := 0 to SheetTags.Count - 1 do
      begin
        SheetName := ParseXMLAttribute(SheetTags[I], 'name');
        
        if SheetName <> '' then
        begin
          Worksheet := TWorksheet.Create(SheetName);
          FWorksheets.Add(Worksheet);
          
          LoadWorksheetData(SheetName, Worksheet, I + 1);
        end;
      end;
    finally
      SheetTags.Free;
    end;
  except
    on E: Exception do
      raise EXlsx4DException.CreateFmt(
        'Failed to load worksheets: %s',
        [E.Message]
      );
  end;
end;

procedure TXLSXEngine.LoadWorksheetData(const ASheetName: string; 
  AWorksheet: TWorksheet; ASheetIndex: Integer);
var
  SheetDataPath: string;
  XMLContent: string;
  RowTags, CellTags: TStringList;
  I, J: Integer;
  CellRef, CellType, CellValue: string;
  Row, Col: Integer;
  Cell: TCell;
  NumValue: Double;
  CellTag: string;
  FormatSettings: TFormatSettings;
begin
  SheetDataPath := TPath.Combine(FTempPath, Format('xl\worksheets\sheet%d.xml', [ASheetIndex]));
  
  if not TFile.Exists(SheetDataPath) then
    Exit;

  FormatSettings := TFormatSettings.Create('en-US');
  FormatSettings.DecimalSeparator := '.';
  
  try
    XMLContent := TFile.ReadAllText(SheetDataPath, TEncoding.UTF8);
    
    RowTags := ExtractAllBetweenTags(XMLContent, '<row ', '</row>');
    try
      for I := 0 to RowTags.Count - 1 do
      begin
        CellTags := ExtractAllBetweenTags(RowTags[I], '<c ', '</c>');
        try
          for J := 0 to CellTags.Count - 1 do
          begin
            CellTag := CellTags[J];
            
            CellRef := ParseXMLAttribute(CellTag, 'r');
            CellType := ParseXMLAttribute(CellTag, 't');
            CellValue := ParseXMLValue(CellTag, 'v');
            
            if not ParseCellReference(CellRef, Row, Col) then
              Continue;
            
            if CellType = 's' then
            begin
              // String compartilhada
              Cell := TCell.Create(Row, Col, GetSharedString(StrToIntDef(CellValue, 0)), ctString);
            end
            else if CellType = 'str' then
            begin
              // String inline
              Cell := TCell.Create(Row, Col, CellValue, ctString);
            end
            else if CellType = 'b' then
            begin
              // Boolean
              Cell := TCell.Create(Row, Col, CellValue = '1', ctBoolean);
            end
            else if CellType = 'e' then
            begin
              // Error
              Cell := TCell.Create(Row, Col, CellValue, ctError);
            end
            else if CellValue <> '' then
            begin
              // Número ou data
              if TryStrToFloat(CellValue, NumValue, FormatSettings) then
                Cell := TCell.Create(Row, Col, NumValue, ctNumber)
              else
                Cell := TCell.Create(Row, Col, CellValue, ctString);
            end
            else
            begin
              // Célula vazia
              Cell := TCell.Empty(Row, Col);
            end;
            
            AWorksheet.AddCell(Cell);
          end;
        finally
          CellTags.Free;
        end;
      end;
    finally
      RowTags.Free;
    end;
  except
    on E: Exception do
      raise EXlsx4DException.CreateFmt(
        'Failed to load worksheet data for sheet "%s": %s',
        [ASheetName, E.Message]
      );
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
