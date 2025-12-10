unit Xlsx4D;

interface

uses
  Xlsx4D.Types, System.SysUtils;

type
  TXlsx4D = class
  private
    FWorkSheets: TWorksheets;
    FFileName: string;
  public
    /// <summary>
    ///   Creates a new instance of TXlsx4D
    /// </summary>
    constructor Create;
    destructor Destroy; override;

    /// <summary>
    ///   Loads an XLSX file from disk
    /// </summary>
    /// <param name="AFileName">Full path to the Excel file (XLSX or XLS format)</param>
    /// <returns>True if file was loaded successfully, False otherwise</returns>
    /// <exception cref="EXlsx4DException">Raised when file does not exist</exception>
    function LoadFromFile(const AFileName: string): Boolean;

    /// <summary>
    ///   Gets the collection of worksheets
    /// </summary>
    /// <returns>The collection of worksheets</returns>    
    function GetWorksheets: TWorksheets;

    /// <summary>
    /// Gets a worksheet by its name
    /// </summary>
    /// <param name="AName">The name of the worksheet to find</param>
    /// <returns>The worksheet if found, nil otherwise</returns>
    function GetWorksheet(const AName: string): TWorksheet;

    /// <summary>
    ///   Gets a worksheet by its index
    /// </summary>
    /// <param name="AIndex">The index of the worksheet to get (0-based)</param>
    /// <returns>The worksheet if found, nil otherwise</returns>
    function GetWorksheetByIndex(const AIndex: Integer): TWorksheet;

    /// <summary>
    ///   Gets the number of worksheets in the workbook
    /// </summary>    
    function GetWorksheetCount: Integer;

    property FileName: string read FFileName;
    property Worksheets: TWorksheets read GetWorksheets;
  end;

implementation

uses 
  Xlsx4D.Engine.XLSX;

{ TXlsx4D }

constructor TXlsx4D.Create;
begin
  inherited Create;
  FWorkSheets := nil;
  FFileName := '';
end;

destructor TXlsx4D.Destroy;
begin
  if FWorkSheets <> nil then
    FWorkSheets.Free;
  inherited;
end;

function TXlsx4D.GetWorksheet(const AName: string): TWorksheet;
begin
  if FWorkSheets = nil then 
    Result := Nil
  else
    Result := FWorkSheets.FindByName(AName);
end;

function TXlsx4D.GetWorksheetByIndex(const AIndex: Integer): TWorksheet;
begin
  if (FWorkSheets = nil) or (AIndex < 0) or (AIndex >= FWorkSheets.Count) then 
    Result := Nil
  else
    Result := FWorkSheets.Items[AIndex];
end;

function TXlsx4D.GetWorksheetCount: Integer;
begin
  if FWorkSheets = nil then
    Result := 0
  else
    Result := FWorkSheets.Count;
end;

function TXlsx4D.GetWorksheets: TWorksheets;
begin
  Result := FWorkSheets;
end;

function TXlsx4D.LoadFromFile(const AFileName: string): Boolean;
var
  Engine: TXLSXEngine;
begin
  Result := False;

  if not FileExists(AFileName) then
    raise EXlsx4DException.CreateFmt(
      'Unable to load file "%s". File does not exist or is not accessible.',
      [AFileName]
    );

  // free previous worksheets if any
  if FWorkSheets <> nil then
    FreeAndNil(FWorkSheets);

  Engine := TXLSXEngine.Create;
  try
    FWorkSheets := Engine.LoadFromFile(AFileName);
    FFileName := AFileName;

    Result := (Assigned(FWorkSheets));
  finally
    Engine.Free;    
  end;
end;

end.

