unit XLSX4D.Types;

interface

type
  TXLSX4DCellDataType = (
    ctString,
    ctNumber,
    ctBoolean,
    ctDate,
    ctFormula,
    ctError
  );

  TXLSX4DCompressionLevel = (
    clNone,
    clFastest,
    clDefault,
    clBest
  );

  TXLSX4DErrorType = (
    etNone,
    etNull,
    etDiv0,
    etValue,
    etRef,
    etName,
    etNum,
    etNA
  );

  TXLSX4DReadOptions = set of (
    roFormulas,
    roFormats,
    roComments,
    roMergedCells,
    roHiddenSheets
  );

  TXLSX4DWriteOptions = set of (
    woFormulas,
    woFormats,
    woComments,
    woMergedCells,
    woAutoFilter
  );

  TXLSX4DCellInfo = record
    Row: Integer;
    Column: Integer;
    Value: Variant;
    Formula: string;
    Format: string;
    DataType: TXLSX4DCellDataType;
  end;

  TXLSX4DRangeInfo = record
    StartRow: Integer;
    StartColumn: Integer;
    EndRow: Integer;
    EndColumn: Integer;
  end;

implementation

end.

