program SimpleReader;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  System.TypInfo,
  Xlsx4D in '..\..\src\Xlsx4D.pas',
  Xlsx4D.Types in '..\..\src\Xlsx4D.Types.pas',
  Xlsx4D.Engine.XLSX in '..\..\src\Xlsx4D.Engine.XLSX.pas',
  Xlsx4D.XML.Parser in '..\..\src\Xlsx4D.XML.Parser.pas';

procedure Op0_Exit;
begin
  WriteLn('=== Saindo do Programa ===');
  WriteLn;
  WriteLn('Pressione ENTER para sair...');
  ReadLn;
end;

procedure Op1_ListSheets;
var
  Xlsx: TXlsx4D;
  I: Integer;
  Sheet: TWorksheet;
begin
  WriteLn('=== Exemplo 1: Listar Planilhas ===');
  WriteLn;

  Xlsx := TXlsx4D.Create;
  try
    WriteLn('Carregando arquivo example.xlsx...');
    if not Xlsx.LoadFromFile('example.xlsx') then
    begin
      WriteLn('ERRO: Não foi possível carregar o arquivo');
      Exit;
    end;

    WriteLn('Arquivo carregado: ', Xlsx.FileName);
    WriteLn('Número de planilhas: ', Xlsx.GetWorksheetCount);
    WriteLn;

    if Xlsx.GetWorksheetCount > 0 then
    begin
      WriteLn('Lista de planilhas:');
      WriteLn('-------------------');
      for I := 0 to Xlsx.GetWorksheetCount - 1 do
      begin
        Sheet := Xlsx.GetWorksheetByIndex(I);
        if Assigned(Sheet) then
        begin
          WriteLn(Format('  %d. "%s" - %d linhas x %d colunas', 
            [I + 1, Sheet.Name, Sheet.RowCount, Sheet.ColCount]));
        end;
      end;
    end
    else
      WriteLn('Nenhuma planilha encontrada no arquivo.');

  finally
    Xlsx.Free;
  end;
end;

procedure Op2_FindSheetByName;
var
  Xlsx: TXlsx4D;
  Sheet: TWorksheet;
  SheetName: string;
begin
  WriteLn('=== Exemplo 2: Buscar Planilha por Nome ===');
  WriteLn;

  Write('Digite o nome da planilha (exemplo: planilha 1): ');
  ReadLn(SheetName);
  WriteLn;

  Xlsx := TXlsx4D.Create;
  try
    WriteLn('Carregando arquivo...');
    Xlsx.LoadFromFile('example.xlsx');

    WriteLn('Buscando planilha "', SheetName, '"...');
    Sheet := Xlsx.GetWorksheet(SheetName);

    if Assigned(Sheet) then
    begin
      WriteLn('Planilha encontrada!');
      WriteLn('Nome: ', Sheet.Name);
      WriteLn('Linhas: ', Sheet.RowCount);
      WriteLn('Colunas: ', Sheet.ColCount);
    end
    else
      WriteLn('Planilha "', SheetName, '" não encontrada.');

  finally
    Xlsx.Free;
  end;
end;

procedure Op3_ReadSpecificCell;
var
  Xlsx: TXlsx4D;
  Sheet: TWorksheet;
  Cell: TCell;
  Row, Col: Integer;
begin
  WriteLn('=== Exemplo 3: Ler Célula Específica ===');
  WriteLn;

  Write('Digite a linha (exemplo: 1): ');
  ReadLn(Row);
  Write('Digite a coluna (exemplo: 1 para A, 2 para B): ');
  ReadLn(Col);
  WriteLn;

  Xlsx := TXlsx4D.Create;
  try
    WriteLn('Carregando arquivo...');
    Xlsx.LoadFromFile('example.xlsx');

    Sheet := Xlsx.GetWorksheetByIndex(0);
    if not Assigned(Sheet) then
    begin
      WriteLn('Nenhuma planilha encontrada');
      Exit;
    end;

    WriteLn(Format('Lendo célula [%d,%d] da planilha "%s"...', [Row, Col, Sheet.Name]));
    Cell := Sheet[Row, Col];

    if Cell.IsEmpty then
      WriteLn('A célula está vazia')
    else
    begin
      WriteLn('Tipo: ', GetEnumName(TypeInfo(TCellType), Ord(Cell.CellType)));
      WriteLn('Valor: ', Cell.AsString);
      
      case Cell.CellType of
        ctNumber: WriteLn('Valor numérico: ', Cell.AsFloat:0:2);
        ctBoolean: WriteLn('Valor booleano: ', Cell.AsBoolean);
      end;
    end;

  finally
    Xlsx.Free;
  end;
end;

procedure Op4_QuickRead;
var
  Xlsx: TXlsx4D;
  Sheet: TWorksheet;
  Row, Col: Integer;
  Cell: TCell;
  MaxRows, MaxCols: Integer;
begin
  WriteLn('=== Exemplo 4: Leitura Rápida ===');
  WriteLn;

  Xlsx := TXlsx4D.Create;
  try
    WriteLn('Carregando arquivo...');
    Xlsx.LoadFromFile('example.xlsx');

    Sheet := Xlsx.GetWorksheetByIndex(0);
    if not Assigned(Sheet) then
    begin
      WriteLn('Nenhuma planilha encontrada');
      Exit;
    end;

    WriteLn('Planilha: ', Sheet.Name);
    WriteLn('Exibindo até 10 linhas e 10 colunas');
    WriteLn;

    MaxRows := Sheet.RowCount;
    if MaxRows > 10 then MaxRows := 10;
    
    MaxCols := Sheet.ColCount;
    if MaxCols > 10 then MaxCols := 10;

    for Row := 1 to MaxRows do
    begin
      Write(Format('Linha %2d: ', [Row]));
      
      for Col := 1 to MaxCols do
      begin
        Cell := Sheet[Row, Col];
        if not Cell.IsEmpty then
          Write('[', Cell.AsString, '] ')
        else
          Write('[      ] ');
      end;
      
      WriteLn;
    end;

  finally
    Xlsx.Free;
  end;
end;

procedure Op5_IterateLines;
var
  Xlsx: TXlsx4D;
  Sheet: TWorksheet;
  Row, Col: Integer;
  Cell: TCell;
  RowCells: TRowCells;
begin
  WriteLn('=== Exemplo 5: Iterar por Linhas ===');
  WriteLn;

  Xlsx := TXlsx4D.Create;
  try
    WriteLn('Carregando arquivo...');
    Xlsx.LoadFromFile('example.xlsx');

    Sheet := Xlsx.GetWorksheetByIndex(0);
    if not Assigned(Sheet) then
    begin
      WriteLn('Nenhuma planilha encontrada');
      Exit;
    end;

    WriteLn('Planilha: ', Sheet.Name);
    WriteLn('Total de linhas: ', Sheet.RowCount);
    WriteLn;

    for Row := 1 to Sheet.RowCount do
    begin
      WriteLn('=== Linha ', Row, ' ===');
      
      RowCells := Sheet.GetRow(Row);
      if Assigned(RowCells) then
      begin
        WriteLn('Total de células nesta linha: ', RowCells.Count);
        
        for Col := 0 to RowCells.Count - 1 do
        begin
          Cell := RowCells[Col];
          WriteLn(Format('  Coluna %d: %s (Tipo: %s)', 
            [Cell.Col, Cell.AsString, GetEnumName(TypeInfo(TCellType), Ord(Cell.CellType))]));
        end;
      end
      else
        WriteLn('  Linha vazia');
      
      WriteLn;
    end;

  finally
    Xlsx.Free;
  end;
end;

procedure ShowMenu;
begin
  WriteLn;
  WriteLn('========================================');
  WriteLn('  XLSX4D - Biblioteca de Leitura      ');
  WriteLn('  de Arquivos Excel para Delphi        ');
  WriteLn('========================================');
  WriteLn;
  WriteLn('IMPORTANTE: coloque um arquivo "example.xlsx" no diretório do executável');
  WriteLn('para testar os exemplos');
  WriteLn;
  WriteLn('Escolha um exemplo: ');
  WriteLn('1 - Listar todas as planilhas');
  WriteLn('2 - Buscar planilha por nome');
  WriteLn('3 - Ler célula específica');
  WriteLn('4 - Leitura rápida');
  WriteLn('5 - Iterar por linhas');
  WriteLn('0 - Sair');
  WriteLn;
end;

var
  Option: Integer;
  Continue: Boolean;
begin
  try
    Continue := True;
    
    while Continue do
    begin
      ShowMenu;
      
      Write('Opção: ');
      ReadLn(Option);
      WriteLn;

      case Option of
        1: Op1_ListSheets;
        2: Op2_FindSheetByName;
        3: Op3_ReadSpecificCell;
        4: Op4_QuickRead;
        5: Op5_IterateLines;
        0: begin
             Op0_Exit;
             Continue := False;
           end;
      else
        WriteLn('Opção não encontrada');
      end;
      
      if Continue then
      begin
        WriteLn;
        WriteLn('Pressione ENTER para voltar ao menu...');
        ReadLn;
      end;
    end;
    
  except
    on E: Exception do
    begin
      WriteLn('ERRO: ', E.Message);
      WriteLn;
      WriteLn('Pressione ENTER para sair...');
      ReadLn;
    end;
  end;
end.
