# XLSX4D

![Delphi](https://img.shields.io/badge/Delphi-10.3+-red.svg)
![Platforms](https://img.shields.io/badge/platform-Win32%20%7C%20Win64-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

> Biblioteca Delphi para leitura de arquivos Excel (XLSX) sem dependências do Microsoft Excel.

## 📋 Sobre

**XLSX4D** é uma biblioteca nativa Delphi para ler arquivos Excel **XLSX** de forma simples e eficiente, sem precisar ter o Microsoft Excel instalado.

### ⚠️ Importante

- ✅ **Formato XLSX** - Totalmente suportado (Excel 2007+)
- ❌ **Formato XLS** - Não suportado (formato legado do Excel 97-2003)

## 📦 Requisitos

- Delphi 10.3 Rio ou superior
- Windows (32-bit ou 64-bit)

> **Nota:** A biblioteca usa apenas componentes nativos do Delphi (RTL e XML RTL), sem dependências externas.

## 🚀 Instalação

### Via Boss

```bash
boss install github.com/gabrielwbenatti/xlsx4d
```

## 💡 Uso Rápido

```pascal
uses
  Xlsx4D, Xlsx4D.Types;

var
  Excel: TXlsx4D;
  Sheet: TWorksheet;
  Cell: TCell;
begin
  Excel := TXlsx4D.Create;
  try
    // Carregar arquivo XLSX
    if Excel.LoadFromFile('planilha.xlsx') then
    begin
      // Acessar primeira planilha
      Sheet := Excel.GetWorksheetByIndex(0);
      
      // Ler célula A1 (linha 1, coluna 1)
      Cell := Sheet[1, 1];
      ShowMessage(Cell.AsString);
      
      // Ler célula B2 (linha 2, coluna 2)
      Cell := Sheet[2, 2];
      ShowMessage('Valor: ' + Cell.AsString);
    end;
  finally
    Excel.Free;
  end;
end;
```