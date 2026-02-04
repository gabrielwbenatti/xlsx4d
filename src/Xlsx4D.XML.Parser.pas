unit Xlsx4D.XML.Parser;

interface

uses
  System.Generics.Collections;

type
  TXMLAttribute = record
    Name: string;
    Value: string;
  end;

  TXMLNode = class
  private
    FName: string;
    FValue: string;
    FAttributes: TList<TXMLAttribute>;
    FChildren: TObjectList<TXMLNode>;
    FParent: TXMLNode;
    function GetAttribute(const AName: string): string;
    function GetChildCount: Integer;
    function GetChild(Index: Integer): TXMLNode;
  public
    constructor Create;
    destructor Destroy; override;

    function HasAttribute(const AName: string): Boolean;
    function FindNode(const AName: string): TXMLNode;
    procedure FindNodes(const AName: string; AList: TObjectList<TXMLNode>);

    property Name: string read FName write FName;
    property Value: string read FValue write FValue;
    property Attributes: TList<TXMLAttribute> read FAttributes write FAttributes;
    property Children: TObjectList<TXMLNode> read FChildren write FChildren;
    property Parent: TXMLNode read FParent write FParent;
    property Attribute[const AName: string]: string read GetAttribute;
    property ChildCount: Integer read GetChildCount;
    property Child[Index: Integer]: TXMLNode read GetChild;
  end;

  TXMLParser = class
  private
    FContent: string;
    FPosition: Integer;
    FLength: Integer;

    function CurrentChar: Char;
    function PeekChar(AOffset: Integer = 1): Char;
    procedure SkipWhitespace;
    procedure Advance(ACount: Integer = 1);
    function ReadUntil(const ADelimiter: string): string;
    function ReadQuotedString: string;
    function DecodeXMLEntities(const AValue: string): string;

    function ParseNode: TXMLNode;
    procedure ParseAttributes(ANode: TXMLNode);
    function IsEndOfContent: Boolean;
  public
    constructor Create;

    function Parse(const AXMLContent: string): TXMLNode;
    function ParseFragment(const AXMLContent: string): TObjectList<TXMLNode>;
  end;

  TXMLHelper = class
    class function GetNodeValue(ANode: TXMLNode; const APath: string): string;
    class function GetNodeAttribute(ANode: TXMLNode; const APath, AAttribute: string): string;
    class function FindNodesByName(ANode: TXMLNode; const AName: string): TObjectList<TXMLNode>;
    class function ExtractValue(const AXMLContent, ATagName: string): string;
    class function ExtractAttribute(const AXMLContent, ATagName, AAttrName: string): string;
  end;

implementation

uses
  System.SysUtils, System.StrUtils;

{ TXMLNode }

constructor TXMLNode.Create;
begin
  inherited Create;
  FAttributes := TList<TXMLAttribute>.Create;
  FChildren := TObjectList<TXMLNode>.Create(True);
  FParent := nil;
end;

destructor TXMLNode.Destroy;
begin
  FAttributes.Free;
  FChildren.Free;
  inherited;
end;

function TXMLNode.GetAttribute(const AName: string): string;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to FAttributes.Count - 1 do
  begin
    if SameText(FAttributes[I].Name, AName) then
    begin
      Result := FAttributes[I].Value;
      Exit;
    end;
  end;
end;

function TXMLNode.GetChild(Index: Integer): TXMLNode;
begin
  if (Index >= 0) and (Index < FChildren.Count) then
    Result := FChildren[Index]
  else
    Result := nil;
end;

function TXMLNode.GetChildCount: Integer;
begin
  Result := FChildren.Count;
end;

function TXMLNode.HasAttribute(const AName: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to FAttributes.Count - 1 do
  begin
    if SameText(FAttributes[I].Name, AName) then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

function TXMLNode.FindNode(const AName: string): TXMLNode;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to FChildren.Count - 1 do
  begin
    if SameText(FChildren[I].Name, AName) then
    begin
      Result := FChildren[I];
      Exit;
    end;
  end;
end;

procedure TXMLNode.FindNodes(const AName: string; AList: TObjectList<TXMLNode>);
var
  I: Integer;
begin
  for I := 0 to FChildren.Count - 1 do
  begin
    if SameText(FChildren[I].Name, AName) then
      AList.Add(FChildren[I]);
  end;
end;

{ TXMLParser }

constructor TXMLParser.Create;
begin
  inherited Create;
end;

function TXMLParser.CurrentChar: Char;
begin
  if FPosition <= FLength then
    Result := FContent[FPosition]
  else
    result := #0;
end;

function TXMLParser.PeekChar(AOffset: Integer): Char;
var
  Pos: Integer;
begin
  Pos := FPosition + AOffset;
  if (Pos > 0) and (Pos <= FLength) then
    Result := FContent[Pos]
  else
    Result := #0;
end;

procedure TXMLParser.SkipWhitespace;
begin
  while (FPosition <= FLength) and (CharInSet(FContent[FPosition], [' ', #9, #10, #13])) do
    Inc(FPosition);
end;

procedure TXMLParser.Advance(ACount: Integer);
begin
  Inc(FPosition, ACount);
end;

function TXMLParser.ReadUntil(const ADelimiter: string): string;
var
  StartPos, DelimPos: Integer;
begin
  StartPos := FPosition;
  DelimPos := PosEx(ADelimiter, FContent, FPosition);

  if DelimPos > 0 then
  begin
    Result := Copy(FContent, StartPos, DelimPos - StartPos);
    FPosition := DelimPos;
  end
  else
  begin
    Result := Copy(FContent, StartPos, FLength - StartPos + 1);
    FPosition := FLength + 1;
  end;
end;

function TXMLParser.ReadQuotedString: string;
var
  QuoteChar: Char;
  StartPos: Integer;
begin
  Result := '';
  QuoteChar := CurrentChar;

  if not CharInSet(QuoteChar, ['''', '"']) then
    Exit;

  Advance(1);
  StartPos := FPosition;

  while (FPosition <= FLength) and (CurrentChar <> QuoteChar) do
    Advance(1);

  Result := Copy(FContent, StartPos, FPosition - StartPos);

  if CurrentChar = QuoteChar then
    Advance(1);

  Result := DecodeXMLEntities(Result);
end;

function TXMLParser.DecodeXMLEntities(const AValue: string): string;
var
  Pos, EndPos, CharCode: Integer;
  EntityStr: string;
begin
  Result := AValue;
  Result := StringReplace(Result, '&lt;', '<', [rfReplaceAll]);
  Result := StringReplace(Result, '&gt;', '>', [rfReplaceAll]);
  Result := StringReplace(Result, '&quot;', '"', [rfReplaceAll]);
  Result := StringReplace(Result, '&apos;', '''', [rfReplaceAll]);
  Result := StringReplace(Result, '&amp;', '&', [rfReplaceAll]);

  Pos := 1;
  while Pos <= Length(Result) do
  begin
    if (Result[Pos] = '&') and (Pos < Length(Result)) and (Result[Pos + 1] = '#') then
    begin
      EndPos := PosEx(';', Result, Pos);
      if EndPos > Pos then
      begin
        EntityStr := Copy(Result, Pos + 2, EndPos - Pos - 2);

        if (Length(EntityStr) > 0) and (EntityStr[1] = 'x') then
        begin
          // hexadecimal
          if TryStrToInt('$' + Copy(EntityStr, 2, MaxInt), CharCode) then
          begin
            Delete(Result, Pos, EndPos - Pos + 1);
            Insert(Char(CharCode), Result, Pos);
          end;
        end
        else
        begin
          // decimal
          if TryStrToInt(EntityStr, CharCode) then
          begin
            Delete(Result, Pos, EndPos - Pos + 1);
            Insert(Char(CharCode), Result, Pos);
          end;
        end;
      end;
      Inc(Pos);
    end;
  end;
end;

function TXMLParser.IsEndOfContent: Boolean;
begin
  Result := FPosition > FLength;
end;

procedure TXMLParser.ParseAttributes(ANode: TXMLNode);
var
  AttrName, AttrValue: string;
  Attr: TXMLAttribute;
begin
  SkipWhitespace;

  while (not IsEndOfContent) and (CurrentChar <> '>') and (CurrentChar <> '/') do
  begin
    AttrName := '';
    while (not IsEndOfContent) and
          (not CharInSet(CurrentChar, [' ', '=', '>', '/', #9, #10, #13])) do
    begin
      AttrName := AttrName + CurrentChar;
      Advance(1);
    end;

    SkipWhitespace;

    if CurrentChar = '=' then
    begin
      Advance(1);
      SkipWhitespace;

      if CharInSet(CurrentChar, ['''', '"']) then
      begin
        AttrValue := ReadQuotedString;

        Attr.Name := AttrName;
        Attr.Value := AttrValue;
        ANode.Attributes.Add(Attr);
      end;
    end;

    SkipWhitespace;
  end;
end;

function TXMLParser.ParseNode: TXMLNode;
var
  Node, ChildNode: TXMLNode;
  NodeName, EndTagName, TextContent: string;
  IsSelfClosing, IsTextNode: Boolean;
begin
  Result := nil;

  // skip comments
  if (CurrentChar = '<') and (PeekChar(1) = '!') and (PeekChar(2) = '-') and (PeekChar(3) = '-') then
  begin
    ReadUntil('-->');
    Advance(3);
    Exit;
  end;

  // skip declaration and instructions
  if (CurrentChar = '<') and (PeekChar(1) = '?') then
  begin
    ReadUntil('?>');
    Advance(2);
    Exit;
  end;

  // skip doctype
  if (CurrentChar = '<') and (PeekChar(1) = '!') then
  begin
    ReadUntil('>');
    Advance(1);
    Exit;
  end;

  if CurrentChar <> '<' then
    Exit;

  Advance(1); // skip '<'

  // check closing tag
  if CurrentChar = '/' then
    Exit;

  Node := TXMLNode.Create;

  NodeName := '';
  while (not IsEndOfContent) and
        (not CharInSet(CurrentChar, [' ', '>', '/', #9, #10, #13])) do
  begin
    NodeName := NodeName + CurrentChar;
    Advance(1);
  end;

  Node.Name := NodeName;

  // parse attributes
  ParseAttributes(Node);

  // check self-closing
  IsSelfClosing := False;
  if CurrentChar = '/' then
  begin
    IsSelfClosing := True;
    Advance(1);
  end;

  if CurrentChar = '>' then
    Advance(1);

  if IsSelfClosing then
  begin
    Result := Node;
    Exit;
  end;

  // parse content and children
  while not IsEndOfContent do
  begin
    SkipWhitespace;

    if IsEndOfContent then
      Break;

    // check closing tag
    if (CurrentChar <> '<') and (PeekChar(1) = '/') then
    begin
      Advance(2); // skip '</'
      EndTagName := '';

      while (not IsEndOfContent) and (CurrentChar <>  '>') do
      begin
        EndTagName := EndTagName + CurrentChar;
        Advance(1);
      end;

      if CurrentChar = '>' then
        Advance(1);

      Break;
    end;

    // check CDATA
    if (CurrentChar = '<') and (PeekChar(1) = '!') and
       (PeekChar(2) = '[') and (Copy(FContent, FPosition, 9) = '<![CDATA[') then
    begin
      Advance(9);
      TextContent := ReadUntil(']]>');
      Node.Value := Node.Value + TextContent;
      Advance(3);
      Continue;
    end;

    // parse child
    if CurrentChar = '<' then
    begin
      ChildNode := ParseNode;
      if ChildNode <> nil then
      begin
        ChildNode.Parent := Node;
        Node.Children.Add(ChildNode);
      end
      else
      begin
        // read content
        TextContent := '';
        while (not IsEndOfContent) and (CurrentChar <> '<') do
        begin
          TextContent := TextContent + CurrentChar;
          Advance(1);
        end;

        if TextContent <> '' then
          Node.Value := Node.Value + DecodeXMLEntities(Trim(TextContent));
      end;
    end;
  end;

  Result := Node;
end;

function TXMLParser.Parse(const AXMLContent: string): TXMLNode;
var
  Node: TXMLNode;
begin
  FContent := AXMLContent;
  FPosition := 1;
  FLength := Length(FContent);

  Result := TXMLNode.Create;
  Result.Name := 'root';

  while not IsEndOfContent do
  begin
    SkipWhitespace;

    if IsEndOfContent then
      Break;

    Node := ParseNode;
    if Node <> nil then
    begin
      Node.Parent := Result;
      Result.Children.Add(Node);
    end;
  end;
end;

function TXMLParser.ParseFragment(
  const AXMLContent: string): TObjectList<TXMLNode>;
var
  Node: TXMLNode;
begin
  FContent := AXMLContent;
  FPosition := 1;
  FLength := Length(FContent);

  Result := TObjectList<TXMLNode>.Create(True);

  while not IsEndOfContent do
  begin
    SkipWhitespace;

    if IsEndOfContent then
      Break;

    Node := ParseNode;
    if Node <> nil then
      Result.Add(Node);
  end;
end;

{ TXMLHelper }

class function TXMLHelper.GetNodeValue(ANode: TXMLNode;
  const APath: string): string;
var
  PathParts: TArray<string>;
  CurrentNode: TXMLNode;
  I: Integer;
begin
  Result := '';

  if ANode = nil then
    Exit;

  if APath = '' then
  begin
    Result := ANode.Value;
    Exit;
  end;

  PathParts := APath.Split(['/']);
  CurrentNode := ANode;

  for I := 0 to High(PathParts) do
  begin
    if PathParts[I] = '' then
      Continue;

    CurrentNode := CurrentNode.FindNode(PathParts[I]);
    if CurrentNode = nil then
      Exit;
  end;

  Result := CurrentNode.Value;
end;

class function TXMLHelper.GetNodeAttribute(ANode: TXMLNode; const APath,
  AAttribute: string): string;
var
  PathParts: TArray<string>;
  CurrentNode: TXMLNode;
  I: Integer;
begin
  Result := '';

  if ANode = nil then
    Exit;

  CurrentNode := ANode;

  if APath <> '' then
  begin
    PathParts := APath.Split(['/']);

    for I := 0 to High(PathParts) do
    begin
      if PathParts[I] = '' then
        Continue;

      CurrentNode := CurrentNode.FindNode(PathParts[I]);
      if CurrentNode = nil then
        Exit;
    end;
  end;

  Result := CurrentNode.Attribute[AAttribute];
end;

class function TXMLHelper.FindNodesByName(ANode: TXMLNode;
  const AName: string): TObjectList<TXMLNode>;

  procedure FindRecursive(ACurrentNode: TXMLNode; const ASearchName: string; AResult: TObjectList<TXMLNode>);
  var
    I: Integer;
  begin
    if ACurrentNode = nil then
      Exit;

    if SameText(ACurrentNode.Name, ASearchName) then
      AResult.Add(ACurrentNode);

    for I := 0 to ACurrentNode.ChildCount - 1 do
      FindRecursive(ACurrentNode.Child[I], ASearchName, AResult);
  end;

begin
  Result := TObjectList<TXMLNode>.Create(False);
  FindRecursive(ANode, AName, Result);
end;

class function TXMLHelper.ExtractValue(const AXMLContent,
  ATagName: string): string;
var
  Parser: TXMLParser;
  Root: TXMLNode;
  Nodes: TObjectList<TXMLNode>;
begin
  Result := '';

  Parser := TXMLParser.Create;
  try
    Root := Parser.Parse(AXMLContent);
    try
      Nodes := FindNodesByName(Root, ATagName);
      try
        if Nodes.Count > 0 then
          Result := Nodes[0].Value;
      finally
        Nodes.Free;
      end;
    finally
      Root.Free;
    end;
  finally
    Parser.Free;
  end;
end;

class function TXMLHelper.ExtractAttribute(const AXMLContent, ATagName,
  AAttrName: string): string;
var
  Parser: TXMLParser;
  Root: TXMLNode;
  Nodes: TObjectList<TXMLNode>;
begin
  Result := '';

  Parser := TXMLParser.Create;
  try
    Root := Parser.Parse(AXMLContent);
    try
      Nodes := FindNodesByName(Root, ATagName);
      try
        if Nodes.Count > 0 then
          Result := Nodes[0].Attribute[AAttrName];
      finally
        Nodes.Free;
      end;
    finally
      Root.Free;
    end;
  finally
    Parser.Free;
  end;
end;

end.

