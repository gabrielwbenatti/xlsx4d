unit Xlsx4D.XML.Parser;

interface

uses
  System.Generics.Collections;

type
  TXMLAttribute = record
    Name: string;
    Value: Variant;
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

    function CurrentChat: Char;
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

{ TXMLNode }

constructor TXMLNode.Create;
begin

end;

destructor TXMLNode.Destroy;
begin

  inherited;
end;

function TXMLNode.FindNode(const AName: string): TXMLNode;
begin

end;

procedure TXMLNode.FindNodes(const AName: string; AList: TObjectList<TXMLNode>);
begin

end;

function TXMLNode.GetAttribute(const AName: string): string;
begin

end;

function TXMLNode.GetChild(Index: Integer): TXMLNode;
begin

end;

function TXMLNode.GetChildCount: Integer;
begin

end;

function TXMLNode.HasAttribute(const AName: string): Boolean;
begin

end;

{ TXMLParser }

procedure TXMLParser.Advance(ACount: Integer);
begin

end;

constructor TXMLParser.Create;
begin

end;

function TXMLParser.CurrentChat: Char;
begin

end;

function TXMLParser.DecodeXMLEntities(const AValue: string): string;
begin

end;

function TXMLParser.IsEndOfContent: Boolean;
begin

end;

function TXMLParser.Parse(const AXMLContent: string): TXMLNode;
begin

end;

procedure TXMLParser.ParseAttributes(ANode: TXMLNode);
begin

end;

function TXMLParser.ParseFragment(
  const AXMLContent: string): TObjectList<TXMLNode>;
begin

end;

function TXMLParser.ParseNode: TXMLNode;
begin

end;

function TXMLParser.PeekChar(AOffset: Integer): Char;
begin

end;

function TXMLParser.ReadQuotedString: string;
begin

end;

function TXMLParser.ReadUntil(const ADelimiter: string): string;
begin

end;

procedure TXMLParser.SkipWhitespace;
begin

end;

{ TXMLHelper }

class function TXMLHelper.ExtractAttribute(const AXMLContent, ATagName,
  AAttrName: string): string;
begin

end;

class function TXMLHelper.ExtractValue(const AXMLContent,
  ATagName: string): string;
begin

end;

class function TXMLHelper.FindNodesByName(ANode: TXMLNode;
  const AName: string): TObjectList<TXMLNode>;
begin

end;

class function TXMLHelper.GetNodeAttribute(ANode: TXMLNode; const APath,
  AAttribute: string): string;
begin

end;

class function TXMLHelper.GetNodeValue(ANode: TXMLNode;
  const APath: string): string;
begin

end;

end.

