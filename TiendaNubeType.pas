type
  /// Records

  /// Deserialize records
  /// <summary>
  /// Language string
  /// </summary>
  TTiendaNubeLangString = record
    en: string;
    es: string;
    pt: string;
  end;

  /// <summary>
  /// Variant record
  /// </summary>
  TTiendaNubeVariant = record
    id: UInt64;
    image_id: UInt64;
    product_id: UInt64;
    stock_management: Boolean;
    stock: Integer;
    price: string;
    sku: string;
    mpn: string;
    barcode: string;
  end;

  /// <summary>
  /// Image record
  /// </summary>
  TTiendaNubeImage = record
    id: UInt64;
    product_id: UInt64;
    src: string;
    position: Integer;
  end;

  /// <summary>
  /// Category record
  /// </summary>
  TTiendaNubeCategory = record
    id: UInt64;
    name: TTiendaNubeLangString;
    description: TTiendaNubeLangString;
    handle: TTiendaNubeLangString;
    parent: UInt64;
  end;

  /// <summary>
  /// Product record
  /// </summary>
  TTiendaNubeProduct = record
    id: UInt64;
    name: TTiendaNubeLangString;
    description: TTiendaNubeLangString;
    handle: TTiendaNubeLangString;
    canonical_url: string;
    variants: TArray<TTiendaNubeVariant>;
    images: TArray<TTiendaNubeImage>;
    categories: TArray<TTiendaNubeCategory>;
  end;

  /// <summary>
  /// Array of products record
  /// </summary>
  TTiendaNubeProducts = TArray<TTiendaNubeProduct>;
  /// Deserialize records end

  TTiendaNubeResponse = record
    Status: Integer;
    Content: String;
    Headers: TStringList;
  end;

  TRequest = record
    time: TDateTime;
    URL: string;
    method: string;
    body: string;
    response: TTiendaNubeResponse;
  end;

  TRequestList = TList<TRequest>;

  /// classes
  TWebBrowserForm = class(TForm)
  private
    { Private declarations }
    FWebBrowser: TWebBrowser;
    FCode: string;
    procedure FOnDocumentComplete(ASender: TObject; const pDisp: IDispatch;
      const URL: OleVariant);
  public
    { Public declarations }
    constructor CreateNew(AOwner: TComponent; Dummy: Integer = 0); override;
    destructor Destroy; override;
    function ShowModalAppId(appId: string): string;
  end;

  TTiendaNubeRequest = class(TComponent)
  private
    FRequests: TRequestList;
    FErrors: TRequestList;
    procedure AddRequest(method, URL, body: string;
      response: TTiendaNubeResponse);
  protected
    procedure DoRequest(URL, token: string; var res: TTiendaNubeResponse;
      method: TRESTRequestMethod = rmGET; body: string = string.Empty);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Requests: TRequestList read FRequests;
    property Errors: TRequestList read FErrors;
  end;

  TTiendaNube = class(TTiendaNubeRequest)
  private
    FormWB: TWebBrowserForm;
    FToken: string;
    FStore: string;
    function ConvertToBase64(const filename: string): string;
    function ModProductInternal(productId, name: string;
      categories: TStringList): string;
    function StrIsInt(s: string): Boolean;
    function StrIsFloat(s: string): Boolean;
    function Auth(Code, ClientID, ClientSecret: string): Boolean;
  public
    // Constructor
    constructor Create(AOwner: TComponent; token, store: string);
      reintroduce; overload;
    destructor Destroy; override;
    // Auth
    function InstallApp(ClientID, ClientSecret: string): Boolean;

    // Products
    function GetProduct(productId: string = string.Empty): string;
    function GetProductSKU(sku: string): string;
    function AddProduct(name: string; categories: TStringList = nil;
      sku: string = string.Empty; price: string = string.Empty;
      stock: string = string.Empty): string;
    function RemoveProduct(productId: string): Boolean;
    function ModProduct(productId, name, sku, price, stock: string;
      categories: TStringList): Boolean;
    // Variant
    function ModVariant(productId, variantId, sku, price,
      stock: string): string;

    // Categories
    function GetCategories(): string;
    function AddCategory(category: string;
      parent: string = string.Empty): string;
    function RemoveCategory(category: string): Boolean;

    // Products images
    function AddImage(productId, Image: string): string;
    function RemoveImage(productId, ImageID: string): Boolean;

    // Deserialize
    function Deserialize<T>(s: string): T;
  published
    property token: string read FToken write FToken;
    property store: string read FStore write FStore;
    property Requests;
    property Errors;
  end;

  TBrowserEmulation = class
  private
    class function GetExeName(): String; inline;
  public const
    // Quelle: https://msdn.microsoft.com/library/ee330730.aspx, Stand: 2017-04-26
    IE11_default = 11000;
    IE11_Quirks = 11001;
    IE10_force = 10001;
    IE10_default = 10000;
    IE9_Quirks = 9999;
    IE9_default = 9000;
    /// <summary>
    /// Webpages containing standards-based !DOCTYPE directives are displayed in IE7
    /// Standards mode. Default value for applications hosting the WebBrowser Control.
    /// </summary>
    IE7_embedded = 7000;
  public
    class procedure SetBrowserEmulation(const Value: DWORD);
  end;