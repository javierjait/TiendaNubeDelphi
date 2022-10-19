unit TiendaNube;

interface

uses
  System.SysUtils, System.Classes, Vcl.Controls, Vcl.Forms,
  SHDocVw, REST.Client, REST.Types, System.JSON, System.JSON.Serializers,
  System.NetEncoding,
  System.Generics.Collections, System.IOUtils, System.Win.Registry,
  System.Types, Windows, Vcl.Dialogs;

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
    promotional_price: string;
    sku: string;
    mpn: string;
    barcode: string;
    weight: string;
    width: string;
    height: string;
    depth: string;
    age_group: string;
    gender: string;
    cost: string;
    created_at: string;
    updated_at: string;
  end;

  /// <summary>
  /// Image record
  /// </summary>
  TTiendaNubeImage = record
    id: UInt64;
    product_id: UInt64;
    src: string;
    position: Integer;
    alt: TArray<string>;
    created_at: string;
    updated_at: string;
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
    subcategories: TArray<UInt64>;
    created_at: string;
    updated_at: string;
  end;

  /// <summary>
  /// Array of category
  /// </summary>
  TTiendaNubeCategories = TArray<TTiendaNubeCategory>;

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
    categories: TTiendaNubeCategories;
    brand: boolean;
    [JsonName('published')]
    prod_published: boolean;
    free_shipping: boolean;
    video_url: string;
    seo_title: TTiendaNubeLangString;
    seo_description: TTiendaNubeLangString;
    tags: string;
    requires_shipping: boolean;
    created_at: string;
    updated_at: string;
  end;

  /// <summary>
  /// Array of product record
  /// </summary>
  TTiendaNubeProducts = TArray<TTiendaNubeProduct>;

  /// <summary>
  /// DateTime record
  /// </summary>
  TTiendaNubeDateTime = record
    date: string;
    timezone_type: integer;
    timezone: string;
  end;

  TTiendaNubePaymentDetail = record
    method: string;
    credit_card_company: string;
    installments: Integer;
  end;

  /// <summary>
  /// Product order record
  /// </summary>
  TTiendaNubeOrderProduct = record
    id: UInt64;
    depth: string;
    height: string;
    name: string;
    price: string;
    product_id: UInt64;
    image: TTiendaNubeImage;
    quatity: Integer;
    free_shipping: boolean;
    weight: string;
    width: string;
    variant_id: UInt64;
    variant_values: TArray<string>;
    properties: TArray<string>;
    sku: string;
    barcode: string;
  end;

  /// <summary>
  /// Address record
  /// </summary>
  TTiendaNubeAddress = record
    address: string;
    city: string;
    default: Boolean;
    floor: string;
    id: UInt64;
    locality: string;
    name: string;
    number: string;
    phone: string;
    province: string;
    zipcode: string;
    customs: string;
  end;

  /// <summary>
  /// Client detail record
  /// </summary>
  TTiendaNubeClientDetail = record
    browser_ip: string;
    user_agent: string;
  end;

  /// <summary>
  /// Order record
  /// </summary>
  TTiendaNubeOrder = record
    id: UInt64;
    token: string;
    store_id: UInt64;
    shipping_min_days: string;
    shipping_max_days: string;
    billing_name: string;
    billing_phone: string;
    billing_address: string;
    billing_floor: string;
    billing_locality: string;
    billing_zipcode: string;
    billing_province: string;
    billing_country: string;
    shipping_cost_owner: string;
    shipping_cost_customer: string;
    subtotal: string;
    coupon: TArray<string>;
    discount: string;
    discount_coupon: string;
    discount_gateway: string;
    total: string;
    total_usd: string;
    checkout_enabled: Boolean;
    weight: string;
    currency: string;
    language: string;
    gateway: string;
    gateway_id: string;
    gateway_name: string;
    shipping: string;
    shipping_option: string;
    shipping_option_code: string;
    shipping_option_reference: string;
    shipping_pickup_details: string;
    shipping_tracking_number: string;
    shipping_tracking_url: string;
    shipping_store_branch_name: string;
    shipping_pickup_type: string;
    storefront: string;
    note: string;
    completed_at: TTiendaNubeDateTime;
    next_action: string;
    payment_details: TTiendaNubePaymentDetail;
    products: TArray<TTiendaNubeOrderProduct>;
    number: Int64;
    cancel_reason: string;
    owner_note: string;
    cancelled_at: string;
    closed_at: string;
    read_at: string;
    status: string;
    payment_status: string;
    gateway_link: string;
    shipping_carrier_name: string;
    shipping_address: TTiendaNubeAddress;
    shipping_status: string;
    shipped_at: string;
    paid_at: string;
    landing_url: string;
    client_details: TTiendaNubeClientDetail;
    app_id: string;
  end;

  /// <summary>
  /// Array of order record
  /// </summary>
  TTiendaNubeOrders = TArray<TTiendaNubeOrder>;
  /// Deserialize records end

  TTiendaNubeResponse = record
    status: Integer;
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
    function AddImage(productId, image: string): string;
    function RemoveImage(productId, ImageID: string): Boolean;

    // Orders
    function GetOrder(orderId: string = string.Empty): string;

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

implementation

const
  apiURL = 'https://api.tiendanube.com/v1/';
  authURL = 'https://www.tiendanube.com/apps/authorize/token';

  /// TTiendaNubeRequest
  // constructor/destroy
constructor TTiendaNubeRequest.Create(AOwner: TComponent);
begin
  inherited;
  FRequests := TRequestList.Create;
  FErrors := TRequestList.Create;
end;

destructor TTiendaNubeRequest.Destroy;
begin
  FErrors.Free;
  FRequests.Free;
  inherited;
end;
// Constructor/destroy end

// private
/// <summary> Add request to request/error list</summary>
/// <param name="method">method</param>
/// <param name="URL">URL</param>
/// <param name="body">Body</param>
/// <param name="response">response record</param>
procedure TTiendaNubeRequest.AddRequest(method, URL, body: string;
  response: TTiendaNubeResponse);
var
  req: TRequest;
begin
  req.time := Now;
  req.method := method;
  req.URL := URL;
  req.body := body;
  req.response := response;
  FRequests.Add(req);
  if not((req.response.status = 200) or (req.response.status = 201)) then
    FErrors.Add(req);
end;

// protected
/// <summary>Do TiendaNube request</summary>
/// <param name="URL">URL</param>
/// <param name="token">Auth token</param>
/// <param name="res">TiendaNube response record</param>
/// <param name="method">return response record</param>
/// <param name="body">Request body</param>
procedure TTiendaNubeRequest.DoRequest(URL, token: string;
  var res: TTiendaNubeResponse; method: TRESTRequestMethod = rmGET;
  body: string = string.Empty);
var
  methodS: string;
  Client: TRestClient;
  request: TRestRequest;
  response: TRestResponse;
begin
  Client := TRestClient.Create(URL);
  request := TRestRequest.Create(Self);
  response := TRestResponse.Create(Self);
  try
    if (URL <> authURL) and (token <> string.Empty) then
      Client.AddParameter('Authentication', 'bearer ' + token, pkHTTPHEADER,
        [poDoNotEncode]);
    request.Client := Client;
    request.response := response;
    request.method := method;
    if body <> string.Empty then
      request.AddBody(body, ctAPPLICATION_JSON);

    request.Execute;
  finally
    res.status := response.StatusCode;
    res.Content := response.Content;
    res.Headers.Assign(response.Headers);

    case method of
      rmPOST:
        methodS := 'POST';
      rmPUT:
        methodS := 'PUT';
      rmGET:
        methodS := 'GET';
      rmDELETE:
        methodS := 'DELETE';
      rmPATCH:
        methodS := 'PATCH';
    end;
    AddRequest(methodS, URL, body, res);

    response.Free;
    request.Free;
    Client.Free;
  end;
end;
/// TTiendaNubeRequest

/// TTiendaNube
// Destroy
constructor TTiendaNube.Create(AOwner: TComponent; token, store: string);
begin
  Create(AOwner);
  FToken := token;
  FStore := store;
end;

destructor TTiendaNube.Destroy;
begin
  inherited;
end;
// Constructor/destroy end

// private
/// <summary>Converts file to base64 strings</summary>
/// <param name="filename">File path to convert</param>
/// <returns>Base64 string. Empty string if there is an error</returns>
function TTiendaNube.ConvertToBase64(const filename: string): string;
var
  inStream: TStream;
  outStream: TStringStream;
begin
  Result := string.Empty;
  inStream := TFileStream.Create(filename, fmOpenRead);
  try
    outStream := TStringStream.Create;
    try
      TNetEncoding.Base64.Encode(inStream, outStream);
      Result := outStream.DataString;
    finally
      outStream.Free;
    end;
  finally
    inStream.Free;
  end;
end;

/// <summary>Check if string is integer</summary>
/// <param name="s">Integer string</param>
/// <returns>True if integer</returns>
function TTiendaNube.StrIsInt(s: string): Boolean;
var
  i: Int64;
begin
  Result := TryStrToInt64(s, i);
end;

/// <summary>
/// Check if string is float.
/// Decimal point should be '.'
/// '100.50' returns True
/// '100,50' returns False
/// </summary>
/// <param name="s">Float string</param>
/// <returns>True if float</returns>
function TTiendaNube.StrIsFloat(s: string): Boolean;
var
  i: Double;
  formatSettings: TFormatSettings;
begin
  formatSettings.DecimalSeparator := '.';
  Result := TryStrToFloat(s, i, formatSettings);
end;
// private end

// public
// auth
/// <summary>
/// Open browser to OAuth TiendaNube
/// </summary>
/// <param name="ClientID">TiendaNube partnet clientID</param>
/// <param name="ClientSecret">TiendaNube partnet client secret</param>
/// <returns>True if can obtain token and store</returns>
function TTiendaNube.InstallApp(ClientID, ClientSecret: string): Boolean;
var
  Code: string;
begin
  Result := False;
  if (ClientID <> string.Empty) and (ClientSecret <> string.Empty) then
  begin
    FormWB := TWebBrowserForm.CreateNew(Self);
    try
      Code := FormWB.ShowModalAppId(ClientID);
      if Code <> string.Empty then
      begin
        // Result := token <> string.Empty;
        Result := Auth(Code, ClientID, ClientSecret);
      end;
    finally
      FormWB.Free;
    end;
  end
end;

/// <summary>Get token and store auth</summary>
/// <param name="Code">OAuth code</param>
/// <param name="ClientID">TiendaNube partnet clientID</param>
/// <param name="ClientSecret">TiendaNube partnet client secret</param>
/// <returns>True if can obtain token and store</returns>
function TTiendaNube.Auth(Code, ClientID, ClientSecret: string): Boolean;
var
  body, resJSON: TJSONObject;
  res: TTiendaNubeResponse;
begin
  Result := False;
  body := TJSONObject.Create;
  res.Headers := TStringList.Create;
  try
    body.AddPair('client_id', ClientID);
    body.AddPair('client_secret', ClientSecret);
    body.AddPair('grant_type', 'authorization_code');
    body.AddPair('code', Code);
    DoRequest(authURL, string.Empty, res, rmPOST, body.ToString);
    if res.status = 200 then
    begin
      Result := true;
      resJSON := TJSONObject.ParseJSONValue(res.Content) as TJSONObject;
      try
        FToken := resJSON.GetValue<string>('access_token');
        FStore := resJSON.GetValue<string>('user_id');
      finally
        resJSON.Free;
      end;
    end;
  finally
    res.Headers.Free;
    body.Free;
  end;
end;
// auth

// products
/// <summary>
/// Get products if productId = string.Empty or not defined
/// Get product from productId
/// </summary>
/// <param name="productId">productId</param>
/// <returns>
/// JSON string
/// Empty string if is an error
/// </returns>
function TTiendaNube.GetProduct(productId: string = string.Empty): string;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  res.Content := string.Empty;
  if FStore <> string.Empty then
  begin
    URL := apiURL + FStore + '/products';
    if StrIsInt(productId) then
      URL := URL + '/' + productId;

    res.Headers := TStringList.Create;
    DoRequest(URL, FToken, res);
    if res.status = 200 then
      Result := res.Content;
    res.Headers.Free;
  end;
end;

/// <summary>
/// Get product from variant SKU
/// </summary>
/// <param name="productSKU">product SKU</param>
/// <returns>
/// JSON string
/// Empty string if is an error
/// </returns>
function TTiendaNube.GetProductSKU(sku: string): string;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  res.Content := string.Empty;
  if (FStore <> string.Empty) and (sku <> string.Empty) then
  begin
    URL := apiURL + FStore + '/products/sku/' + sku;
    res.Headers := TStringList.Create;
    DoRequest(URL, FToken, res);
    if res.status = 200 then
      Result := res.Content;
    res.Headers.Free;
  end;
end;

/// <summary>
/// Add product
/// </summary>
/// <param name="name">Product name. Required.</param>
/// <param name="categories">Categories ID list</param>
/// <param name="sku">Variant/Product SKU</param>
/// <param name="price">
/// Price. Float string. Decimal separator should be '.'
/// '100.50' is valid
/// '100,50' is not valid
/// </param>
/// <param name="stock">Stock (integer). If stock=string.Empty creates stock_managment=false variant</param>
/// <returns>
/// JSON string
/// Empty string if is an error
/// </returns>
function TTiendaNube.AddProduct(name: string; categories: TStringList = nil;
  sku: string = string.Empty; price: string = string.Empty;
  stock: string = string.Empty): string;
var
  URL: string;
  body, variant: TJSONObject;
  variants, cat: TJSONArray;
  res: TTiendaNubeResponse;
  i: Integer;
begin
  if name <> string.Empty then
  begin
    body := TJSONObject.Create;

    res.Headers := TStringList.Create;
    try
      body.AddPair('name', name);
      if categories <> nil then
      begin
        cat := TJSONArray.Create();
        for i := 0 to categories.Count - 1 do
          cat.Add(categories.Strings[i]);

        body.AddPair('categories', cat);
      end;
      if (sku <> string.Empty) or (price <> string.Empty) or
        (stock <> string.Empty) then
      begin
        variant := TJSONObject.Create;
        variants := TJSONArray.Create;
        if sku <> string.Empty then
          variant.AddPair('sku', sku);
        if StrIsFloat(price) then
          variant.AddPair('price', price);
        variants.Add(variant);
        if StrIsInt(stock) then
          variant.AddPair('stock', stock);
        body.AddPair('variants', variants);
      end;

      URL := apiURL + FStore + '/products';
      DoRequest(URL, FToken, res, rmPOST, body.ToString);
      if res.status = 201 then
        Result := res.Content;
    finally
      res.Headers.Free;
      body.Free;
    end;
  end;
end;

/// <summary>
/// Remove product
/// </summary>
/// <param name="productId">Product ID</param>
/// <returns>
/// True => product removed
/// </returns>
function TTiendaNube.RemoveProduct(productId: string): Boolean;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  Result := False;
  if (FStore <> string.Empty) and (StrIsInt(productId)) then
  begin
    URL := apiURL + FStore + '/products/' + productId;
    res.Headers := TStringList.Create;
    DoRequest(URL, FToken, res, rmDELETE);
    Result := res.status = 200;
    res.Headers.Free;
  end;
end;

/// <summary>
/// Modify product
/// </summary>
/// <param name="productId">Product ID</param>
/// <param name="name">Product name</param>
/// <param name="categories">Categories ID list</param>
/// <returns>
/// JSON string
/// Empty string if is an error
/// </returns>
function TTiendaNube.ModProductInternal(productId, name: string;
  categories: TStringList): string;
var
  URL: string;
  res: TTiendaNubeResponse;
  body: TJSONObject;
  cat: TJSONArray;
  i: Integer;
begin
  Result := string.Empty;
  body := TJSONObject.Create;
  res.Headers := TStringList.Create;
  try
    URL := apiURL + FStore + '/products/' + productId;
    if name <> string.Empty then
      body.AddPair('name', name);

    if categories <> nil then
    begin
      cat := TJSONArray.Create();
      for i := 0 to categories.Count - 1 do
        cat.Add(categories.Strings[i]);

      body.AddPair('categories', cat);
    end;
    DoRequest(URL, FToken, res, rmPUT, body.ToString);
    if res.status = 200 then
      Result := res.Content;
  finally
    res.Headers.Free;
    body.Free;
  end;
end;

/// <summary>
/// Modify product/variant
/// Only modifies not empty values
/// </summary>
/// <param name="productId">Product ID</param>
/// <param name="name">Product name</param>
/// <param name="sku">Variant/Product SKU</param>
/// <param name="price">
/// Price. Float string. Decimal separator should be '.'
/// '100.50' is valid
/// '100,50' is not valid
/// </param>
/// <param name="stock">Stock (integer)</param>
/// <param name="categories">Categories ID list</param>
/// <returns>
/// True => The product was modified
/// </returns>
function TTiendaNube.ModProduct(productId, name, sku, price, stock: string;
  categories: TStringList): Boolean;
var
  variantId, res: string;
  prod: TTiendaNubeProduct;
begin
  Result := False;
  res := string.Empty;
  if (FStore <> string.Empty) and (StrIsInt(productId)) then
  begin
    if (name <> string.Empty) or (categories <> nil) then
    begin
      res := ModProductInternal(productId, name, categories);
      Result := res <> string.Empty;
    end
    else if (sku <> string.Empty) or (price <> string.Empty) or
      (stock <> string.Empty) then
      res := GetProduct(productId);

    if (sku <> string.Empty) or (StrIsFloat(price)) or (StrIsInt(stock)) then
    begin
      prod := Deserialize<TTiendaNubeProduct>(res);
      if length(prod.variants) > 0 then
      begin
        variantId := prod.variants[0].id.ToString;
        res := ModVariant(productId, variantId, sku, price, stock);
        Result := res <> string.Empty;
      end;
    end;
  end;
end;

/// <summary>
/// Modify product/variant
/// Only modifies not empty values
/// </summary>
/// <param name="productId">Product ID</param>
/// <param name="name">Product name</param>
/// <param name="sku">Variant/Product SKU</param>
/// <param name="price">
/// Price. Float string. Decimal separator should be '.'
/// '100.50' is valid
/// '100,50' is not valid
/// </param>
/// <param name="stock">Stock (integer)</param>
/// <param name="categories">Categories ID list. Nil is empty list.</param>
/// <returns>
/// True => The product was modified
/// </returns>
function TTiendaNube.ModVariant(productId, variantId, sku, price,
  stock: string): string;
var
  URL: string;
  body: TJSONObject;
  res: TTiendaNubeResponse;
begin
  if (FStore <> string.Empty) and (StrIsInt(productId)) and (StrIsInt(variantId))
  then
  begin
    body := TJSONObject.Create;
    res.Headers := TStringList.Create;
    try
      if sku <> string.Empty then
        body.AddPair('sku', sku);
      if StrIsFloat(price) then
        body.AddPair('price', price);
      if StrIsInt(stock) then
        body.AddPair('stock', stock);
      URL := apiURL + FStore + '/products/' + productId + '/variants/' +
        variantId;

      DoRequest(URL, FToken, res, rmPUT, body.ToString);
      if res.status = 200 then
        Result := res.Content;
    finally
      res.Headers.Free;
      body.Free;
    end;
  end;
end;
// products end

// categories
/// <summary>
/// Get category list
/// </summary>
/// <returns>
/// JSON string
/// Empty string if is an error
/// </returns>
function TTiendaNube.GetCategories: string;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  if FStore <> string.Empty then
  begin
    res.Headers := TStringList.Create;
    URL := apiURL + FStore + '/categories';
    DoRequest(URL, FToken, res);
    if res.status = 200 then
      Result := res.Content;
    res.Headers.Free;
  end;
end;

/// <summary>
/// Add category
/// If parent. This would be sub-category
/// </summary>
/// <param name="category">Category name
/// <param name="parant">
/// Parent category id
/// Not required
/// </param>
/// <returns>
/// JSON string
/// Empty string if is an error
/// </returns>
function TTiendaNube.AddCategory(category: string;
  parent: string = string.Empty): string;
var
  URL: string;
  body: TJSONObject;
  res: TTiendaNubeResponse;
begin
  if FStore <> string.Empty then
  begin
    res.Headers := TStringList.Create;
    body := TJSONObject.Create;
    try
      body.AddPair('name', category);
      if (StrIsInt(parent)) then
        body.AddPair('parent', parent);

      URL := apiURL + FStore + '/categories';
      DoRequest(URL, FToken, res, rmPOST, body.ToString);
      if res.status = 200 then
        Result := res.Content;
    finally
      res.Headers.Free;
      body.Free;
    end;
  end;
end;

/// <summary>
/// Removes category
/// </summary>
/// <param name="category">Category ID</param>
/// <returns>
/// True => The category was removed
/// </returns>
function TTiendaNube.RemoveCategory(category: string): Boolean;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  Result := False;
  if (FStore <> string.Empty) and (StrIsInt(category)) then
  begin
    URL := apiURL + FStore + '/categories/' + category;
    DoRequest(URL, FToken, res, rmDELETE);
    Result := res.status = 200;
    res.Headers.Free;
  end;
end;
// categories end

// products image
/// <summary>
/// Add image to product
/// </summary>
/// <param name="productId">productId to add image</param>
/// <param name="Image">Image file path</param>
/// <returns>
/// imageId if created.
/// Empty string if is an error.
/// </returns>
function TTiendaNube.AddImage(productId, image: string): string;
var
  URL, Base64, filename: string;
  body, resJSON: TJSONObject;
  res: TTiendaNubeResponse;
begin
  if (FStore <> string.Empty) and (StrIsInt(productId)) then
  begin
    resJSON := nil;
    body := TJSONObject.Create;
    res.Headers := TStringList.Create;
    try
      Base64 := ConvertToBase64(image);
      URL := apiURL + FStore + '/products/' + productId + '/images';
      filename := ExtractFileName(image);
      body.AddPair('filename', filename);
      body.AddPair('attachment', Base64);
      DoRequest(URL, FToken, res, rmPOST, body.ToString);
      if res.status = 201 then
      begin
        resJSON := TJSONObject.ParseJSONValue(res.Content) as TJSONObject;
        Result := resJSON.GetValue<string>('id');
      end;
    finally
      if resJSON <> nil then
        resJSON.Free;
      res.Headers.Free;
      body.Free;
    end;
  end;
end;

/// <summary>
/// Removes image
/// </summary>
/// <param name="productId">productId to remove image</param>
/// <param name="ImageID">Image to remove</param>
/// <returns>
/// True => The image was removed
/// </returns>
function TTiendaNube.RemoveImage(productId, ImageID: string): Boolean;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  Result := False;
  if (FStore <> string.Empty) and (StrIsInt(productId)) and (StrIsInt(ImageID))
  then
  begin
    URL := apiURL + FStore + '/products/' + productId + '/images/' + ImageID;
    res.Headers := TStringList.Create;
    DoRequest(URL, FToken, res, rmDELETE);
    Result := res.status = 200;
    res.Headers.Free;
  end;
end;
// products image

// products
/// <summary>
/// Get orders if orderId = string.Empty or not defined
/// Get order from orderId
/// </summary>
/// <param name="orderId">orderId</param>
/// <returns>
/// JSON string
/// Empty string
function TTiendaNube.GetOrder(orderId: string = string.Empty): string;
var
  URL: string;
  res: TTiendaNubeResponse;
begin
  res.Content := string.Empty;
  if FStore <> string.Empty then
  begin
    URL := apiURL + FStore + '/orders';
    if StrIsInt(orderId) then
      URL := URL + '/' + orderId;

    res.Headers := TStringList.Create;
    DoRequest(URL, FToken, res);
    if res.status = 200 then
      Result := res.Content;
    res.Headers.Free;
  end;
end;

/// <summary>
/// Deserialize JSON to object
/// </summary>
/// <param name="s">JSON string</param>
/// <returns>
/// Returns T object from JSON
/// </returns>
function TTiendaNube.Deserialize<T>(s: string): T;
var
  d: TJSONSerializer;
begin
  Result := default (T);

  if not string.IsNullOrWhiteSpace(s) then
  begin
    d := TJSONSerializer.Create;
    try
      try
        Result := d.Deserialize<T>(s);
      except
      end;
    finally
      d.Free;
    end;
  end;
end;
// public end
/// TTiendaNube

/// TWebBrowserForm
constructor TWebBrowserForm.CreateNew(AOwner: TComponent; Dummy: Integer = 0);
begin
  inherited CreateNew(AOwner);

  width := Trunc(Screen.width * 0.7);
  height := Trunc(Screen.height * 0.6);
  position := poScreenCenter;

  FWebBrowser := TWebBrowser.Create(Self);
  FWebBrowser.Visible := true;
  FWebBrowser.Silent := true;
  FWebBrowser.SelectedEngine := EdgeIfAvailable;
  TWinControl(FWebBrowser).parent := Self;
  FWebBrowser.OnNavigateComplete2 := FOnDocumentComplete;
end;

/// <summary>
/// Navigate in TBrowser and Shows form modal
/// </summary>
/// <param name="appId">
/// clientId/appId from TiendaNube partner
/// </param>
function TWebBrowserForm.ShowModalAppId(appId: string): string;
var
  URL: String;
begin
  URL := 'https://tiendanube.com/apps/' + appId + '/authorize';
  FWebBrowser.Navigate(URL);
  FWebBrowser.Align := alClient;
  ShowModal;

  Result := FCode;
end;

procedure TWebBrowserForm.FOnDocumentComplete(ASender: TObject;
  const pDisp: IDispatch; const URL: OleVariant);
var
  URLstring: String;
  isLocalhost: Boolean;
  codePos: Integer;
begin
  try
    URLstring := URL;
    // URLstring := 'http://localhost/?code=1234567890';
    isLocalhost := URLstring.Contains('http://localhost');

    if isLocalhost then
    begin
      codePos := URLstring.IndexOf('code=');
      if codePos <> -1 then
      begin
        FCode := URLstring.Substring(codePos + 5);
        ModalResult := mrOk;
      end;
    end;
  finally
  end;
end;

destructor TWebBrowserForm.Destroy;
begin
  FWebBrowser.OnNavigateComplete2 := nil;
  FWebBrowser.Free;
  inherited;
end;
/// TWebBrowserForm

/// TBrowserEmulationAdjuster
class function TBrowserEmulation.GetExeName(): String;
begin
  Result := TPath.GetFileName(ParamStr(0));
end;

class procedure TBrowserEmulation.SetBrowserEmulation(const Value: DWORD);
const
  registryPath =
    'Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION';
var
  Registry: TRegistry;
  exeName: String;
begin
  exeName := GetExeName();

  Registry := TRegistry.Create(KEY_SET_VALUE);
  try
    Registry.RootKey := HKEY_CURRENT_USER;
    Win32Check(Registry.OpenKey(registryPath, true));
    Registry.WriteInteger(exeName, Value)
  finally
    Registry.Destroy();
  end;
end;

/// TBrowserEmulationAdjuster end
end.
