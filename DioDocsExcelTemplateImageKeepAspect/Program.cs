// See https://aka.ms/new-console-template for more information
using DioDocsExcelTemplateImageKeepAspect;
using GrapeCity.Documents.Excel;

Console.WriteLine("画像のアスペクト比を維持して出力");

// 新規ワークブックの作成
var workbook = new Workbook();

// 帳票テンプレートを読み込む
workbook.Open(@"resource\\Template_ImageTemplate.xlsx");

#region データの初期化
using var fs1 = new FileStream(@"resource\\AR.png", FileMode.Open);
var image1 = new byte[fs1.Length];
await fs1.ReadAsync(image1);

using var fs2 = new FileStream(@"resource\\ARJS.png", FileMode.Open);
var image2 = new byte[fs2.Length];
await fs2.ReadAsync(image2);

var datasource = new List<Products>();

var product1 = new Products();
product1.ProductName = "ActiveReports";
product1.ProductImage = image1;
datasource.Add(product1);

var product2 = new Products();
product2.ProductName = "ActiveReportsJS";
product2.ProductImage = image2;
datasource.Add(product2);
#endregion

// グローバル設定を追加
workbook.Names.Add("TemplateOptions.KeepLineSize", "true");

// データソースを追加
workbook.AddDataSource("ds", datasource);

// データを連結して帳票を作成
workbook.ProcessTemplate();

// xlsx ファイルに保存
workbook.Save("Result.xlsx");