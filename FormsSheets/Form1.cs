using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace FormsSheets
{
    public partial class Form1 : Form
    {
        private static string ClientSecret = "client_secret.json";
        private static readonly string[] ScopesSheets = { SheetsService.Scope.Spreadsheets };
        private static readonly string AppName = "GoogleDriveAPIStart";
        private static readonly string SpreadsheetId = "1xFHgjLLInDenc3g1UpydfGET6nViTcsgu0v7Rr8fl7k";     //название таблицы
        private const string Range = "'Лист1'!C2";
        private static string Data = "7";
        public Form1()
        {
            InitializeComponent();
            Console.WriteLine("Get Creds");
            var credential = GetSheetCredentials();
            Console.WriteLine("Get service");
            var service = GetService(credential);
            Console.WriteLine("Fill data");
            FillSpreadsheet(service, SpreadsheetId, Data);
            Console.WriteLine("Getting result");
            string result = GetFirstCell(service, Range, SpreadsheetId);
            Console.WriteLine("recult: {0}", result);
            Console.WriteLine("Done");
        }
        private static UserCredential GetSheetCredentials()
        {
            using (var stream = new FileStream(ClientSecret, FileMode.Open, FileAccess.Read))
            {
                var creadPath = Path.Combine(Directory.GetCurrentDirectory(), "sheetsCreds.json");

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    ScopesSheets,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(creadPath, true)).Result;
            }
        }
        private static SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = AppName
            });
        }
        private static void FillSpreadsheet(SheetsService service, string spreadsheetId, string data)
        {

            List<Request> requests = new List<Request>();

            List<CellData> values = new List<CellData>();
            int n = 0;

            values.Add(new CellData
            {
                UserEnteredValue = new ExtendedValue
                {
                    StringValue = data
                }
            });

            requests.Add(
                new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Start = new GridCoordinate
                        {
                            SheetId = 0,
                            RowIndex = n,
                            ColumnIndex = 0
                        },
                        Rows = new List<RowData> { new RowData { Values = values } },
                        Fields = "userEnteredValue"
                    }
                }
           );

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };

            service.Spreadsheets.BatchUpdate(busr, spreadsheetId).Execute();
        }

        private static string GetFirstCell(SheetsService service, string range, string spreadsheetId)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();

            string result = null;

            foreach (var value in response.Values)
            {
                result += "" + value[0];
            }

            return result;
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var credential = GetSheetCredentials();
            var service = GetService(credential);
            FillSpreadsheet(service, SpreadsheetId, textBox1.Text);
        }
    }
}
