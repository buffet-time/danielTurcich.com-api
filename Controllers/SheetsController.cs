using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Web.Http.Cors;

namespace DanielTurcichComApi.Controllers
{
    [EnableCors(origins: "https://danielturcich.com/, localhost", headers: "*", methods: "*")]
    [ApiController]
    public class SheetsController : ControllerBase
    {
        static readonly string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static readonly string ApplicationName = "DanielTurcich.com API";

        //  pm2 start "dotnet DanielTurcichComApi.dll" --name api
        // pm2 delete api

        // Returns all rows in a specified spreadsheet if no row is passed
        // If row is passed then it only returns that specific row
        [HttpGet]
        [Route("sheets/{spreadsheetId}/{range}/{row?}")]
        public dynamic Get(string spreadsheetId, string range, int? row = null)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            var sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;

            if (row != null)
            {
                return values[row.GetValueOrDefault() - 2];
            }
            else
			{
                return values;
            }
        }
    }
}