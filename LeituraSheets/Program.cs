using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;

namespace MyFirstProject
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "My First Project";
        static readonly string spreadsheetId = "1KxMpbnwCF_9WwfV1JXnf3FkXfzNhW9muvVFguu_aXeg";
        static readonly string sheet = "engenharia_de_software";
        static SheetsService service;

        static void Main(string[] args)
         {
            GoogleCredential credential;

            using (var stream =
                new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            WorkSheet();
            Show();
        }

        static void WorkSheet()
        {
            int count = 4;
            var range = $"{sheet}!A4:H27";
            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    double classes = 60, m, naf;
                    double faultLimit = (0.25 * classes);
                    int fault = Convert.ToInt32(row[2]);
                    int P1 = Convert.ToInt32(row[3]);
                    int P2 = Convert.ToInt32(row[4]);
                    int P3 = Convert.ToInt32(row[5]);

                    m = (P1 + P2 + P3) / 3;
                                    
                    if(fault > faultLimit)
                    {
                        UpdateNaf(0, count);
                        UpdateSituation("Reprovado por Falta", count);
                    }
                    else
                    {
                        if (m < 50)
                        {
                            UpdateNaf(0, count);
                            UpdateSituation("Reprovado por Nota", count);
                        }
                        else if ((m >= 50) && (m < 70))
                        {
                            naf = (50 * 2) - m;
                            UpdateNaf(Math.Round(naf), count);
                            UpdateSituation("Exame Final", count);
                        }
                        else if (m >= 70)
                        {
                            UpdateNaf(0, count);
                            UpdateSituation("Aprovado", count);
                        }
                    }                    
                    
                    count += 1;
                }
            }
            else
            {
                Console.WriteLine("Nenhum dado encontrado!");
            }
        }

        static void Show()
        {
            var range = $"{sheet}!A4:H27";
            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {                    
                    // Print columns A to F, which correspond to indices 0 and 4.
                    Console.WriteLine("{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7}", row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]);
                }
            }
            else
            {
                Console.WriteLine("Nenhum dado encontrado!");
            }
        }

        static void UpdateSituation(string situation, int line)
        {
            var range = $"{sheet}!G" + line;
            var valueRange = new ValueRange();

            var oblist = new List<object>() { situation };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.
                Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }

        static void UpdateNaf(double p, int line)
        {
            var range = $"{sheet}!H" + line;
            var valueRange = new ValueRange();

            var oblist = new List<object>() { p };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();
        }
    }
}