using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.ComponentModel;

namespace BandiDownloader
{
    class Program
    {
        // URL dell'API - questo deve essere l'endpoint corretto (modifica se necessario)
        private static readonly string apiUrl = "https://www.dati.lombardia.it/resource/bukx-h2uy.json";

        static async Task Main(string[] args)
        {
            Console.WriteLine("Scaricamento dati dei bandi in corso...");

            try
            {
                // Recupera i dati dall'API
                var bandiData = await GetBandiDataAsync();

                // Esporta i dati in un file Excel
                ExportToExcel(bandiData, "BandiRegioneLombardia.xlsx");

                Console.WriteLine("Dati scaricati con successo. File Excel generato: BandiRegioneLombardia.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore: {ex.Message}");
            }
        }

        // Metodo per ottenere i dati dall'API
        static async Task<List<Bando>> GetBandiDataAsync()
        {
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = await client.GetAsync(apiUrl);

                if (!response.IsSuccessStatusCode)
                    throw new Exception($"Errore nella chiamata API: {response.StatusCode}");

                string content = await response.Content.ReadAsStringAsync();

                // Deserializza il JSON in una lista di oggetti Bando
                return JsonConvert.DeserializeObject<List<Bando>>(content);
            }
        }
        static void ExportToExcel(List<Bando> bandiData, string fileName)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; // Necessario per EPPlus

            using (ExcelPackage excel = new ExcelPackage())
            {
                var workSheet = excel.Workbook.Worksheets.Add("Bandi");

                // Intestazioni delle colonne
                workSheet.Cells[1, 1].Value = "Codice Bando";
                workSheet.Cells[1, 2].Value = "Titolo Bando";
                workSheet.Cells[1, 3].Value = "Direzione Generale";
                workSheet.Cells[1, 4].Value = "Ente";
                workSheet.Cells[1, 5].Value = "Apertura Adesione";
                workSheet.Cells[1, 6].Value = "Chiusura Adesione";
                workSheet.Cells[1, 7].Value = "Tipo Strumento";
                workSheet.Cells[1, 8].Value = "Presentato";

                // Popolamento dei dati
                for (int i = 0; i < bandiData.Count; i++)
                {
                    workSheet.Cells[i + 2, 1].Value = bandiData[i].CodiceBando;
                    workSheet.Cells[i + 2, 2].Value = bandiData[i].TitoloBando;
                    workSheet.Cells[i + 2, 3].Value = bandiData[i].DirezioneGenerale;
                    workSheet.Cells[i + 2, 4].Value = bandiData[i].Ente;
                    workSheet.Cells[i + 2, 5].Value = bandiData[i].AperturaAdesione;
                    workSheet.Cells[i + 2, 6].Value = bandiData[i].ChiusuraAdesione;
                    workSheet.Cells[i + 2, 7].Value = bandiData[i].TipoStrumento;
                    workSheet.Cells[i + 2, 8].Value = bandiData[i].Presentato;
                }

                // Salvataggio del file Excel
                FileInfo excelFile = new FileInfo(fileName);
                excel.SaveAs(excelFile);
            }
        }

    }

    // Modello per il deserializzare il JSON dei bandi
    public class Bando
    {
        [JsonProperty("codice_bando")]
        public string CodiceBando { get; set; }

        [JsonProperty("titolo_bando")]
        public string TitoloBando { get; set; }

        [JsonProperty("direzione_generale")]
        public string DirezioneGenerale { get; set; }

        [JsonProperty("ente")]
        public string Ente { get; set; }

        [JsonProperty("apertura_adesione")]
        public string AperturaAdesione { get; set; }

        [JsonProperty("chiusura_adesione")]
        public string ChiusuraAdesione { get; set; }

        [JsonProperty("tipo_strumento")]
        public string TipoStrumento { get; set; }

        [JsonProperty("presentato")]
        public string Presentato { get; set; }
    }

}
