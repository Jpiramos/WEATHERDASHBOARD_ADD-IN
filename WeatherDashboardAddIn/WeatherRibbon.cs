using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http;
using System.Windows.Forms;

namespace WeatherDashboardAddIn
{
    public partial class WeatherRibbon
    {



        private void WeatherRibbon_Load(object sender, RibbonUIEventArgs e)
        {


        }

        private async void btnBuscarClima_Click(object sender, RibbonControlEventArgs e)
        {
            string apiKey = "3a45bd3fab25bc775596be00e632641a";
            string cidade = txtCidade.Text;
            string cidadeCodificada = Uri.EscapeDataString(cidade);
            string apiUrl = $"https://api.openweathermap.org/data/2.5/weather?q={cidadeCodificada}&appid={apiKey}&units=metric&lang=pt";
            // Força o uso do TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            try
            {
                // Criar um handler sem proxy
                HttpClientHandler handler = new HttpClientHandler()
                {
                    Proxy = null, 
                    UseProxy = false  
                };

                // Usar o handler no HttpClient
                using (HttpClient client = new HttpClient(handler))
                {
                    HttpResponseMessage response = await client.GetAsync(apiUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResponse = await response.Content.ReadAsStringAsync();
                        dynamic dadosClima = JsonConvert.DeserializeObject(jsonResponse);

                        // Atualizar a planilha
                        var planilha = Globals.ThisAddIn.Application.ActiveSheet;

                        // Adicionando um título em negrito para as colunas
                        planilha.Cells[1, 1].Value2 = "Cidade";
                        planilha.Cells[1, 2].Value2 = cidade;

                        planilha.Cells[2, 1].Value2 = "Descrição";
                        planilha.Cells[2, 2].Value2 = dadosClima.weather[0].description;

                        planilha.Cells[3, 1].Value2 = "Temperatura (°C)";
                        planilha.Cells[3, 2].Value2 = dadosClima.main.temp;

                        planilha.Cells[4, 1].Value2 = "Umidade (%)";
                        planilha.Cells[4, 2].Value2 = dadosClima.main.humidity;

                        planilha.Columns["A:B"].AutoFit();

                        // Aplicar uma cor de fundo para os cabeçalhos
                        planilha.Range["A1:B1"].Interior.Color = System.Drawing.Color.LightSkyBlue;
                        planilha.Range["A1:B1"].Font.Bold = true;
                        planilha.Range["A1:B1"].Font.Color = System.Drawing.Color.White;

                        // Aplicar bordas para as células preenchidas
                        planilha.Range["A1:B4"].Borders.LineStyle = 1;

                        // Alterar o estilo das células com os dados
                        planilha.Range["A2:B4"].Font.Color = System.Drawing.Color.Black;
                        planilha.Range["A2:B4"].Font.Size = 12;

                        // Alinhar o texto centralizado nas células da segunda coluna
                        planilha.Range["A1:B4"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        planilha.Range["A2:B4"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                        // Adicionar uma linha de separação ao redor do conjunto de dados
                        planilha.Range["A1:B1"].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
                        planilha.Range["A1:B1"].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.Gray;

                        // Melhorar o visual das células de dados
                        planilha.Range["A2:B4"].Interior.Color = System.Drawing.Color.AliceBlue;

                    }
                    else
                    {
                        MessageBox.Show($"Erro ao buscar dados: Código de status HTTP {response.StatusCode}");
                        string errorMessage = await response.Content.ReadAsStringAsync();
                        MessageBox.Show($"Mensagem do servidor: {errorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                {
                    MessageBox.Show($"Erro: {ex.Message}\nDetalhes: {ex.InnerException.Message}");
                }
                else
                {
                    MessageBox.Show($"Erro: {ex.Message}");
                }
                Console.WriteLine($"Erro ao enviar a solicitação: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
