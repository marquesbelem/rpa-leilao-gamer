using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace rpa_leilao_gamer
{
    public class AutomationWeb
    {
        private IWebDriver m_Driver;
        private const int ImplicitWait = 3; 
        public void Init()
        {
            string url;
            DisplayUrl(out url);
            int pages;
            DisplayPages(out pages);

            Console.WriteLine("$$$ Programa sendo executado ... aguarde e não feche essa janela $$$");

            try
            {
                m_Driver = new ChromeDriver();

                List<LoteModel> lotesByPages = new List<LoteModel>();

                for (int i = 1; i <= pages; i++)
                {
                    var fullUrl = $"{url}?page={i}";
                    m_Driver.Navigate().GoToUrl(fullUrl);
                    m_Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(ImplicitWait);
                    lotesByPages.AddRange(GetLotesInPage());
                }

                SaveExcel(lotesByPages);
                Console.WriteLine("$$$ Dados salvos no Excel com sucesso! $$$");
                m_Driver.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao executar {ex.Message}");
            }
        }

        void DisplayPages(out int pages)
        {
            Console.WriteLine("$$$ Digite quantidade de paginas $$$");
            pages = Int32.Parse(Console.ReadLine());
            if (pages <= 0)
            {
                Console.WriteLine("A quantidade de paginas tem que ser maior de 0. \n Encerrando o programa.");
                return;
            }
        }

        void DisplayUrl(out string url)
        {
            Console.WriteLine("$$$ Digite a URL da página principal $$$");
            Console.WriteLine("$$$              ATENÇÃO             $$$");
            Console.WriteLine("$$$ Não colocar o final da URl '?page=29' $$$");
            Console.WriteLine("$$$ Exemplo de URL 'https://www.nossoleilao.com.br/leilao/315/lotes' $$$");

            url = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(url))
            {
                Console.WriteLine("A URL não pode estar vazia. \n Encerrando o programa.");
                return;
            }
        }

        private List<LoteModel> GetLotesInPage()
        {
            IReadOnlyCollection<IWebElement> loteElements = m_Driver.FindElements(By.ClassName("lote"));
            List<LoteModel> lotes = new List<LoteModel>();

            foreach (IWebElement element in loteElements)
            {
                IWebElement datailLink = element.FindElement(By.CssSelector("a.btn.btn-block.btn-dark"));
                string href = datailLink.GetDomAttribute("href");
                var model = CreateModel(element.Text, href);
                lotes.Add(model);
            }

            return lotes;
        }

        LoteModel CreateModel(string texto, string link)
        {
            string[] linhas = texto.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            string maxCast = "";

            for (int i = 0; i < linhas.Length; i++)
            {
                if (linhas[i].Contains("R$"))
                {
                    maxCast = linhas[i];
                    break;
                }
            }

            var model = new LoteModel(linhas[0], linhas[1], maxCast, link);
            return model;
        }

        void SaveExcel(List<LoteModel> dados)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Lotes");
                worksheet.Cells[1, 1].Value = "Lote";
                worksheet.Cells[1, 2].Value = "Link";
                worksheet.Cells[1, 3].Value = "Descrição";
                worksheet.Cells[1, 4].Value = "Maior Lance";

                for (int i = 0; i < dados.Count; i++)
                {
                    var dado = dados[i];
                    worksheet.Cells[i + 2, 1].Value = dado.Lote;
                    worksheet.Cells[i + 2, 2].Value = dado.Link;
                    worksheet.Cells[i + 2, 3].Value = dado.Description;
                    worksheet.Cells[i + 2, 4].Value = dado.MaxCast;
                }

                FileInfo fileInfo = new FileInfo("Lotes.xlsx");
                package.SaveAs(fileInfo);
            }
        }
    }
}
