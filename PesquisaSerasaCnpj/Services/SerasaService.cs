namespace Services.SerasaService;

using Entity.ObjetoExcelEmpresas;
using Ganss.Excel;

public class SerasaService
{

    IConfiguration _configuration;
    private string UrlSerasa { get; set; }
    private string UserSerasa { get; set; }
    private string SenhaSerasa { get; set; }
    private string ProdutoSerasa { get; set; }

    public SerasaService()
    {
        _configuration = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json")
            .Build();

        UrlSerasa = _configuration["Serasa:Url"]!;
        UserSerasa = _configuration["Serasa:Usuario"]!;
        SenhaSerasa = _configuration["Serasa:Senha"]!;
        ProdutoSerasa = _configuration["Serasa:Produto"]!;
    }

    private string TratarTelefoneSerasa(string retornoSerasa)
    {
        var ddd = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(49, 2).Trim();
        var primeiraMetade = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(52, 4).Trim();
        var segundaMetade = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(56, 4).Trim();
        return $"({ddd}) {primeiraMetade}-{segundaMetade}";
    }

    private string TratarEnderecoSerasa(string retornoSerasa)
    {
        var endereco = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010103"))!.Substring(6, 70).Trim();
        var cidade = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(6, 30).Trim();
        var uf = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(36, 2).Trim();
        var cep = retornoSerasa.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(39, 7).Trim();
        return $"{endereco}, {cidade} - {uf} - CEP: {cep}";
    }

    public async Task<string> RecuperarDadosSerasa()
    {

        HttpClient client = new HttpClient();
        string[] listaErros = new string[] { };
        var cnpjLoopAtual = string.Empty;
        var excel = new ExcelMapper(Environment.CurrentDirectory + "\\Input\\Base.xlsx");
        var empresas = excel.Fetch<ObjetoExcelEmpresas>().ToList();

        foreach (var empresa in empresas)
        {
            try
            {
                cnpjLoopAtual = empresa.Cnpj;
                var produto = ProdutoSerasa;
                var usuario = UserSerasa;
                var senha = SenhaSerasa;
                var cnpjTratado = empresa.Cnpj.Substring(0, empresa.Cnpj.Length - 6).PadLeft(9, '0');
                var url = UrlSerasa.Replace("{logon}", usuario).Replace("{senha}", senha).Replace("{produto}", produto).Replace("{documento}", cnpjTratado);
                string body = string.Empty;
                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    var response = client.Send(request);
                    response.EnsureSuccessStatusCode();
                    body = await response.Content.ReadAsStringAsync();
                    if (body.Contains("USUARIO BLOQUEADO")) throw new Exception("Usuário bloqueado!");

                    var empresaAtualizada = new ObjetoExcelEmpresas
                    {
                        Cnpj = empresa.Cnpj,
                        Nome = empresa.Nome,
                        Fone = TratarTelefoneSerasa(body),
                        Endereco = TratarEnderecoSerasa(body),
                        Email = body.Split("#L").FirstOrDefault(x => x.StartsWith("010104"))!.Substring(143, 70)
                    };

                    empresas.FirstOrDefault(x => x.Cnpj == empresa.Cnpj)!.Fone = empresaAtualizada.Fone;
                    empresas.FirstOrDefault(x => x.Cnpj == empresa.Cnpj)!.Endereco = empresaAtualizada.Endereco;
                    empresas.FirstOrDefault(x => x.Cnpj == empresa.Cnpj)!.Email = empresaAtualizada.Email;

                    excel.Save(Environment.CurrentDirectory + "\\Input\\Base.xlsx", empresas);

                    Thread.Sleep(5000);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                listaErros.Append($"CNPJ: {cnpjLoopAtual}");
                continue;
            }
        }

        return "Dados integrados com sucesso! Cnpj's com erro: " + string.Join(", ", listaErros);
    }

}