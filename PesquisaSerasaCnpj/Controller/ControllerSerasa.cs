using Microsoft.AspNetCore.Mvc;
using Services.SerasaService;

namespace Controller.ControllerSerasa;

[Controller]
public class ControllerSerasa : ControllerBase {

    private readonly SerasaService _serasaService;

    public ControllerSerasa() {
        _serasaService = new SerasaService();
    }

    /// <summary>
    /// Recupera os dados do Serasa
    /// </summary>
    /// <returns></returns>
    [HttpGet("RecuperarDadosSerasa")]
    public async Task<string> RecuperarDadosSerasa() {
        return await _serasaService.RecuperarDadosSerasa();
    } 

}