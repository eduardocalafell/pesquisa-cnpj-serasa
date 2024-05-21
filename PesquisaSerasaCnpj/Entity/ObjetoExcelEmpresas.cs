using Ganss.Excel;

namespace Entity.ObjetoExcelEmpresas;

public class ObjetoExcelEmpresas {

    [Column("CNPJ")]
    public string Cnpj { get; set; } = String.Empty;

    [Column("Nome")]
    public string Nome { get; set; } = String.Empty;

    [Column("Fone")]
    public string? Fone { get; set; }

    [Column("Endereço")]
    public string? Endereco { get; set; }

    [Column("Email")]
    public string? Email { get; set; }
}