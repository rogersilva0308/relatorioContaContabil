package br.com.feltex.excel;

public class GerarRelatorio {
    public static void main(String[] args) {
        var lerArquivoRelatorioExcel = new LerArquivoRelatorioExcel();
        var registros = lerArquivoRelatorioExcel.lerArquivo("listaEntrada.xlsx");

        registros = lerArquivoRelatorioExcel.criarRegistrosPosterioresFaltantes(registros);

        var criaArquivoRelatorioExcel = new CriaArquivoRelatorioExcel();
        criaArquivoRelatorioExcel.criarArquivo("listaSaida.xlsx", registros);
    }
}
