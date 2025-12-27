package br.com.feltex.excel;

import br.com.feltex.excel.modelo.RegistroContabil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

//@Slf4j
public class CriaArquivoRelatorioExcel {
    public void criarArquivo(final String nomeArquivo, final List<RegistroContabil> registros) {
//        log.info("Gerando o arquivo {}", nomeArquivo);
        System.out.println("Gerando o arquivo: " +  nomeArquivo);

        try (var workbook = new XSSFWorkbook();
             var outputStream = new FileOutputStream(nomeArquivo)) {
            var planilha = workbook.createSheet("Movimentações");
            int numeroDaLinha = 0;

            adicionarCabecalho(planilha, numeroDaLinha++);

            for (RegistroContabil registro : registros) {
                var linha = planilha.createRow(numeroDaLinha++);
                adicionarCelula(linha, 0, registro.getAno());
                adicionarCelula(linha, 1, registro.getMes());
                adicionarCelula(linha, 2, registro.getContaContabil());
                adicionarCelula(linha, 3, registro.getIc1());
                adicionarCelula(linha, 4, registro.getTipo1());
                adicionarCelula(linha, 5, registro.getIc2());
                adicionarCelula(linha, 6, registro.getTipo2());
                adicionarCelula(linha, 7, registro.getIc3());
                adicionarCelula(linha, 8, registro.getTipo3());
                adicionarCelula(linha, 9, registro.getIc4());
                adicionarCelula(linha, 10, registro.getTipo4());
                adicionarCelula(linha, 11, registro.getIc5());
                adicionarCelula(linha, 12, registro.getTipo5());
                adicionarCelula(linha, 13, registro.getIc6());
                adicionarCelula(linha, 14, registro.getTipo6());
                adicionarCelula(linha, 15, registro.getSaldoInicial());
                adicionarCelula(linha, 16, registro.getValorDebito());
                adicionarCelula(linha, 17, registro.getValorCredito());
                adicionarCelula(linha, 18, registro.getSaldoFinal());
                adicionarCelula(linha, 19, registro.getChaveTabela());
                adicionarCelula(linha, 20, registro.getLinha());
                adicionarCelula(linha, 21, registro.getChave());
            }

            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
//            log.error("Arquivo não encontrado: {}", nomeArquivo);
            System.out.println("Arquivo não encontrado: " + nomeArquivo);
        } catch (IOException e) {
//            log.error("Erro ao processar o arquivo: {} ", nomeArquivo);
            System.out.println("Erro ao processar o arquivo: " + nomeArquivo);
        }
//        log.info("Arquivo gerado com sucesso!");
        System.out.println("Arquivo gerado com sucesso!");
    }

    private void adicionarCabecalho(XSSFSheet planilha, int numeroLinha) {
        var linha = planilha.createRow(numeroLinha);

        adicionarCelula(linha, 0,"Ano");
        adicionarCelula(linha, 1,"Mes");
        adicionarCelula(linha, 2,"ContaContabil");
        adicionarCelula(linha, 3,"Ic1");
        adicionarCelula(linha, 4,"Tipo1");
        adicionarCelula(linha, 5,"Ic2");
        adicionarCelula(linha, 6,"Tipo2");
        adicionarCelula(linha, 7,"Ic3");
        adicionarCelula(linha, 8,"Tipo3");
        adicionarCelula(linha, 9,"Ic4");
        adicionarCelula(linha, 10,"Tipo4");
        adicionarCelula(linha, 11,"Ic5");
        adicionarCelula(linha, 12,"Tipo5");
        adicionarCelula(linha, 13,"Ic6");
        adicionarCelula(linha, 14,"Tipo6");
        adicionarCelula(linha, 15,"SaldoInicial");
        adicionarCelula(linha, 16,"ValorDebito");
        adicionarCelula(linha, 17,"ValorCredito");
        adicionarCelula(linha, 18,"SaldoFinal");
        adicionarCelula(linha, 19,"ChaveTabela");
        adicionarCelula(linha, 20,"Linha");
        adicionarCelula(linha, 21,"Chave");
    }

    private void adicionarCelula(Row linha, int coluna, String valor) {
        Cell cell = linha.createCell(coluna);
        cell.setCellValue(valor);
    }

    private void adicionarCelula(Row linha, int coluna, int valor) {
        Cell cell = linha.createCell(coluna);
        cell.setCellValue(valor);
    }
    private void adicionarCelula(Row linha, int coluna, BigDecimal valor) {
        Cell cell = linha.createCell(coluna);
        cell.setCellValue(valor.doubleValue());
    }
}
