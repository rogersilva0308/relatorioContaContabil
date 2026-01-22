package br.com.feltex.excel;

import br.com.feltex.excel.modelo.RegistroContabil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

//@Slf4j
public class LerArquivoRelatorioExcel {
    public List<RegistroContabil> lerArquivo(final String nomeArquivo) {
//        log.info("Lendo arquivo {}", nomeArquivo);
        System.out.println("Lendo arquivo: " +  nomeArquivo);
        List<RegistroContabil> registros = new ArrayList<>();

        try (FileInputStream excelFile = new FileInputStream(nomeArquivo)) {
            var workbook = new XSSFWorkbook(excelFile);
            var primeiraAba = workbook.getSheetAt(0);

            int contadorLinha = 0;
            int chaveAnterior = 0;
            int mesAnterior = 1;

            System.out.println("Planilha de entrada Original: " + nomeArquivo + " - Total de linhas: " +  (primeiraAba.getPhysicalNumberOfRows() - 1));

            for (Row linha : primeiraAba) {
                if (++contadorLinha == 1) continue; // Ignora cabeçalho

                int chaveAtual = Integer.parseInt(linha.getCell(21).getStringCellValue());
                int mesAtual = (int) linha.getCell(1).getNumericCellValue();

                if (chaveAnterior == chaveAtual) {
                    if ((mesAnterior + 1) != mesAtual) {
                        int novasLinhas = (mesAtual - (mesAnterior + 1));
                        criarRegistrosAnterioresFaltantes(linha, registros, novasLinhas);
                    }
                }

                adicionarRegistro(linha, registros);
                mesAnterior = mesAtual;
                chaveAnterior = chaveAtual;
                System.out.println("Linha: " +  contadorLinha);
            }

        } catch (FileNotFoundException e) {
//            log.error("Arquivo não encontrado {}", nomeArquivo);
            System.out.println("Arquivo não encontrado: " + nomeArquivo);
        } catch (IOException e) {
//            log.error("Erro ao processar o arquivo {}", nomeArquivo);
            System.out.println("Erro ao processar o arquivo: " + nomeArquivo);
        }
//        log.info("Total de registros lidos {}", registros.size());
        return registros;
    }

    private void adicionarRegistro(Row linha, List<RegistroContabil> registros) {
        var registro = RegistroContabil.builder()
                .ano((int) linha.getCell(0).getNumericCellValue())
                .mes((int) linha.getCell(1).getNumericCellValue())
                .contaContabil((int) linha.getCell(2).getNumericCellValue())
                .ic1((linha.getCell(3) == null ? "" : linha.getCell(3).getStringCellValue()))
                .tipo1((linha.getCell(4) == null ? "" : linha.getCell(4).getStringCellValue()))
                .ic2((linha.getCell(5) == null ? "" : linha.getCell(5).getStringCellValue()))
                .tipo2((linha.getCell(6) == null ? "" : linha.getCell(6).getStringCellValue()))
                .ic3((linha.getCell(7) == null ? "" : linha.getCell(7).getStringCellValue()))
                .tipo3((linha.getCell(8) == null ? "" : linha.getCell(8).getStringCellValue()))
                .ic4((linha.getCell(9) == null ? "" : linha.getCell(9).getStringCellValue()))
                .tipo4((linha.getCell(10) == null ? "" : linha.getCell(10).getStringCellValue()))
                .ic5((linha.getCell(11) == null ? "" : linha.getCell(11).getStringCellValue()))
                .tipo5((linha.getCell(12) == null ? "" : linha.getCell(12).getStringCellValue()))
                .ic6((linha.getCell(13) == null ? "" : linha.getCell(13).getStringCellValue()))
                .tipo6((linha.getCell(14) == null ? "" : linha.getCell(14).getStringCellValue()))
                .saldoInicial(BigDecimal.valueOf((double) linha.getCell(15).getNumericCellValue()))
                .valorDebito(BigDecimal.valueOf((double) linha.getCell(16).getNumericCellValue()))
                .valorCredito(BigDecimal.valueOf((double) linha.getCell(17).getNumericCellValue()))
                .saldoFinal(BigDecimal.valueOf((double) linha.getCell(18).getNumericCellValue()))
                .chaveTabela(linha.getCell(19).getStringCellValue())
//                .linha((int) linha.getCell(20).getNumericCellValue())
//                .chave((int) linha.getCell(21).getNumericCellValue())
                .linha(Integer.parseInt(linha.getCell(20).getStringCellValue()))
                .chave(Integer.parseInt(linha.getCell(21).getStringCellValue()))
                .build();

        registros.add(registro);
//        log.info("Lendo registro {}", registro);
    }

    private void criarRegistrosAnterioresFaltantes(Row linha, List<RegistroContabil> registros, int novasLinhas) {

        int mesInicial = ((int) linha.getCell(1).getNumericCellValue()) - novasLinhas;
        for (int i = 1; i <= novasLinhas; i++) {
            var registro = RegistroContabil.builder()
                    .ano((int) linha.getCell(0).getNumericCellValue())
                    .mes(mesInicial++)
                    .contaContabil((int) linha.getCell(2).getNumericCellValue())
                    .ic1((linha.getCell(3) == null ? "" : linha.getCell(3).getStringCellValue()))
                    .tipo1((linha.getCell(4) == null ? "" : linha.getCell(4).getStringCellValue()))
                    .ic2((linha.getCell(5) == null ? "" : linha.getCell(5).getStringCellValue()))
                    .tipo2((linha.getCell(6) == null ? "" : linha.getCell(6).getStringCellValue()))
                    .ic3((linha.getCell(7) == null ? "" : linha.getCell(7).getStringCellValue()))
                    .tipo3((linha.getCell(8) == null ? "" : linha.getCell(8).getStringCellValue()))
                    .ic4((linha.getCell(9) == null ? "" : linha.getCell(9).getStringCellValue()))
                    .tipo4((linha.getCell(10) == null ? "" : linha.getCell(10).getStringCellValue()))
                    .ic5((linha.getCell(11) == null ? "" : linha.getCell(11).getStringCellValue()))
                    .tipo5((linha.getCell(12) == null ? "" : linha.getCell(12).getStringCellValue()))
                    .ic6((linha.getCell(13) == null ? "" : linha.getCell(13).getStringCellValue()))
                    .tipo6((linha.getCell(14) == null ? "" : linha.getCell(14).getStringCellValue()))
                    .saldoInicial(BigDecimal.valueOf((double) linha.getCell(15).getNumericCellValue()))
                    .valorDebito(BigDecimal.ZERO)
                    .valorCredito(BigDecimal.ZERO)
                    .saldoFinal(BigDecimal.valueOf((double) linha.getCell(15).getNumericCellValue())) // saldoInicial e saldoFinal são iguais nos novos registros
                    .chaveTabela(linha.getCell(19).getStringCellValue())
                    .linha(Integer.parseInt(linha.getCell(20).getStringCellValue()))
                    .chave(Integer.parseInt(linha.getCell(21).getStringCellValue()))
                    .build();

            registros.add(registro);
//            log.info("Novo registro criado {}", registro);
        }
    }

    public List<RegistroContabil> criarRegistrosPosterioresFaltantes(List<RegistroContabil> registros) {
        List<RegistroContabil> novaLista = new ArrayList<>();
        RegistroContabil registroAnterior = null;
        int chaveAnterior = 0;
        int contInteracao = 0;
        int ultimoMesContabil = 0;

        Scanner kb = new Scanner(System.in);
        System.out.print("Digite o último mês contábil: ");

        while (true)
            try {
                ultimoMesContabil = Integer.parseInt(kb.nextLine());
                break;
            } catch (NumberFormatException nfe) {
                System.out.print("Try again: ");
            }

        if(ultimoMesContabil == 0){
            ultimoMesContabil = LocalDate.now().getMonthValue();
        }

        for (RegistroContabil registroAtual : registros) {
            contInteracao++;
            if (chaveAnterior != registroAtual.getChave()) {
                if (chaveAnterior != 0 && registroAnterior.getMes() != ultimoMesContabil && registroAnterior.getSaldoFinal().compareTo(BigDecimal.ZERO) != 0) {
                    int quantidadeMesesCriar = ultimoMesContabil - registroAnterior.getMes();
                    criarMesesPosterioresFaltantes(registroAnterior, quantidadeMesesCriar, novaLista);
                }
            }
            novaLista.add(registroAtual);

            if(contInteracao == registros.size() && registroAtual.getSaldoFinal().compareTo(BigDecimal.ZERO) != 0){
                int quantidadeMesesCriar = ultimoMesContabil - registroAtual.getMes();
                criarMesesPosterioresFaltantes(registroAtual, quantidadeMesesCriar, novaLista);
            }

            registroAnterior = registroAtual;
            chaveAnterior = registroAtual.getChave();
        }

        System.out.println("Planilha processada - Total de linhas: " +  novaLista.size());
        return novaLista;
    }

    private void criarMesesPosterioresFaltantes(RegistroContabil registroAnterior, int quantidadeMesesCriar, List<RegistroContabil> novaLista) {
        int mesInicial = registroAnterior.getMes();
        int linhaInicial = registroAnterior.getLinha();
        for (int i = 1; i <= quantidadeMesesCriar; i++) {
            var novoRegistro = RegistroContabil.builder()
                    .ano(registroAnterior.getAno())
                    .mes(++mesInicial)
                    .contaContabil(registroAnterior.getContaContabil())
                    .ic1(registroAnterior.getIc1())
                    .tipo1(registroAnterior.getTipo1())
                    .ic2(registroAnterior.getIc2())
                    .tipo2(registroAnterior.getTipo2())
                    .ic3(registroAnterior.getIc3())
                    .tipo3(registroAnterior.getTipo3())
                    .ic4(registroAnterior.getIc4())
                    .tipo4(registroAnterior.getTipo4())
                    .ic5(registroAnterior.getIc5())
                    .tipo5(registroAnterior.getTipo5())
                    .ic6(registroAnterior.getIc6())
                    .tipo6(registroAnterior.getTipo6())
                    .saldoInicial(registroAnterior.getSaldoFinal()) // saldoInicial e saldoFinal são iguais nos novos registros
                    .valorDebito(BigDecimal.ZERO)
                    .valorCredito(BigDecimal.ZERO)
                    .saldoFinal(registroAnterior.getSaldoFinal())
                    .chaveTabela(registroAnterior.getChaveTabela())
                    .linha(++linhaInicial)
                    .chave(registroAnterior.getChave())
                    .build();

            novaLista.add(novoRegistro);
//            log.info("Novo registro criado {}", novoRegistro);
        }
    }

    public void imprimirRegistro(){
//        System.out.println(" COLUNA 1 Ano " + (int) linha.getCell(0).getNumericCellValue());
//        System.out.println(" COLUNA 2 Mes " + (int) linha.getCell(1).getNumericCellValue());
//        System.out.println(" COLUNA 3 ContaContabil " + (int) linha.getCell(2).getNumericCellValue());
//        System.out.println(" COLUNA 4 IC1 " + linha.getCell(3).getStringCellValue());
//        System.out.println(" COLUNA 5 TIPO1 " + linha.getCell(4).getStringCellValue());
//        System.out.println(" COLUNA 6 IC2 " + linha.getCell(5).getStringCellValue());
//        System.out.println(" COLUNA 7 TIPO2 " + linha.getCell(6).getStringCellValue());
//        System.out.println(" COLUNA 8 IC3 " + linha.getCell(7).getStringCellValue());
//        System.out.println(" COLUNA 8 TIPO3 " + linha.getCell(8).getStringCellValue());
//        System.out.println(" COLUNA 10 IC4 " + linha.getCell(9).getStringCellValue());
//        System.out.println(" COLUNA 11 TIPO4 " + linha.getCell(10).getStringCellValue());
//        System.out.println(" COLUNA 12 IC5 " + (linha.getCell(11) == null ? "" : linha.getCell(11).getStringCellValue()));
//        System.out.println(" COLUNA 13 TIPO5 " + (linha.getCell(12) == null ? "" : linha.getCell(12).getStringCellValue()));
//        System.out.println(" COLUNA 14 IC6 " + (linha.getCell(13) == null ? "" : linha.getCell(13).getStringCellValue()));
//        System.out.println(" COLUNA 15 TIPO6 " + (linha.getCell(14) == null ? "" : linha.getCell(14).getStringCellValue()));
//        System.out.println(" COLUNA 16 SaldoInicial " + new BigDecimal((double) linha.getCell(15).getNumericCellValue()));
//        System.out.println(" COLUNA 17 ValorDebito " + new BigDecimal((double) linha.getCell(16).getNumericCellValue()));
//        System.out.println(" COLUNA 18 ValorCredito " + new BigDecimal((double) linha.getCell(17).getNumericCellValue()));
//        System.out.println(" COLUNA 19 SaldoFinal " + new BigDecimal((double) linha.getCell(18).getNumericCellValue()));
//        System.out.println(" COLUNA 20 ChaveTabela " + linha.getCell(19).getStringCellValue());
//        System.out.println(" COLUNA 21 Linha " + (int) linha.getCell(20).getNumericCellValue());
//        System.out.println(" COLUNA 22 Chave " + (int) linha.getCell(21).getNumericCellValue());
    }
}

