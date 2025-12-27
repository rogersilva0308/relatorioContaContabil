package br.com.feltex.excel.modelo;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

import java.math.BigDecimal;

@Data
@Builder
@AllArgsConstructor
public class RegistroContabil {
    private Integer ano;
    private Integer mes;
    private Integer contaContabil;
    private String ic1;
    private String tipo1;
    private String ic2;
    private String tipo2;
    private String ic3;
    private String tipo3;
    private String ic4;
    private String tipo4;
    private String ic5;
    private String tipo5;
    private String ic6;
    private String tipo6;
    private BigDecimal saldoInicial;
    private BigDecimal valorDebito;
    private BigDecimal valorCredito;
    private BigDecimal saldoFinal;
    private String chaveTabela;
    private Integer linha;
    private Integer chave;
}
