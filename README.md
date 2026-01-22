# Feltex-excel-util

Uma pequena utilidade Java que demonstra como trabalhar com arquivos Microsoft Excel usando o framework Apache POI. O projeto lê um arquivo de entrada com registros contábeis, processa os dados e gera um relatório/arquivo Excel de saída.

## Objetivo

Este repositório tem fins didáticos: mostrar leitura e escrita de planilhas Excel em Java, boas práticas mínimas de projeto, e integração com bibliotecas comuns (Apache POI, Lombok, Logback).

## O que a aplicação faz

- Lê um arquivo Excel de entrada (planilha com registros contábeis).
- Converte cada linha em objetos de domínio (`RegistroContabil`).
- Gera um arquivo Excel de relatório resumido/transformado com base nos registros lidos.

Arquivos de exemplo incluídos no repositório:

- `ListaEntrada.xlsx` — exemplo de planilha de entrada correta.
- `ListaEntrada_ERRADA.xlsx` — exemplo que contém entradas inválidas (útil para testes).
- `listaSaida.xlsx` — exemplo de saída gerada (modelo/resultado).

## Estrutura do projeto

- `src/main/java/br/com/feltex/excel/`
  - `LerArquivoRelatorioExcel.java` — lógica de leitura/parsing do Excel para objetos.
  - `CriaArquivoRelatorioExcel.java` — lógica de criação/escrita do arquivo Excel de saída.
  - `GerarRelatorio.java` — classe com método `main` que orquestra a leitura e gravação (ponto de entrada).
  - `modelo/RegistroContabil.java` — classe modelo que representa uma linha/registro contábil.
- `src/main/resources/logback.xml` — configuração de logging (Logback).
- `pom.xml` — arquivo de build Maven com dependências do projeto.

> Nota: o diretório `target/` contém classes compiladas e artefatos gerados pelo Maven.

## Tecnologias e bibliotecas usadas

- Java 11 (ou superior compatível)
- Maven — ferramenta de build e gerenciamento de dependências
- Apache POI — leitura e escrita de arquivos Microsoft Excel (.xls/.xlsx)
- Lombok — redução de boilerplate para modelos (getters/setters/builders)
- Logback — implementação de logging configurada via `logback.xml`
- JUnit (opcional) — framework de testes (caso queira adicionar testes)

## Como compilar e executar

1. Verifique se você tem Java 11+ e Maven instalados.
2. Abra um terminal na raiz do projeto (onde está o `pom.xml`).

Comandos sugeridos (Windows PowerShell):

```
# Compilar o projeto
mvn clean package

# Executar diretamente a classe principal usando a pasta target/classes no classpath
java -cp target\classes br.com.feltex.excel.GerarRelatorio
```

Observação: se preferir, você pode empacotar um JAR executável via `maven-assembly-plugin` ou `maven-shade-plugin` e executá-lo com `java -jar` — o projeto atual inclui `MANIFEST.MF` em `src/main/resources/META-INF` mas a geração de um JAR executável pode requerer configuração adicional no `pom.xml`.

## Parâmetros e arquivos de entrada

Por convenção, coloque o arquivo de entrada (por exemplo `ListaEntrada.xlsx`) na raiz do projeto ou informe o caminho completo na classe `GerarRelatorio` caso ela suporte argumentos. A aplicação vai:

- Ler o arquivo de entrada
- Validar/converter cada linha para `RegistroContabil`
- Registrar logs de progresso/erro via Logback
- Gerar um arquivo Excel de saída com o relatório

Verifique os exemplos `ListaEntrada.xlsx` e `ListaEntrada_ERRADA.xlsx` para entender o formato esperado das colunas.

## Possíveis melhorias e próximos passos

- Adicionar parsing de argumentos (ex.: `--input <arquivo> --output <arquivo>`).
- Implementar testes unitários para as classes de leitura e escrita.
- Gerar JAR executável no `pom.xml` para simplificar a execução.
- Tratar mais cenários de validação/erros nas planilhas (linhas vazias, formatos incorretos, células nulas).

## Contribuições

Contribuições são bem-vindas: abra uma issue descrevendo o que deseja melhorar ou envie um pull request com mudanças pequenas e bem documentadas.

## Referências

- Apache POI: https://poi.apache.org/
- Lombok: https://projectlombok.org/
- Logback: http://logback.qos.ch/

---
Arquivo gerado/atualizado automaticamente pelo assistente.
