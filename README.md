# Exportação de Declarações de IRRF para PDF
Macro VBA para gerar declarações individuais em PDF a partir de dados cadastrados no Excel.
## Objetivo
Automatizar o preenchimento de um modelo de declaração e a exportação de um PDF para cada investidor marcado para processamento.
O documento está configurado para o ano-calendário de 2025, incluindo informações de capital investido, AFAC, devolução de AFAC e distribuição de dividendos.
## Estrutura do arquivo
A pasta de trabalho deve conter as seguintes abas:
| Aba | Finalidade |
|---|---|
| `RESUMO` | Armazena os dados dos investidores e o controle de geração |
| `Exportar` | Contém o modelo utilizado para gerar os PDFs |
A macro utiliza `ThisWorkbook`, portanto as abas devem estar no mesmo arquivo que contém o código.
## Dados da aba RESUMO
A linha 1 deve conter os cabeçalhos. Os registros são processados a partir da linha 2.
| Coluna | Informação | Utilização |
|---|---|---|
| A | Referência para localizar a última linha | Determina o limite da leitura |
| B | Identificador | Compõe o nome do PDF |
| C | Nome do investidor | Preenche a declaração e compõe o nome do PDF |
| D | Identificação do investidor | Incluída no texto de apresentação |
| E | Filial | Identifica a empresa no documento |
| F | CNPJ da filial | Preenche a identificação da empresa |
| G | Saldo em 31/12/2024 | Capital investido no início do período |
| H | AFAC em 2025 | Aportes do período |
| I | Devolução de AFAC em 2025 | Capital devolvido no período |
| J | Saldo em 31/12/2025 | Capital investido no final do período |
| K | Distribuição de dividendos em 2025 | Dividendos do período |
| M | Status | Controla a geração do PDF |
## Modelo da aba Exportar
A macro preenche as seguintes células:
| Célula | Conteúdo |
|---|---|
| A8 | Texto de apresentação com nome e identificação do investidor |
| A10 | Nome da filial |
| A12 | CNPJ da filial |
| A14 | Saldo do capital investido em 31/12/2024 |
| A16 | AFAC em 2025 |
| A18 | Devolução de AFAC em 2025 |
| A20 | Saldo do capital investido em 31/12/2025 |
| A22 | Distribuição de dividendos em 2025 |
O modelo é reutilizado a cada registro. Ao final da execução, permanece preenchido com os dados do último investidor processado.
## Funcionamento
1. Localiza a última linha preenchida na coluna A da aba `RESUMO`.
2. Percorre os registros a partir da linha 2.
3. Processa somente as linhas com o status **GERAR** na coluna M.
4. Lê os dados do investidor e os valores financeiros.
5. Preenche o modelo na aba `Exportar`.
6. Exporta a declaração para PDF.
7. Altera o status do registro para `GERADO`.
8. Exibe o tempo total de execução em segundos.
A macro utiliza os valores informados na aba `RESUMO`. Não realiza cálculos de saldo nem validações dos valores financeiros.
## Pasta de destino
O caminho está definido diretamente no código:
```text
C:\Users\Eduardo Souza\Downloads\
```
Para utilizar em outro computador ou salvar em outra pasta, ajuste a variável `caminho`.
A macro não cria a pasta de destino.
## Nome dos arquivos
Os PDFs seguem o padrão:
```text
[Identificador da coluna B] - Declaração IRRF - [Nome da coluna C].pdf
```
Exemplo:
```text
001 - Declaração IRRF - Investidor Exemplo.pdf
```
O código não limpa os caracteres utilizados no nome do arquivo nem verifica previamente se já existe um PDF com o mesmo nome.
## Configurações do PDF
| Configuração | Definição |
|---|---|
| Formato | PDF |
| Qualidade | Padrão |
| Inclusão das propriedades do documento | Ativada |
| Respeitar a área de impressão | Ativado |
| Abrir o PDF após a geração | Desativado |
O layout e a área de impressão da aba `Exportar` devem estar preparados antes da execução.
## Controle de geração
| Status na coluna M | Ação |
|---|---|
| `GERAR` | Preenche o modelo e exporta o PDF |
| `GERADO` | Ignora o registro |
| Vazio ou outro conteúdo | Ignora o registro |
O código não remove espaços do status. Utilize o texto `GERAR` conforme o padrão definido.
Para gerar novamente um documento, altere o status do registro para `GERAR`.
A alteração para `GERADO` ocorre na planilha, mas a macro não salva automaticamente a pasta de trabalho.
## Envio de e-mail
O bloco de envio de e-mail está comentado. Na versão atual, nenhuma mensagem é criada ou enviada.
Entretanto, esta instrução permanece ativa no início da macro:
```vba
Set OutlookApp = CreateObject("Outlook.Application")
```
Assim, a execução ainda tenta inicializar a automação do Outlook, mesmo quando apenas gera PDFs.
A referência comentada para o destinatário utiliza a coluna J da aba `RESUMO`. Essa coluna está atualmente destinada ao saldo em 31/12/2025 e deve ser revista antes de ativar o envio.
## Limitações
- Não há tratamento de erros para falhas na exportação, no acesso à pasta ou na inicialização do Outlook.
- O caminho de saída está fixo no código.
- Os anos e as datas do documento estão fixos.
- Os registros abaixo da última célula preenchida na coluna A não são processados.
- Algumas variáveis utilizadas não estão declaradas no trecho original.
- O código não valida a identificação dos investidores, os CNPJs ou os valores.
- O nome do arquivo utiliza “Declaração IRRF”, enquanto o documento apresenta informações de capital, AFAC e dividendos. A nomenclatura deve ser validada pelo responsável antes da distribuição.
## Privacidade
O código contém um caminho local com nome de usuário. Antes de publicar no repositório, substitua esse caminho por um exemplo genérico.
Não inclua no repositório planilhas preenchidas ou PDFs com nomes, identificações e valores reais dos investidores.
