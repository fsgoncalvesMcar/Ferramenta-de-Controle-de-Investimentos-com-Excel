# Ferramenta-de-Controle-de-Investimentos-com-Excel

O projeto consiste na criação de uma ferramenta simples em Excel para simular investimentos em fundos imobiliários. A planilha permite calcular o valor total investido, o patrimônio acumulado e os dividendos mensais, ajudando o usuário a entender melhor o impacto de seus investimentos ao longo do tempo. Este modelo pode ser usado como base para futuras expansões e personalizações, oferecendo uma solução prática e acessível para investidores iniciantes.

1. Estrutura do Workbook
Transações: registro de todas as compras, vendas e taxas.

Preços: cotação atual de cada FII (pode ser manual ou via Power Query).

Dividendos: calendário de proventos (data de pagamento × valor por cota).

Resumo: principais métricas e gráficos de acompanhamento.

2. Aba “Transações”
Converta esse intervalo em Tabela (Ctrl + T) com colunas:

Data

FII (ticker, ex: “KNRI11”)

Tipo (“Compra”/“Venda”)

Qtde Cotas

Preço Unitário

Taxas (em R$)

Valor Líquido → =SE([@Tipo]="Compra",[@Qtde Cotas]*[@Preço Unitário]+[@Taxas],-[@Qtde Cotas]*[@Preço Unitário]-[@Taxas])

Use Validação de Dados para o ticker e tipo, evitando digitação livre.

3. Aba “Preços”
Crie uma Tabela com colunas: FII | Preço Atual

Atualização manual ou via Power Query (Dados → Obter Dados → Web/JSON de alguma API de FIIs).

4. Aba “Dividendos”
Tabela com: Data Pagto | FII | Valor por Cota.

Registre cada provento conforme anunciado pelas gestoras.

5. Cálculos na aba “Resumo”
Posição Atual por FII

Qtde Total: =SOMASES(Transações[Qtde Cotas];Transações[FII];A2;Transações[Tipo];"Compra") - SOMASES(...;Transações[Tipo];"Venda")

Custo Médio: =SOMASES(Transações[Valor Líquido];Transações[FII];A2;Transações[Tipo];"Compra")/QtdeTotal

Valor de Mercado

=QtdeTotal * PROCV(A2;Preços!$A:$B;2;0)

Total Investido

=SOMASES(Transações[Valor Líquido];Transações[Tipo];"Compra")

Patrimônio Acumulado

=SOMA(Resumo[Valor de Mercado])

Dividendos Mensais

Crie coluna “Mês” em Dividendos: =TEXTO([@Data Pagto];"yyyy-mm")

Use Tabela Dinâmica ou SOMASES para totalizar por mês: =SOMASES(Dividendos[Valor por Cota];Dividendos[FII];A2;Dividendos[Mês];E$1) * QtdeTotal

6. Dashboard e Gráficos
Gráfico de linhas do patrimônio ao longo do tempo (use datas e valor de mercado diário ou mensal).

Gráfico de colunas para dividendos recebidos por mês.

Segmentação de dados (Slicers) para filtrar por FII ou período.

7. Boas Práticas
Use Tabelas nomeadas: facilita fórmulas dinâmicas e atualização de intervalos.

Documente processos mais complexos com Comentários de Célula.

Proteja células de fórmulas para evitar alterações acidentais.

Mantenha versão de backup ou controle via GitHub (com arquivos .xlsx ou exportados como .csv).

8. Possíveis Extensões
Simulação de aportes periódicos: planilha que adiciona compras todo mês em X valor.

Reinvestimento automático de dividendos: macro ou fórmulas para reinvestir proventos.

Projeção de cenários: vars de crescimento de cotação e yields.

Dashboard no Power BI: conecte o Excel para visuais interativos.
