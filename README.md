# Mercado do Ponto | Voucher Analytics 2026

<p align="center">
  <img src="images/IMG_1377.PNG" alt="Mercado do Ponto" width="500">
</p>

## Visão geral

Projeto de análise financeira de vouchers do **Mercado do Ponto**, desenvolvido no **Microsoft Excel 365**. O trabalho transforma extratos de transações da **Alelo** e guias de reembolso da **VR** em indicadores para monitorar movimentação bruta, recebimentos líquidos, descontos, taxas efetivas e evolução mensal.

> **Importante:** a Alelo apresenta **transações de venda** e a VR apresenta **guias de reembolso**. As bases têm granularidades e critérios temporais diferentes. Portanto, a soma dos documentos **não é uma medida validada de faturamento por competência** e o valor médio das guias VR **não é ticket médio de compras**.

## Problema de negócio

Como um mercado de bairro pode entender o custo financeiro dos vouchers, acompanhar os valores líquidos recebidos e identificar oportunidades de melhoria na gestão de meios de pagamento usando os documentos disponíveis?

## Objetivos

- Consolidar e rastrear documentos de duas operadoras de voucher.
- Quantificar valores brutos, líquidos e diferenças identificadas.
- Calcular taxas efetivas observadas: `(valor_bruto - valor_liquido) / valor_bruto`.
- Comparar a evolução mensal segundo o critério de data disponível em cada fonte.
- Examinar o mix de cartões Alelo e identificar prioridades para revisão de custos.
- Estabelecer uma base inicial para conciliação financeira futura.

## Escopo e fontes

| Fonte | Arquivos | Unidade de análise | Critério mensal |
|---|---:|---|---|
| Alelo | 5 | Transação de venda | Data da venda |
| VR | 5 | Guia de reembolso | Data do pagamento |
| **Total** | **10** | **Unidades distintas** | **Não homogêneo** |

Arquivos identificados de **junho a outubro de 2026**. Datas e registros devem ser interpretados conforme o conteúdo dos extratos, não apenas pelo nome do arquivo.

## Principais indicadores

| Indicador | Resultado |
|---|---:|
| Valor bruto dos documentos | R$ 8.599,93 |
| Valor líquido dos documentos | R$ 7.934,07 |
| Diferença bruto menos líquido | R$ 665,86 |
| Taxa efetiva consolidada | 7,74% |
| Transações Alelo | 90 |
| Guias VR | 120 |

### Comparação entre operadoras

| Indicador | Alelo | VR |
|---|---:|---:|
| Registros | 90 transações | 120 guias |
| Bruto | R$ 2.365,57 | R$ 6.234,36 |
| Líquido | R$ 2.180,82 | R$ 5.753,25 |
| Diferença | R$ 184,75 | R$ 481,11 |
| Taxa efetiva | 7,81% | 7,72% |

**Nota:** as diferenças bruto–líquido podem incluir taxas, ajustes e outros descontos presentes nos documentos. Não se deve equipará-las automaticamente à taxa contratual de adquirência.

## Perguntas de negócio respondidas

1. **Quanto foi movimentado nos documentos?** R$ 8.599,93 brutos e R$ 7.934,07 líquidos, sem equivaler necessariamente ao faturamento de vendas do período.
2. **Qual o custo observado dos descontos?** R$ 665,86, ou 7,74% do bruto consolidado.
3. **Qual operadora tem maior taxa efetiva?** Alelo (7,81%), ante VR (7,72%), diferença de 0,09 ponto percentual; comparar com cautela.
4. **Onde está a maior concentração financeira?** VR, com aproximadamente 72,5% do bruto documental.
5. **Qual o ticket médio Alelo?** R$ 26,28 por transação, calculado como R$ 2.365,57 / 90.
6. **Qual o valor médio de guia VR?** R$ 51,95 por guia, não por compra.
7. **Qual mês apresentou maior valor combinado?** Setembro, cerca de R$ 2.029,17, conforme os critérios de agrupamento distintos das fontes.
8. **Como identificar modalidades mais custosas?** Usando `Mix_Alelo`, com valores e taxas por tipo de cartão.
9. **O que falta para conciliação completa?** Arquivos bancários de crédito, identificadores vinculáveis e regras de liquidação.
10. **Qual ação priorizar?** Revisar a composição dos descontos e estabelecer conciliação recorrente antes de negociar condições comerciais.

## Estrutura da pasta de trabalho

| Aba | Finalidade |
|---|---|
| `Dashboard` | KPIs e comparação mensal |
| `Resumo_Mensal` | Totais por mês e operadora |
| `Operadoras` | Comparativo bruto, líquido e taxa |
| `Mix_Alelo` | Tipos de cartão e indicadores Alelo |
| `Alelo_Vendas` | Base de transações importadas |
| `VR_Reembolsos` | Base de guias importadas |
| `Auditoria_Fontes` | Rastreabilidade e controles de importação |
| `LEIA_ME` | Guia de interpretação do Excel |

## Metodologia

1. **Leitura das fontes:** dez planilhas de duas operadoras.
2. **Padronização:** datas, valores, campos de identificação e arquivo de origem.
3. **Tratamento:** cálculo da diferença entre bruto e líquido em cada registro.
4. **Agregação:** Alelo por mês da venda; VR por mês de pagamento.
5. **Modelagem analítica:** tabelas detalhadas, resumos, comparativos e indicadores.
6. **Validação inicial:** contagem de registros, rastreabilidade por arquivo e verificação de diferenças negativas.
7. **Visualização:** painel gerencial em Excel com fórmulas e gráfico.

**Fórmulas-chave**

```excel
=Valor_Bruto-Valor_Liquido
=SEERRO(Descontos/Valor_Bruto;0)
=SEERRO(Valor_Bruto/Quantidade_Transacoes;0)
```

As fórmulas acima são representações conceituais com nomes descritivos; as referências reais estão na planilha.

## Como utilizar

1. Abra `Mercado_do_Ponto_Voucher_Analytics_2026.xlsx` no Excel 365.
2. Comece pela aba `LEIA_ME` para conhecer os critérios.
3. Consulte o `Dashboard` e o `Resumo_Mensal`.
4. Analise `Operadoras` e `Mix_Alelo` para investigar descontos.
5. Confira os registros nas abas detalhadas e a origem em `Auditoria_Fontes`.

## Tecnologias e competências

**Excel 365**, fórmulas de agregação, modelagem analítica, análise financeira, qualidade de dados, rastreabilidade, construção de KPIs, visualização e tradução de dados em decisões de negócio.

## Contexto do projeto

Projeto de portfólio desenvolvido a partir de uma necessidade prática de gestão financeira de um mercado de bairro. As interpretações e decisões propostas são hipóteses gerenciais a validar com os contratos e extratos bancários.

# Perguntas e respostas de negócio - Mercado do Ponto Voucher Analytics

**Escopo:** cinco arquivos Alelo e cinco arquivos VR, identificados de junho a outubro de 2026. Alelo = transações; VR = guias de reembolso. Os critérios de data não são equivalentes.

## 1. Quanto foi movimentado?
**Resposta:** R$ 8.599,93 brutos e R$ 7.934,07 líquidos nos documentos importados.

**Ação:** usar como referência documental, não como faturamento por competência.

## 2. Quanto foi descontado?
**Resposta:** R$ 665,86 (7,74% do bruto).

**Ação:** verificar os itens que compõem a diferença entre bruto e líquido, inclusive ajustes.

## 3. Qual operadora teve maior taxa efetiva?
**Resposta:** Alelo 7,81%; VR 7,72%. Diferença de 0,09 ponto percentual.

**Ação:** confrontar com contratos; a natureza dos documentos é diferente.

## 4. Qual fonte concentra maior valor bruto?
**Resposta:** VR, com R$ 6.234,36 (aproximadamente 72,5% do total documental); Alelo, R$ 2.365,57.

**Ação:** priorizar a revisão dos valores VR por materialidade, sem interpretar como participação de vendas.

## 5. Qual o mês de maior movimentação combinada?
**Resposta:** Setembro, aproximadamente R$ 2.029,17, sob agrupamentos distintos (venda Alelo e pagamento VR).

**Ação:** não concluir que foi o maior mês de vendas sem alinhar as datas de competência.

## 6. Qual o ticket médio Alelo?
**Resposta:** R$ 26,28 por transação (R$ 2.365,57 / 90).

**Ação:** comparar esse indicador com períodos futuros e outras formas de pagamento.

## 7. Qual o valor médio das guias VR?
**Resposta:** R$ 51,95 por guia (R$ 6.234,36 / 120).

**Ação:** utilizar como métrica administrativa, não como ticket médio de cliente.

## 8. Quais tipos de cartão Alelo têm maior custo proporcional?
**Resposta:** A aba `Mix_Alelo` detalha as taxas por modalidade; a conclusão deve ser extraída da ordenação das taxas efetivas e da representatividade de cada categoria.

**Ação:** investigar modalidades com alto custo e volume relevante.

## Conclusão

O projeto **Mercado do Ponto - Voucher Analytics 2026** demonstra como dados financeiros operacionais podem ser transformados em informações estratégicas para apoiar a gestão de um negócio varejista.

Por meio do Excel 365, foi possível consolidar e analisar documentos das operadoras Alelo e VR, estruturar indicadores financeiros e desenvolver um dashboard gerencial para acompanhar valores brutos, recebimentos líquidos, descontos e taxas efetivas.

A análise identificou uma movimentação bruta de **R$ 8.599,93**, com **R$ 665,86 em descontos**, representando uma taxa efetiva consolidada de **7,74%**. Esses resultados evidenciam a importância de monitorar os custos associados aos meios de pagamento e identificar oportunidades de melhoria na gestão financeira.

Além dos indicadores, o projeto reforçou a importância da **qualidade dos dados, padronização das informações e interpretação correta das diferentes fontes**, especialmente ao trabalhar com registros de vendas e guias de reembolso.

Mais do que construir uma planilha ou dashboard, o objetivo foi aplicar conceitos de **Data Analytics, Business Intelligence e análise de negócios** para responder a perguntas reais e transformar dados em informações úteis para a tomada de decisão.

Como evolução, o projeto poderá incorporar conciliação bancária automatizada, análises de rentabilidade, acompanhamento de prazos de recebimento e integração com ferramentas como Python, SQL e Power BI.

**O principal aprendizado é que o valor da análise de dados não está apenas nos indicadores apresentados, mas na capacidade de compreender o negócio, identificar problemas e apoiar decisões fundamentadas em dados confiáveis.**


