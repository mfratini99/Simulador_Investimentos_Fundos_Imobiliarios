# Simulador de Investimentos em Fundos Imobiliários (FIIs)

Projeto desenvolvido como parte do desafio de Excel, com o objetivo de
aplicar conceitos de planilhas, fórmulas financeiras, referências, busca
de dados e visualização em uma ferramenta prática de simulação de
investimentos em Fundos Imobiliários (FIIs).

## 📌 Sobre o projeto

A planilha permite simular a evolução de um investimento mensal
considerando:

-   valor investido mensalmente;
-   quantidade de anos de investimento;
-   taxa de rendimento mensal;
-   rendimento da carteira;
-   patrimônio acumulado;
-   estimativa de dividendos mensais;
-   diferentes horizontes de investimento;
-   perfil de investidor e distribuição sugerida entre tipos de FII.

O projeto foi estruturado para que o usuário possa alterar os principais
parâmetros e observar automaticamente os resultados da simulação.

> **Observação:** os valores apresentados são provenientes de uma
> simulação educacional e não constituem recomendação de investimento.

## 🎯 Objetivos

-   Aplicar funções financeiras do Excel em uma situação prática;
-   Automatizar cálculos de crescimento patrimonial;
-   Estimar dividendos a partir do patrimônio acumulado;
-   Utilizar funções de busca e referências entre planilhas;
-   Trabalhar com diferentes cenários de longo prazo;
-   Organizar uma ferramenta de simulação de forma clara e funcional;
-   Utilizar o GitHub para documentar e compartilhar o projeto.

## 🧮 Principais recursos utilizados

### Função `VF`

A função financeira `vf` (Valor Futuro) é utilizada para calcular o
patrimônio projetado a partir de aportes mensais e de uma taxa de
rendimento.

Exemplo utilizado na planilha:

``` excel
=vf($D$19,$A24*12,$D$17*-1)
```

Essa lógica permite projetar o valor acumulado para diferentes períodos,
como 2, 5, 10, 20 e 30 anos.

### Cálculo de dividendos

A estimativa de dividendos mensais é calculada a partir do patrimônio
acumulado e do rendimento da carteira:

``` excel
=patrimonio*rendimento_carteira
```

### Função `PROCV`

A função `PROCV` é utilizada para buscar o percentual de alocação
correspondente ao perfil selecionado e ao tipo de FII:

``` excel
=PROCV($C$32&"-"&B36,Planilha2!A:D,4,FALSE)
```

A chave é formada pela combinação do perfil e do tipo de FII.

### Soma da distribuição

O valor total destinado aos diferentes tipos de FII é consolidado com:

``` excel
=SOMA(D36:D41)
```

## 📊 Estrutura da planilha

### Planilha1

É a área principal do simulador. Nela estão concentrados:

-   configurações iniciais;
-   investimento mensal;
-   prazo do investimento;
-   taxa de rendimento mensal;
-   patrimônio acumulado;
-   dividendos mensais;
-   cenários de 2, 5, 10, 20 e 30 anos;
-   perfil do investidor;
-   distribuição sugerida entre tipos de FII.

### Planilha2

Funciona como uma tabela de referência para os perfis de investimento.

São considerados três perfis:

-   **Conservador**
-   **Moderado**
-   **Agressivo**

E seis tipos de FII:

-   Papel
-   Tijolo
-   Híbridos
-   FOFs
-   Desenvolvimento
-   Hotelarias

Os percentuais de cada combinação são utilizados pela função `PROCV`
para alimentar automaticamente a distribuição apresentada na planilha
principal.

## 🔎 Exemplo de simulação

Na configuração utilizada no arquivo, foram considerados:

  Parâmetro                            Valor
  --------------------------- --------------
  Salário                       R\$ 2.000,00
  Investimento mensal             R\$ 200,00
  Prazo principal                     5 anos
  Taxa de rendimento mensal           1,079%
  Rendimento da carteira                0,6%

Com esses parâmetros, a simulação de 5 anos resulta em aproximadamente:

-   **Patrimônio acumulado:** R\$ 16.755,38
-   **Dividendos mensais estimados:** R\$ 100,53

Os valores são apenas exemplos calculados a partir dos parâmetros
presentes no arquivo e podem mudar quando o usuário altera as entradas.

## 📈 Cenários de longo prazo

Mantendo os mesmos parâmetros da simulação, os valores projetados são
aproximadamente:

    Período   Patrimônio projetado   Dividendos mensais estimados
  --------- ---------------------- ------------------------------
     2 anos           R\$ 5.445,53                      R\$ 32,67
     5 anos          R\$ 16.755,38                     R\$ 100,53
    10 anos          R\$ 48.656,84                     R\$ 291,94
    20 anos         R\$ 225.039,68                   R\$ 1.350,24
    30 anos         R\$ 864.433,93                   R\$ 5.186,60

Os resultados acima representam uma projeção matemática baseada nas
taxas informadas na planilha.

## 👤 Distribuição por perfil

Para o perfil **Moderado**, utilizado no exemplo, a distribuição
configurada é:

  Tipo de FII         Percentual
  ----------------- ------------
  Papel                      32%
  Tijolo                     35%
  Híbridos                    8%
  FOFs                        5%
  Desenvolvimento            10%
  Hotelarias                 10%
  **Total**             **100%**

Considerando um aporte mensal de R\$ 200,00, a distribuição
correspondente seria:

  Tipo de FII           Valor mensal
  ----------------- ----------------
  Papel                    R\$ 64,00
  Tijolo                   R\$ 70,00
  Híbridos                 R\$ 16,00
  FOFs                     R\$ 10,00
  Desenvolvimento          R\$ 20,00
  Hotelarias               R\$ 20,00
  **Total**           **R\$ 200,00**

## 🛠️ Conceitos de Excel aplicados

O projeto reúne diferentes conceitos trabalhados durante o curso,
incluindo:

-   funções financeiras;
-   referências absolutas e relativas;
-   fórmulas matemáticas;
-   funções de busca;
-   concatenação de informações;
-   tabelas de referência;
-   cenários de simulação;
-   organização e formatação de dados;
-   gráficos para visualização;
-   automação de cálculos através de fórmulas.

## 📁 Arquivos

-   `Book 2.xlsx` --- arquivo principal contendo o simulador
    desenvolvido no Excel.
-   `README.md` --- documentação do projeto.

## 🚀 Conclusão

O projeto demonstra a aplicação prática do Excel na construção de uma
ferramenta de simulação, combinando cálculos financeiros, automação por
fórmulas, tabelas de referência e visualização de informações.

A atividade também permitiu exercitar a documentação técnica de uma
solução e o uso do GitHub para disponibilização e compartilhamento do
projeto.
