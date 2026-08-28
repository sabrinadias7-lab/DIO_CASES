# 📊 Simulador de Investimentos em Fundos Imobiliários (FIIs)

Projeto desenvolvido como desafio prático do Bootcamp/curso da [DIO](https://www.dio.me/), com o objetivo de aplicar conceitos de Excel na construção de uma ferramenta de simulação de investimentos em Fundos Imobiliários (FIIs).

## 📌 Sobre o desafio

O objetivo era construir, em Excel, uma planilha que ajudasse o usuário a simular investimentos mensais em FIIs, respondendo perguntas como:

- Quanto investir por mês?
- Por quanto tempo?
- Qual a taxa de rendimento esperada?
- Quanto de patrimônio isso vai gerar no futuro?
- Quanto isso representa em dividendos mensais?

A proposta era automatizar esses cálculos, dando ao usuário uma visão clara do potencial retorno do seu investimento, com base em dados reais de mercado (rendimento médio de carteira, taxa mensal, etc.).

## 🎯 Objetivos de aprendizagem

- Criar ferramentas de simulação de investimentos em Excel;
- Aplicar cálculos financeiros como rendimento mensal e cálculo de dividendos;
- Documentar processos técnicos de forma clara e estruturada;
- Utilizar o GitHub como ferramenta de compartilhamento de documentação técnica.

## 🧠 Como a planilha funciona

O arquivo `Investimentos.xlsx` está dividido em duas abas:

### 🗂 Planilha1 — Simulador principal

**Configurações**
Bloco inicial com os dados de base do usuário:
- **Salário**
- **Rendimento de carteira** (percentual médio de rendimento mensal dos FIIs)
- **Sugestão de investimento** — calculada automaticamente como 30% do salário

**Investimento mensal**
Bloco onde o usuário informa os parâmetros da simulação:
- Quanto pretende investir por mês
- Por quantos anos pretende investir
- Qual a taxa de rendimento mensal esperada

A partir disso, a planilha calcula automaticamente:
- **Patrimônio acumulado**, usando a função financeira `VF` (Valor Futuro)
- **Dividendos mensais estimados**, multiplicando o patrimônio pelo rendimento de carteira

**Cenários comparativos**
Uma tabela que projeta o patrimônio acumulado e os dividendos mensais para diferentes horizontes de tempo (2, 5, 10, 20 e 30 anos), permitindo comparar o efeito do tempo sobre o investimento.

**Perfil de investidor e alocação por tipo de FII**
Com base em um **perfil de investidor** (Conservador, Moderado ou Agressivo), a planilha distribui automaticamente o valor mensal a ser investido entre 6 categorias de fundos imobiliários:

| Tipo de FII | Descrição resumida |
|---|---|
| PAPEL | Fundos de recebíveis (CRIs) |
| TIJOLO | Fundos de imóveis físicos |
| HÍBRIDOS | Combinação de papel e tijolo |
| FOFs | Fundos de fundos |
| DESENVOLVIMENTO | Fundos voltados a novos empreendimentos |
| HOTELARIAS | Fundos do setor hoteleiro |

O percentual sugerido para cada categoria muda conforme o perfil escolhido, buscando esses valores na aba de referência (Planilha2).

### 🗂 Planilha2 — Base de referência

Contém uma matriz com os percentuais recomendados de alocação para cada combinação de **Perfil x Tipo de FII** (ex: `Agressivo-PAPEL`, `Moderado-TIJOLO`), que serve como tabela de consulta para a Planilha1.

## 🧮 Fórmulas e conceitos de Excel aplicados

- **`VF` (Valor Futuro):** cálculo do patrimônio acumulado a partir de aportes mensais constantes, taxa de rendimento e prazo. *(Em inglês essa função se chama `FV` — é o mesmo nome salvo internamente no arquivo, o Excel só traduz o que aparece na tela.)*
- **`PROCV` (VLOOKUP):** busca dos percentuais de alocação na Planilha2, usando uma chave combinada.
- **Concatenação de texto (`&`):** criação de uma chave única (`Perfil & "-" & Tipo de FII`) para permitir a busca cruzada entre as duas categorias com um único PROCV.
- **Nomes de intervalo (Named Ranges):** células importantes foram nomeadas (`Salario`, `Aporte`, `Qtd_anos`, `Taxa_mensal`, `Rendimento_Carteira`, `Patrimonio`) para deixar as fórmulas mais legíveis, por exemplo:
  ```
  =VF(Taxa_mensal; Qtd_anos*12; Aporte)*-1
  =Patrimonio*Rendimento_Carteira
  ```
- **Referências absolutas e relativas (`$`):** usadas para travar células de parâmetros fixos ao copiar fórmulas entre linhas.

## 🖥 Como usar a planilha

1. Baixe o arquivo `Investimentos.xlsx` deste repositório;
2. Preencha os campos de entrada na aba **Planilha1**: salário, rendimento de carteira, valor a investir por mês, prazo em anos e taxa de rendimento mensal;
3. Escolha o **perfil de investidor** (Conservador, Moderado ou Agressivo);
4. Veja automaticamente o patrimônio projetado, os dividendos mensais e a distribuição sugerida do aporte entre os tipos de FII;
5. Consulte a tabela de **Cenários** para comparar o resultado em diferentes prazos.

## 🖼 Capturas de tela

<img width="1067" height="348" alt="image" src="https://github.com/user-attachments/assets/ed6454e3-94d5-4422-9c54-29967e87e6ee" />

<img width="1036" height="432" alt="image" src="https://github.com/user-attachments/assets/072e434c-783c-4bb5-9a1e-b798e0fa2402" />

<img width="1018" height="707" alt="image" src="https://github.com/user-attachments/assets/1eef5d08-1e25-4e89-a6d5-a7e683915701" />


## 🚀 Tecnologias utilizadas

- Microsoft Excel (fórmulas financeiras, PROCV, nomes de intervalo)

## 📚 Principais aprendizados

- Como estruturar uma planilha financeira de forma organizada, separando entradas do usuário, cálculos e resultados;
- Como usar a função `VF` para projetar valor futuro de aportes recorrentes;
- Como cruzar duas categorias de informação (perfil e tipo de fundo) com PROCV usando uma chave concatenada, evitando a necessidade de várias tabelas separadas;
- A importância de nomear intervalos para tornar fórmulas mais legíveis e fáceis de auditar.

## 🔗 Referências

- [Planilha de referência resolvida (DIO)](https://hermes.dio.me/files/assets/a04b81b1-8e35-4e72-aeb9-98aed8ed4403.xlsx)
- [Trilha/Bootcamp DIO](https://www.dio.me/)

## 👩‍💻 Autora

Desenvolvido por **[Sabrina Dias]** como parte do desafio de Excel da DIO.


