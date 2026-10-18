# 🧾 Organizador de Imposto de Renda (Excel)

Planilha em Excel que **reúne e organiza as informações essenciais para a declaração do Imposto de Renda (IRPF)**, com menu de navegação, validações automáticas, resumo visual e links rápidos. Tudo construído apenas com recursos nativos do Excel, sem macros (VBA).

Projeto desenvolvido como desafio do bootcamp da **[DIO](https://www.dio.me/)**.

> ⚠️ **Aviso:** esta planilha é uma ferramenta de apoio para organizar documentos e valores. Ela **não substitui** o programa oficial da Receita Federal nem orientação contábil. Confira sempre as regras do ano no site da Receita.

---

## 📋 Sumário

- [Visão geral](#-visão-geral)
- [Estrutura das abas](#-estrutura-das-abas)
- [Validações de dados](#-validações-de-dados)
- [Recursos do Excel utilizados](#-recursos-do-excel-utilizados)
- [Como usar](#-como-usar)
- [Convenções visuais](#-convenções-visuais)
- [Personalização](#-personalização)
- [Aprendizados](#-aprendizados)
- [Autor](#-autor)

---

## 🔎 Visão geral

O objetivo é funcionar como um **agregador de dados**: em vez de espalhar informes de bancos, holerites e recibos em vários lugares, o usuário lança tudo em um só arquivo, com entradas controladas e conferidas automaticamente, e chega à hora de declarar com os números em mãos.

**Principais funcionalidades**

- 🧭 **Menu de navegação** com botões para cada etapa e botão de retorno em todas as abas
- ✅ **Status automático** de preenchimento (ex.: "15 de 18 campos ok", "2 informe(s) pendente(s)")
- 🛡️ **Validações de entrada**: CPF, e-mail, datas, valores, listas suspensas
- 📊 **Resumo do ano** com cards, tabela mensal e gráfico de entradas
- 🔗 **Links rápidos** para os portais oficiais e atalho para abrir cada comprovante
- 🔒 **Fórmulas protegidas** (proteção de planilha sem senha) para evitar apagar cálculos sem querer

---

## 🗂️ Estrutura das abas

| Aba | Função |
|---|---|
| **INÍCIO** | Menu principal, ano-calendário, resumo do ano, entradas mês a mês (com gráfico) e links rápidos |
| **TITULAR** | Dados pessoais, endereço, contato e situação da declaração |
| **RENDIMENTOS** | Saldos em 31/12 e rendimentos do ano por banco, com controle de informes recebidos |
| **ENTRADAS** | Lançamento de holerites, notas de PJ e outras receitas (200 linhas), com resumo por categoria |
| **DEDUÇÕES** | Despesas dedutíveis (saúde, educação, previdência privada etc.) com totais por tipo |
| **LISTAS** *(oculta)* | Listas de apoio das listas suspensas (bancos, categorias, UFs etc.) |

### INÍCIO
Define o **ano-calendário** (célula amarela), que controla as validações de data de todo o arquivo. Mostra o status de cada etapa, os totais do ano e o gráfico de entradas por mês.

### TITULAR
Campos agrupados em Identificação, Endereço, Contato e Situação da declaração. Uma coluna de status confere cada campo (`✔`, `⚠ Obrigatório`, `✖ inválido`). O endereço completo e a idade são gerados por fórmula.

### RENDIMENTOS
Um banco por linha (15 linhas). Informa tipo de produto, saldo em 31/12, rendimentos no ano e se o informe foi recebido. A coluna **Abrir** transforma o nome do arquivo ou link do informe em atalho clicável.

### ENTRADAS
Data, categoria, fonte pagadora, valor bruto, IR retido e valor líquido (calculado). O painel lateral soma por categoria e confere se a soma bate com o total.

### DEDUÇÕES
Data, tipo, prestador, CPF/CNPJ, valor e comprovante. Esta aba apenas soma os valores; **os limites de dedução mudam a cada ano e não estão embutidos**.

---

## ✅ Validações de dados

| Campo | Regra |
|---|---|
| Ano-calendário | Número inteiro entre 2000 e 2100 |
| Nome | Mínimo de 5 caracteres |
| CPF | 11 dígitos numéricos + conferência dos dígitos verificadores (rejeita sequências repetidas como 111.111.111-11) |
| Data de nascimento | Entre 01/01/1900 e hoje |
| Título de eleitor, CEP, telefone, celular | Somente números, com tamanho coerente |
| Estado civil, UF, banco, produto, categoria, tipo de dedução, SIM/NÃO | Lista suspensa (valor fora da lista é bloqueado) |
| E-mail | Precisa conter `@` e ter tamanho mínimo |
| Datas de entradas e deduções | Dentro do ano-calendário definido em **INÍCIO** |
| Valores monetários | Numéricos e maiores ou iguais a zero (deduções: maior que zero) |
| IR retido | Não pode ser maior que o valor bruto da mesma linha |

Todas as validações exibem **mensagem de erro** explicando o que corrigir (estilo "Parar").

---

## 🧰 Recursos do Excel utilizados

- **Validação de dados:** lista, número inteiro, decimal, data, comprimento de texto e personalizada
- **Nomes definidos:** `ANO_BASE`, `BANCOS`, `CATEGORIAS`, `PRODUTOS`, `DEDUCOES`, `ESTADO_CIVIL`, `UFS`, `SIMNAO`
- **Fórmulas:** `IF`, `AND`, `OR`, `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `COUNT`, `LEFT`, `MID`, `MOD`, `TEXT`, `DATEDIF`, `TODAY`, `FIND`, `SEARCH`, `ROUND`, `HYPERLINK`
- **Hiperlinks internos e externos** para navegação e links rápidos
- **Formatação condicional** nos status (verde, amarelo e vermelho)
- **Gráfico de colunas** com as entradas mensais
- **Aba oculta** para dados de apoio
- **Proteção de planilha** (sem senha) com células de entrada desbloqueadas
- **Painéis congelados** e configuração de impressão (ajustar à largura da página)

### Destaque: validação do CPF por fórmula
O status do CPF recalcula os dois dígitos verificadores diretamente na célula, usando `SUMPRODUCT` com constantes de matriz e `MOD`, sem precisar de macro.

---

## 🚀 Como usar

1. Baixe o arquivo `Organizador_Imposto_de_Renda.xlsx` e abra no Excel.
2. Em **INÍCIO**, confira o **ano-calendário** (célula amarela) e ajuste se necessário.
3. Siga o menu, na ordem: **TITULAR → RENDIMENTOS → ENTRADAS → DEDUÇÕES**.
4. Preencha apenas as **células amarelas com texto azul**.
5. Acompanhe a coluna de status e o resumo em **INÍCIO** até tudo ficar ✔.
6. Use os números consolidados para preencher a declaração no programa oficial da Receita Federal.

> 💡 Os dados que vêm preenchidos são **exemplos fictícios** (a primeira linha de cada tabela e toda a aba TITULAR). Sobrescreva ou apague antes de usar.

---

## 🎨 Convenções visuais

| Estilo | Significado |
|---|---|
| Fundo amarelo, texto azul | Célula de preenchimento do usuário |
| Fundo branco ou cinza, texto preto | Fórmula (protegida) |
| Botões verde-água no menu | Navegação entre abas |
| `✔` verde, `⚠` amarelo, `✖` vermelho | Status de conferência |

---

## 🔧 Personalização

- **Desproteger as abas:** *Revisão → Desproteger Planilha* (não há senha).
- **Adicionar bancos, categorias ou tipos:** reexiba a aba **LISTAS**, escreva o novo item no fim da coluna e ajuste o intervalo em *Fórmulas → Gerenciador de Nomes*.
- **Aumentar o número de linhas:** copie a última linha da tabela e ajuste os intervalos das fórmulas de totais.

---

## 📚 Aprendizados

Neste desafio pratiquei:

- Modelagem de uma planilha como pequena aplicação, separando entrada, cálculo e apresentação
- Validação de dados para impedir erros na origem
- Fórmulas de busca, soma condicional e conferência (reconciliação entre abas)
- Navegação com hiperlinks e organização visual para quem não conhece o arquivo
- Documentação do projeto para uso por terceiros

---

## 👤 Autor

**Jonathan Moreira**
- LinkedIn: [seu-perfil](https://www.linkedin.com/in/jonathangmoreira/)

---

<sub>Projeto educacional desenvolvido no desafio da DIO. Os dados de exemplo são fictícios.</sub>
