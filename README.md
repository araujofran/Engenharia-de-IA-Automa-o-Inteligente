

## 🚀 **README.md — Projeto: Condições Salariais com Automation Anywhere**

# 🤖 Projeto de Automação: Atualização Salarial com Condicionais no Excel

**Autor:** Francisco Ferreira de Araujo  
GitHub: https://github.com/araujofran  
LinkedIn: https://www.linkedin.com/in/francisco-ferreira-de-araujo-1b432033/  



## 🎯 Objetivo do Projeto

Este projeto tem como objetivo **automatizar o cálculo de novos salários** de colaboradores com base em **regras condicionais**, utilizando o **Automation Anywhere A360** integrado com um arquivo **Excel**.

O bot:
- Lê os dados de uma planilha `.xlsx`
- Limpa e converte os valores salariais para numérico
- Analisa regras condicionais por empresa + faixa salarial
- Aplica percentuais de aumento
- Escreve o novo salário diretamente no Excel



## 📊 Estrutura dos Dados da Planilha

| First Name | Last Name | Company | Current Salary | New Salary |
|-----------|-----------|---------|----------------|------------|
| Texto     | Texto     | Texto   | $?xx,xxx.xx    | Calculado pelo Bot |

O **bot preenche a coluna "New Salary"** com os valores calculados.



## 🧠 Regras Condicionais Utilizadas

| Empresa | Salário < 40.000 | Salário ≥ 40.000 |
|--------|------------------|------------------|
| **NanoTech** | +10% | +5% |
| **PicoTech** | +15% | +10% |
| **MicroTech** | +5% | +5% |



## 🏗️ Fluxo do Bot — Passo a Passo Explicado

A seguir, cada ação utilizada no projeto é detalhada com o motivo e o resultado esperado.



### ▶️ **1. Abrir o Excel**


Excel Advanced: Open "C:\Bots\universidadedevRPA\salaries.xlsx"


- Carrega o arquivo onde os dados serão lidos e escritos.



### 📍 **2. Posicionar na coluna New Salary**


Excel Advanced: Go to specific cell "E2"


- Preparação para escrita dos novos salários.



### 🔁 **3. Loop por cada linha da planilha**


Loop: For each row in worksheet assign to $rExcelRow$


- Permite processar linha a linha.
- A variável `$rExcelRow$` contém os dados da linha atual.



### 🔡 **4–5. Remover caracteres inválidos do salário**
#### 4️⃣ Substituir `$` por vazio


String: Substituir "$" por "" em $rExcelRow("Current Salary")$
Saída: $sSalaryStep1$


#### 5️⃣ Substituir vírgula `,` por vazio


String: Substituir "," por "" em $sSalaryStep1$
Saída: $sSalaryClean$



📌 Objetivo: deixar o número em formato **45678.00**, sem símbolo de moeda.



### 🔢 **6. Converter salário para Número**


String: Em número
Entrada: $sSalaryClean$
Saída: $nSalary$


Agora já é possível comparar valores numéricos no **IF**.



### ⚙️ **7–15: Aplicação das Regras Condicionais**

#### 🟣 PicoTech


If $rExcelRow("Company")$ = "PicoTech"
If $nSalary$ < 40000
$sSalary$ = $nSalary$ * 1.15
Else
$sSalary$ = $nSalary$ * 1.10



#### 🔵 MicroTech


Else If $rExcelRow("Company")$ = "MicroTech"
$sSalary$ = $nSalary$ * 1.05



#### 🟢 NanoTech


Else If $rExcelRow("Company")$ = "NanoTech"
If $nSalary$ < 40000
$sSalary$ = $nSalary$ * 1.10
Else
$sSalary$ = $nSalary$ * 1.05


📌 A variável utilizada para escrita é sempre `$sSalary$`.



### ✍️ **16. Escrever o novo salário no Excel**

Excel Advanced: Set cell = $sSalary$





### ⬇️ **17. Avançar uma linha**


Excel Advanced: Go to next row



## ❗ Problemas Encontrados e Soluções

| Problema | Causa | Solução |
|---------|------|---------|
| Erro ao converter salário para número | Valor continha `$` e `,` | Foi necessário usar **duas ações Substituir** antes |
| `$Year$`, `$Month$` não apareciam corretamente | Variáveis do sistema não são aceitas em concatenação | Implementado método alternativo baseado em substituição |
| Campo “To number” rejeitava input | O Excel retornava string inválida | Criadas variáveis intermediárias |

🔍 Todos os erros foram corrigidos com uso de:
- Variáveis intermediárias
- Sanitização de strings
- Execução em ordem correta das ações



## 📌 Tecnologias Utilizadas

- **Automation Anywhere A360**
- Módulo **Excel Advanced**
- Ações da categoria **String**
- Condições **If / Else If / Else**
- Estrutura de **Loop**



### 🧾 Resultado Final — Salários Calculados pelo Bot

Abaixo está o resultado gerado após a aplicação automática das regras condicionais:

| First Name | Company   | Current Salary | Novo Percentual Aplicado | New Salary |
| ---------- | --------- | -------------- | ------------------------ | ---------- |
| Addison    | NanoTech  | $45,678.00     | 5%                       | $47,961.90 |
| Dan        | PicoTech  | $34,567.00     | 15%                      | $39,752.05 |
| Gabriel    | PicoTech  | $54,321.00     | 10%                      | $59,753.10 |
| Irma       | NanoTech  | $24,680.00     | 10%                      | $27,148.00 |
| Jack       | MicroTech | $46,825.00     | 5%                       | $49,166.25 |
| Leon       | PicoTech  | $36,951.00     | 15%                      | $42,593.65 |
| Luke       | NanoTech  | $45,682.00     | 5%                       | $47,966.10 |
| Rithy      | PicoTech  | $35,791.00     | 15%                      | $41,159.65 |
| Sarah      | NanoTech  | $27,381.00     | 10%                      | $30,119.10 |
| Seth       | NanoTech  | $38,192.00     | 10%                      | $42,011.20 |

📌 Essa tabela é gerada automaticamente ao rodar o bot.


### 🧾 Resultado Final — Salários Calculados pelo Bot

Após a execução completa da automação, os novos salários foram calculados conforme as regras condicionais:

| First Name | Company | Current Salary | Novo Percentual Aplicado | New Salary |
|-----------|---------|----------------|-------------------------|-----------|
| Addison | NanoTech | $45,678.00 | 5% | $47,961.90 |
| Dan | PicoTech | $34,567.00 | 15% | $39,752.05 |
| Gabriel | PicoTech | $54,321.00 | 10% | $59,753.10 |
| Irma | NanoTech | $24,680.00 | 10% | $27,148.00 |
| Jack | MicroTech | $46,825.00 | 5% | $49,166.25 |
| Leon | PicoTech | $36,951.00 | 15% | $42,593.65 |
| Luke | NanoTech | $45,682.00 | 5% | $47,966.10 |
| Rithy | PicoTech | $35,791.00 | 15% | $41,159.65 |
| Sarah | NanoTech | $27,381.00 | 10% | $30,119.10 |
| Seth | NanoTech | $38,192.00 | 10% | $42,011.20 |

> Tabela gerada automaticamente a partir da execução do bot desenvolvido no Automation Anywhere.


## 🧑‍💻 Autor

**Francisco Ferreira de Araujo**  
🔗 GitHub: https://github.com/araujofran  
🔗 LinkedIn: https://www.linkedin.com/in/francisco-ferreira-de-araujo-1b432033/  



Se este projeto foi útil para você, ⭐ **considere deixar uma estrela no repositório!**



