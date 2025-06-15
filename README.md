
# 📊 financialAnalysis – Automação de Conciliação Financeira em Excel (VBA)

Projeto de automação em **VBA para Excel**, desenvolvido para **facilitar e agilizar a conciliação de notas fiscais com os recebimentos financeiros** de uma empresa.

Essa solução realiza o cruzamento de dados entre relatórios de **boletos**, **cartões**, **PIX**, **transferências** e **extratos bancários**, proporcionando maior controle e precisão sobre as entradas financeiras.

---

## ✅ Funcionalidades principais

- Conciliação automática de boletos pagos com o extrato bancário
- Conciliação de recebimentos via cartão (crédito e débito) com o extrato
- Desmembramento dos recebimentos por venda (valor líquido e taxas)
- Exportação de planilhas por mês e por banco
- Geração de relatórios detalhados, prontos para análise ou auditoria

---

## 📁 Estrutura do Projeto

| Arquivo / Macro         | Função                                                                                   |
|-------------------------|------------------------------------------------------------------------------------------|
| **macroFormulario**     | Exporta uma nova planilha `.xlsx` por **mês** e por **banco**, com base nos filtros selecionados pelo usuário. |
| **subBoletoExt**         | Faz a conciliação dos **boletos emitidos** com o **extrato bancário**, identificando pagamentos e conferindo valores. |
| **subCartaooExt**        | Concilia os **recebimentos de cartões** com o **extrato bancário**, verificando os repasses das operadoras. |
| **subDestrinchamento**   | Detalha o **montante recebido por cartão**, separando **valor líquido de cada venda** e **taxas de maquininha**. |

---

## 📝 Pré-requisitos

- Microsoft Excel (com suporte a Macros - habilitação de conteúdo VBA)
- Permissão para habilitar macros no Excel
- Relatórios financeiros nos formatos esperados (extrato, boletos, cartões, etc.)

---

## 🚀 Como usar

1. **Clone ou baixe o repositório**:

```bash
git clone https://github.com/devlunah/financialAnalysis.git
```

2. **Abra o arquivo modelo**:

- Utilize o arquivo:  
`Planilha Conciliacoes_Modelo.xlsm`

3. **Habilite as macros no Excel**.

4. **Importe seus dados**:

- Cole os relatórios de extratos, boletos e cartões nas abas correspondentes.

5. **Execute as macros conforme necessário**:

| Caso de uso                            | Macro a executar          |
|----------------------------------------|---------------------------|
| Exportar nova planilha por mês/banco   | `macroFormulario`         |
| Conciliar boletos com extrato          | `subBoletoExt`            |
| Conciliar cartões com extrato          | `subCartaooExt`           |
| Desmembrar valores de cartão           | `subDestrinchamento`      |

---

## 📌 Exemplo de fluxo de trabalho

1. Importar o relatório de extrato bancário.
2. Importar os relatórios de boletos e cartões.
3. Rodar as macros de conciliação.
4. Gerar relatórios mensais/exportações conforme a necessidade.
5. Validar os resultados e utilizar para prestação de contas ou análise.

---

## ✅ Benefícios da automação

- Redução de tempo na conciliação financeira
- Diminuição de erros manuais
- Facilidade de auditoria
- Melhor controle de fluxo de caixa e recebíveis
- Processo replicável para diferentes bancos e meses

---

## 👨‍💻 Autor

**Lunah Carvalho**  
[GitHub Profile](https://github.com/devlunah)

---

## 📢 Contribuições e melhorias

Sinta-se à vontade para abrir issues ou enviar pull requests com sugestões de melhoria.

---

## 📃 Licença

Este projeto está sob a licença **MIT**.
