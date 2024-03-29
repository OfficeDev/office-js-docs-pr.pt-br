---
title: Chamar APIs JavaScript do Excel de uma função personalizada
description: Saiba quais APIs JavaScript do Excel você pode chamar de sua função personalizada.
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d1cbf6d07e4ede5b8309e899828f8f1d8ad1fa0
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464829"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Chamar APIs JavaScript do Excel de uma função personalizada

Chame APIs JavaScript do Excel de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos. Chamar APIs JavaScript do Excel por meio de uma função personalizada pode ser útil quando:

- Uma função personalizada precisa obter informações do Excel antes do cálculo. Essas informações podem incluir propriedades do documento, formatos de intervalo, partes XML personalizadas, um nome de pasta de trabalho ou outras informações específicas do Excel.
- Uma função personalizada definirá o formato de número da célula para os valores retornados após o cálculo.

> [!IMPORTANT]
> Para chamar APIs JavaScript do Excel de sua função personalizada, você precisará usar um [runtime compartilhado](../testing/runtimes.md#shared-runtime). Consulte [Configurar seu Suplemento do Office para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

## <a name="code-sample"></a>Exemplo de código

Para chamar APIs JavaScript do Excel de uma função personalizada, primeiro você precisa de um contexto. Use o [objeto Excel.RequestContext](/javascript/api/excel/excel.requestcontext) para obter um contexto. Em seguida, use o contexto para chamar as APIs necessárias na pasta de trabalho.

O exemplo de código a seguir mostra como usar `Excel.RequestContext` para obter um valor de uma célula na pasta de trabalho. Neste exemplo, o parâmetro `address` é passado para o método [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) da API JavaScript do Excel e deve ser inserido como uma cadeia de caracteres. Por exemplo, a função personalizada inserida na interface do usuário do Excel `=CONTOSO.GETRANGEVALUE("A1")`deve seguir o padrão, `"A1"` onde está o endereço da célula da qual recuperar o valor.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 const context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Limitações de chamar APIs JavaScript do Excel por meio de uma função personalizada

Um suplemento de funções personalizadas pode chamar APIs JavaScript do Excel, mas você deve ter cuidado com quais APIs ele chama. Não chame APIs JavaScript do Excel de uma função personalizada que altere células fora da célula que executa a função personalizada. Alterar outras células ou o ambiente do Excel pode resultar em baixo desempenho, tempos limite e loops infinitos no aplicativo excel. Isso significa que suas funções personalizadas não devem fazer o seguinte:

- Inserir, excluir ou formatar células na planilha.
- Altere o valor de outra célula.
- Mover, renomear, excluir ou adicionar planilhas a uma pasta de trabalho.
- Adicione nomes a uma pasta de trabalho.
- Definir propriedades.
- Altere qualquer uma das opções de ambiente do Excel, como modo de cálculo ou exibições de tela.

O suplemento de funções personalizadas pode ler informações de células fora da célula que executa a função personalizada, mas não deve executar operações de gravação em outras células. Em vez disso, faça alterações em outras células ou no ambiente do Excel a partir do contexto de um botão da faixa de opções ou de um painel de tarefas. Além disso, os cálculos de função personalizados não devem ser executados enquanto um recálculo do Excel está ocorrendo, pois esse cenário cria resultados imprevisíveis.

## <a name="next-steps"></a>Próximas etapas

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Confira também

- [Compartilhar dados e eventos entre funções personalizadas do Excel e o tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configurar seu Suplemento do Office para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
