---
title: Chamar Excel APIs JavaScript de uma função personalizada
description: Saiba quais Excel APIs JavaScript que você pode chamar de sua função personalizada.
ms.date: 08/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7b60f3fbdeb317169800c688b77982580dfbf8c4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744392"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Chamar Excel APIs JavaScript de uma função personalizada

Chame Excel APIs JavaScript de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos. Chamar Excel APIs JavaScript por meio de uma função personalizada pode ser útil quando:

- Uma função personalizada precisa obter informações de Excel antes do cálculo. Essas informações podem incluir propriedades de documento, formatos de intervalo, partes XML personalizadas, um nome da Excel de trabalho ou outras informações específicas.
- Uma função personalizada definirá o formato de número da célula para os valores de retorno após o cálculo.

> [!IMPORTANT]
> Para chamar Excel APIs JavaScript de sua função personalizada, você precisará usar um tempo de execução JavaScript compartilhado. Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

## <a name="code-sample"></a>Exemplo de código

Para chamar Excel APIs JavaScript de uma função personalizada, primeiro você precisa de um contexto. Use o [Excel. Objeto RequestContext](/javascript/api/excel/excel.requestcontext) para obter um contexto. Em seguida, use o contexto para chamar as APIs de que você precisa na guia de trabalho.

O exemplo de código a seguir mostra como usar `Excel.RequestContext` para obter um valor de uma célula na lista de trabalho. Neste exemplo, o `address` parâmetro é passado para o método Excel API JavaScript [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) e deve ser inserido como uma cadeia de caracteres. Por exemplo, a função `=CONTOSO.GETRANGEVALUE("A1")`personalizada inserida na interface do usuário Excel deve seguir o padrão , `"A1"` onde é o endereço da célula da qual recuperar o valor.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Limitações de chamar Excel APIs JavaScript por meio de uma função personalizada

Não chame Excel APIs JavaScript de uma função personalizada que altere o ambiente de Excel. Isso significa que suas funções personalizadas não devem fazer nada do seguinte:

- Inserir, excluir ou formatar células na planilha.
- Altere o valor de outra célula.
- Mover, renomear, excluir ou adicionar planilhas a uma planilha.
- Altere qualquer uma das opções de ambiente, como modo de cálculo ou modos de exibição de tela.
- Adicione nomes a uma lista de trabalho.
- Definir propriedades ou executar a maioria dos métodos.

Alterar Excel pode resultar em desempenho ruim, tempo de insufinições e loops infinitos. Os cálculos de função personalizados não devem ser executados enquanto um recálculo Excel está ocorrendo, pois resultará em resultados imprevisíveis.

Em vez disso, faça alterações Excel do contexto de um botão de faixa de opções ou painel de tarefas.

## <a name="next-steps"></a>Próximas etapas

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Confira também

- [Compartilhar dados e eventos entre Excel funções personalizadas e tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
