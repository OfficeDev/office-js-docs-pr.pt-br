---
title: Chamar APIs JavaScript do Excel de uma função personalizada
description: Saiba quais APIs JavaScript do Excel você pode chamar de sua função personalizada.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613903"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Chamar APIs JavaScript do Excel de uma função personalizada

Chame APIs JavaScript do Excel de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos. Chamar APIs JavaScript do Excel por meio de uma função personalizada pode ser útil quando:

- Uma função personalizada precisa obter informações do Excel antes do cálculo. Essas informações podem incluir propriedades de documento, formatos de intervalo, partes XML personalizadas, um nome da planilha ou outras informações específicas do Excel.
- Uma função personalizada definirá o formato de número da célula para os valores de retorno após o cálculo.

> [!IMPORTANT]
> Para chamar APIs JavaScript do Excel de sua função personalizada, você precisará usar um tempo de execução JavaScript compartilhado. Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

## <a name="code-sample"></a>Exemplo de código

Para chamar APIs JavaScript do Excel de uma função personalizada, primeiro você precisa de um contexto. Use o [objeto Excel.RequestContext](/javascript/api/excel/excel.requestcontext) para obter um contexto. Em seguida, use o contexto para chamar as APIs de que você precisa na guia de trabalho.

O exemplo de código a seguir mostra como usar para obter um `Excel.RequestContext` valor de uma célula na lista de trabalho. Neste exemplo, o parâmetro é passado para o `address` método [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) da API JavaScript do Excel e deve ser inserido como uma cadeia de caracteres. Por exemplo, a função personalizada inserida na interface do usuário do Excel deve seguir o padrão , onde é o endereço da célula da qual `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` recuperar o valor.

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
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Limitações de chamar APIs JavaScript do Excel por meio de uma função personalizada

Não chame APIs JavaScript do Excel de uma função personalizada que altere o ambiente do Excel. Isso significa que suas funções personalizadas não devem fazer nada do seguinte:

- Inserir, excluir ou formatar células na planilha.
- Altere o valor de outra célula.
- Mover, renomear, excluir ou adicionar planilhas a uma planilha.
- Altere qualquer uma das opções de ambiente, como modo de cálculo ou modos de exibição de tela.
- Adicione nomes a uma lista de trabalho.
- Definir propriedades ou executar a maioria dos métodos.

Alterar o Excel pode resultar em desempenho ruim, tempo de insufinições e loops infinitos. Os cálculos de função personalizada não devem ser executados enquanto um recálculo do Excel está ocorrendo, pois resultará em resultados imprevisíveis.

Em vez disso, faça alterações no Excel a partir do contexto de um botão de faixa de opções ou do painel de tarefas.

## <a name="next-steps"></a>Próximas etapas

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Confira também

- [Compartilhar dados e eventos entre funções personalizadas do Excel e tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
