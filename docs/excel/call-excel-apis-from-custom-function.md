---
title: Chamar as APIs do Microsoft Excel a partir de uma função personalizada
description: Saiba quais APIs do Microsoft Excel você pode chamar a partir de sua função personalizada.
ms.date: 02/06/2020
localization_priority: Normal
ms.openlocfilehash: e22ed897e95a74707bd0d8bded3f8dca724731d1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719340"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a>Chamar as APIs do Microsoft Excel a partir de uma função personalizada

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

Chamar as APIs do Excel do Office. js de suas funções personalizadas para obter dados de intervalo e obter mais contexto para seus cálculos.

Chamar APIs do Office. js por meio de uma função personalizada pode ser útil quando:

- Uma função personalizada precisa obter informações do Excel antes do cálculo. Essas informações podem incluir propriedades de documento, formatos de intervalo, partes XML personalizadas, um nome de pasta de trabalho ou outras informações específicas do Excel.
- Uma função personalizada definirá o formato de número da célula para os valores de retorno após o cálculo.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="code-sample"></a>Exemplo de código

Para chamar as APIs do Office. js, você precisa primeiro de um contexto. Use o `Excel.RequestContext` objeto para obter um contexto. Em seguida, use o contexto para chamar as APIs de que você precisa na pasta de trabalho.

O exemplo de código a seguir mostra como obter um intervalo de valores da pasta de trabalho.

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a>Limitações da chamada do Office. js por meio de uma função personalizada

Não chame APIs do Office. js de uma função personalizada que altere o ambiente do Excel. Isso significa que suas funções personalizadas não devem fazer o seguinte:

- Inserir, excluir ou Formatar células na planilha.
- Altera o valor de outra célula.
- Mover, renomear, excluir ou adicionar planilhas a uma pasta de trabalho.
- Alterar qualquer uma das opções de ambiente, como modo de cálculo ou modos de exibição de tela.
- Adicionar nomes a uma pasta de trabalho.
- Definir propriedades ou executar a maioria dos métodos.

Alterar o Excel pode resultar em desempenho ruim, tempo limite e loops infinitos. Cálculos de função personalizada não devem ser executados enquanto um recálculo do Excel está ocorrendo, pois resultará em resultados imprevisíveis.

Em vez disso, faça alterações no Excel a partir do contexto de um botão da faixa de opções ou de um painel de tarefas.

## <a name="next-steps"></a>Próximas etapas

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Também confira

- [Compartilhar dados e eventos entre as funções personalizadas do Excel e o tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)