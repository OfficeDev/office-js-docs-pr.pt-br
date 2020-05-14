---
title: Chamar as APIs do Microsoft Excel a partir de uma função personalizada
description: Saiba quais APIs do Microsoft Excel você pode chamar a partir de sua função personalizada.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a24cdfba2d79b6e2ad165765d22cd77743047d34
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217876"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a>Chamar as APIs do Microsoft Excel a partir de uma função personalizada

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

## <a name="see-also"></a>Confira também

- [Compartilhar dados e eventos entre as funções personalizadas do Excel e o tutorial do painel de tarefas](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)