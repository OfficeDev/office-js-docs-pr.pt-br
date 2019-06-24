---
ms.date: 06/17/2019
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis nas funções
localization_priority: Normal
ms.openlocfilehash: bcaef092ec386a7d80760c1e2a567b9de1fdad21
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127812"
---
# <a name="volatile-values-in-functions"></a>Valores voláteis nas funções

Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada. O valor pode ser alterado mesmo se nenhum argumento da função for alterado. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`. Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

As funções personalizadas permitem que você crie suas próprias funções voláteis, o que pode ser útil ao lidar com datas, horas, números aleatórios e modelagem. Por exemplo, as simulações do [Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) exigem a geração de entradas aleatórias para determinar uma solução ideal.

Se escolher gerar automaticamente o arquivo JSON, declare uma função volátil com a marca `@volatile`de comentário JSDOC. Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

Um exemplo de uma função personalizada volátil segue, que simula a transferência de um ou mais de seis lados.

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>Próximas etapas
Saiba como [salvar o estado em suas funções personalizadas](custom-functions-save-state.md).

## <a name="see-also"></a>Confira também

* [Opções de parâmetros de funções personalizadas](custom-functions-parameter-options.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
