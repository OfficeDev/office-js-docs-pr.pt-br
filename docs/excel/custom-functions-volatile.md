---
ms.date: 01/14/2020
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis nas funções
localization_priority: Normal
ms.openlocfilehash: a318c87cc5b5f45bf3b1f5fe1341b7008f5a3d2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609787"
---
# <a name="volatile-values-in-functions"></a>Valores voláteis nas funções

Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada. O valor pode ser alterado mesmo se nenhum argumento da função for alterado. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`. Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

As funções personalizadas permitem que você crie suas próprias funções voláteis, o que pode ser útil ao lidar com datas, horas, números aleatórios e modelagem. Por exemplo, as [simulações do Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) exigem a geração de entradas aleatórias para determinar uma solução ideal.

Se escolher gerar automaticamente o arquivo JSON, declare uma função volátil com a marca de comentário JSDoc `@volatile` . Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

Um exemplo de uma função personalizada volátil segue, que simula a transferência de um ou mais de seis lados.

![Um gif mostrando uma função personalizada, retornando um valor aleatório para simular a rolagem de um e seis lados](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>Próximas etapas
* Saiba mais sobre [as opções de parâmetro de funções personalizadas](custom-functions-parameter-options.md).

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
