---
ms.date: 01/14/2020
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis nas funções
localization_priority: Normal
ms.openlocfilehash: 0f530e9d67894ebbc13c8b8a13e6219571c96ff1
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071627"
---
# <a name="volatile-values-in-functions"></a>Valores voláteis nas funções

Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada. O valor pode ser alterado mesmo se nenhum argumento da função for alterado. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`. Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

As funções personalizadas permitem que você crie suas próprias funções voláteis, o que pode ser útil ao lidar com datas, horas, números aleatórios e modelagem. Por exemplo, as [simulações do Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) exigem a geração de entradas aleatórias para determinar uma solução ideal.

Se escolher gerar automaticamente o arquivo JSON, declare uma função volátil com a marca de comentário JSDoc `@volatile` . Para obter mais informações sobre a autogeração, consulte [AutoGenerate metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).

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

* [Criar manualmente metadados JSON para funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
