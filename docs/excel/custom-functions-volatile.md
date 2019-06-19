---
ms.date: 06/17/2019
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis nas funções
localization_priority: Normal
ms.openlocfilehash: 0edf4071ce366c40300663233f1de318a544169b
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059843"
---
# <a name="volatile-values-in-functions"></a>Valores voláteis nas funções

Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada. O valor pode ser alterado mesmo se nenhum argumento da função for alterado. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

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
