---
ms.date: 01/14/2020
description: Aprenda a implementar funções personalizadas de streaming voláteis e offline.
title: Valores voláteis nas funções
ms.localizationpriority: medium
ms.openlocfilehash: 401be3e04a7b36a226547175df4311fc653c027a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744475"
---
# <a name="volatile-values-in-functions"></a>Valores voláteis nas funções

Funções voláteis são funções nas quais o valor muda cada vez que a célula é calculada. O valor pode mudar mesmo que nenhum dos argumentos da função mude. Essas funções são recalculadas sempre que o Excel recalcular. Por exemplo, imagine uma célula que chame a função `NOW`. Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`. Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Funções personalizadas permitem que você crie suas próprias funções voláteis, que podem ser úteis ao lidar com datas, horas, números aleatórios e modelagem. Por exemplo, [as simulações de Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) exigem a geração de entradas aleatórias para determinar uma solução ideal.

Se optar por gerar automaticamente seu arquivo JSON, declare uma função volátil com a marca de comentário `@volatile`JSDoc . Para obter mais informações sobre a geração automática, consulte [Metadados JSON de geração automática para funções personalizadas](custom-functions-json-autogeneration.md).

Segue-se um exemplo de uma função personalizada volátil, que simula a rolagem de um dado de seis lados.

![GIF mostrando uma função personalizada retornando um valor aleatório para simular a rolagem de um dado de seis lados.](../images/six-sided-die.gif)

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
* Saiba mais [sobre as opções de parâmetros de funções personalizadas](custom-functions-parameter-options.md).

## <a name="see-also"></a>Confira também

* [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
