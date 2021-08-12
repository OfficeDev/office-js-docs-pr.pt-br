---
title: Adiar a execução enquanto a célula está sendo editada
description: Saiba como atrasar a execução do método Excel.run quando uma célula está sendo editada.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 41bbfba3894bcef0c1fd075ce76557dfdc4ba4721b7bc7b19ca21756b86ccc4d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084278"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Adiar a execução enquanto a célula está sendo editada

`Excel.run`tem uma sobrecarga que leva em um [Excel. Objeto RunOptions.](/javascript/api/excel/excel.runoptions) Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada. A propriedade a seguir tem suporte no momento.

- `delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula. Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula. Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário). O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
