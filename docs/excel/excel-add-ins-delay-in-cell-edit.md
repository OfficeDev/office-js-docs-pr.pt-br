---
title: Atrasar a execução durante a edição da célula
description: Saiba como atrasar a execução do método Excel. Run quando uma célula estiver sendo editada.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: eb33f4cb7cce3b1f8642e00f432e708e90b5b895
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409373"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Atrasar a execução durante a edição da célula

`Excel.run` tem uma sobrecarga que utiliza um objeto [Excel. RunOptions](/javascript/api/excel/excel.runoptions) . Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada. A propriedade a seguir tem suporte no momento:

* `delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula. Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula. Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário). O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
