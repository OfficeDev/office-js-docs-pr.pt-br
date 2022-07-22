---
title: Adiar a execução enquanto a célula está sendo editada
description: Saiba como atrasar a execução da função Excel.run quando uma célula está sendo editada.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c434fddf70c89d49712c96a42db772d67168a1fb
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958529"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Adiar a execução enquanto a célula está sendo editada

`Excel.run` tem uma sobrecarga que usa um [objeto Excel.RunOptions](/javascript/api/excel/excel.runoptions) . Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada. No momento, há suporte para a propriedade a seguir.

- `delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula. Quando `true`, a solicitação em lote é atrasada e é executada quando o usuário sai do modo de edição de célula. Quando `false`, a solicitação em lote falhará automaticamente se o usuário estiver no modo de edição de célula (causando um erro para alcançar o usuário). O comportamento padrão sem nenhuma `delayForCellEdit` propriedade especificada é equivalente a quando ele é `false`.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
