---
title: Adiar a execução enquanto a célula está sendo editada
description: Saiba como atrasar a execução do método Excel.run quando uma célula está sendo editada.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1abcdb382150db486033b32d2521207ab0b7f28f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889216"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Adiar a execução enquanto a célula está sendo editada

`Excel.run` tem uma sobrecarga que usa um [objeto Excel.RunOptions](/javascript/api/excel/excel.runoptions) . Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada. No momento, há suporte para a propriedade a seguir.

- `delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula. Quando `true`, a solicitação em lote é atrasada e é executada quando o usuário sai do modo de edição de célula. Quando `false`, a solicitação em lote falhará automaticamente se o usuário estiver no modo de edição de célula (causando um erro para alcançar o usuário). O comportamento padrão sem nenhuma `delayForCellEdit` propriedade especificada é equivalente a quando ele é `false`.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
