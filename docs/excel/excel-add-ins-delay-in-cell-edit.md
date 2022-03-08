---
title: Adiar a execução enquanto a célula está sendo editada
description: Saiba como atrasar a execução do método Excel.run quando uma célula está sendo editada.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c5609fbb2a39d6ecc69063d4bccdfbc1da1c102d
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340803"
---
# <a name="delay-execution-while-cell-is-being-edited"></a>Adiar a execução enquanto a célula está sendo editada

`Excel.run`tem uma sobrecarga que leva em um [Excel. Objeto RunOptions](/javascript/api/excel/excel.runoptions). Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada. A propriedade a seguir tem suporte no momento.

- `delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula. Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula. Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário). O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.

```js
await Excel.run({ delayForCellEdit: true }, async (context) => { ... });
```
