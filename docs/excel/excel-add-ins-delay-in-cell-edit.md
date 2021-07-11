---
title: Adiar a execução enquanto a célula está sendo editada
description: Saiba como atrasar a execução do método Excel.run quando uma célula está sendo editada.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: b7b28064ef4d313639391e63cba780351b5623f9
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349515"
---
# <a name="delay-execution-while-cell-is-being-edited"></a><span data-ttu-id="2ae69-103">Adiar a execução enquanto a célula está sendo editada</span><span class="sxs-lookup"><span data-stu-id="2ae69-103">Delay execution while cell is being edited</span></span>

<span data-ttu-id="2ae69-104">`Excel.run`tem uma sobrecarga que leva em um [Excel. Objeto RunOptions.](/javascript/api/excel/excel.runoptions)</span><span class="sxs-lookup"><span data-stu-id="2ae69-104">`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="2ae69-105">Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada.</span><span class="sxs-lookup"><span data-stu-id="2ae69-105">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="2ae69-106">A propriedade a seguir tem suporte no momento.</span><span class="sxs-lookup"><span data-stu-id="2ae69-106">The following property is currently supported.</span></span>

- <span data-ttu-id="2ae69-107">`delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula.</span><span class="sxs-lookup"><span data-stu-id="2ae69-107">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="2ae69-108">Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula.</span><span class="sxs-lookup"><span data-stu-id="2ae69-108">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="2ae69-109">Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário).</span><span class="sxs-lookup"><span data-stu-id="2ae69-109">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="2ae69-110">O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.</span><span class="sxs-lookup"><span data-stu-id="2ae69-110">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
