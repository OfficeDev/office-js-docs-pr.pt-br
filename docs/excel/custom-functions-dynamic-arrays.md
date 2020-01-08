---
ms.date: 12/18/2019
description: Retornar vários resultados de sua função personalizada em um suplemento do Office Excel.
title: Retornar vários resultados de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: 687ffcd66cff16d92fec372a778fe94bad7b38d5
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2020
ms.locfileid: "40970367"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="94e48-103">Retornar vários resultados de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="94e48-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="94e48-104">Você pode retornar vários resultados de sua função personalizada que serão retornadas às células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="94e48-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="94e48-105">Esse comportamento é chamado de despejo.</span><span class="sxs-lookup"><span data-stu-id="94e48-105">This behavior is called spilling.</span></span> <span data-ttu-id="94e48-106">Quando sua função personalizada retorna uma matriz de resultados, ela é conhecida como uma fórmula de matriz dinâmica.</span><span class="sxs-lookup"><span data-stu-id="94e48-106">When your custom function returns an array of results, it is known as a dynamic array formula.</span></span> <span data-ttu-id="94e48-107">Para obter mais informações sobre fórmulas de matriz dinâmicas no Excel, consulte [matrizes dinâmicas e comportamento de matriz despejada](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span><span class="sxs-lookup"><span data-stu-id="94e48-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="94e48-108">A imagem a seguir mostra como a função **Sort** é despejada nas células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="94e48-108">The following image shows how the **SORT** function spills down into neighboring cells.</span></span> <span data-ttu-id="94e48-109">Sua função personalizada também pode retornar vários resultados como este.</span><span class="sxs-lookup"><span data-stu-id="94e48-109">Your custom function can also return multiple results like this.</span></span>

![Captura de tela da função SORT que exibe vários resultados em várias células.](../images/dynamic-array-spill.png)

<span data-ttu-id="94e48-111">Para criar uma função personalizada que seja uma fórmula de matriz dinâmica, ela deve retornar uma matriz bidimensional de valores.</span><span class="sxs-lookup"><span data-stu-id="94e48-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="94e48-112">Se os resultados forem despejados em células vizinhas que já possuem valores, a fórmula exibirá um **#SPILL!**</span><span class="sxs-lookup"><span data-stu-id="94e48-112">If the results spill into neighboring cells that already have values, the formula will display a **#SPILL!**</span></span> <span data-ttu-id="94e48-113">.</span><span class="sxs-lookup"><span data-stu-id="94e48-113">error.</span></span> 

<span data-ttu-id="94e48-114">O exemplo a seguir mostra como retornar uma matriz dinâmica que derrama.</span><span class="sxs-lookup"><span data-stu-id="94e48-114">The following example shows how to return a dynamic array that spills down.</span></span>

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

<span data-ttu-id="94e48-115">O exemplo a seguir mostra como retornar uma matriz dinâmica que despeja à direita.</span><span class="sxs-lookup"><span data-stu-id="94e48-115">The following example shows how to return a dynamic array that spills right.</span></span> 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

<span data-ttu-id="94e48-116">O exemplo a seguir mostra como retornar uma matriz dinâmica que é despejada para baixo e para a direita.</span><span class="sxs-lookup"><span data-stu-id="94e48-116">The following example shows how to return a dynamic array that spills both down and right.</span></span>

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a><span data-ttu-id="94e48-117">Confira também</span><span class="sxs-lookup"><span data-stu-id="94e48-117">See also</span></span>

- [<span data-ttu-id="94e48-118">Matrizes dinâmicas e comportamento de matriz derramada</span><span class="sxs-lookup"><span data-stu-id="94e48-118">Dynamic arrays and spilled array behavior</span></span>](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="94e48-119">Opções para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="94e48-119">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)