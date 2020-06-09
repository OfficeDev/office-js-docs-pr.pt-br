---
ms.date: 05/11/2020
description: Retornar vários resultados de sua função personalizada em um suplemento do Office Excel.
title: Retornar vários resultados de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: e25965277fbbe1c39007f79f401bf62b25760488
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609647"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="4afa6-103">Retornar vários resultados de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="4afa6-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="4afa6-104">Você pode retornar vários resultados de sua função personalizada que serão retornadas às células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="4afa6-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="4afa6-105">Esse comportamento é chamado de despejo.</span><span class="sxs-lookup"><span data-stu-id="4afa6-105">This behavior is called spilling.</span></span> <span data-ttu-id="4afa6-106">Quando sua função personalizada retorna uma matriz de resultados, ela é conhecida como uma fórmula de matriz dinâmica.</span><span class="sxs-lookup"><span data-stu-id="4afa6-106">When your custom function returns an array of results, it's known as a dynamic array formula.</span></span> <span data-ttu-id="4afa6-107">Para obter mais informações sobre fórmulas de matriz dinâmicas no Excel, consulte [matrizes dinâmicas e comportamento de matriz despejada](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span><span class="sxs-lookup"><span data-stu-id="4afa6-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="4afa6-108">A imagem a seguir mostra como a `SORT` função é despejada para baixo nas células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="4afa6-108">The following image shows how the `SORT` function spills down into neighboring cells.</span></span> <span data-ttu-id="4afa6-109">Sua função personalizada também pode retornar vários resultados como este.</span><span class="sxs-lookup"><span data-stu-id="4afa6-109">Your custom function can also return multiple results like this.</span></span>

![Captura de tela da função "SORT" exibindo vários resultados em várias células.](../images/dynamic-array-spill.png)

<span data-ttu-id="4afa6-111">Para criar uma função personalizada que seja uma fórmula de matriz dinâmica, ela deve retornar uma matriz bidimensional de valores.</span><span class="sxs-lookup"><span data-stu-id="4afa6-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="4afa6-112">Se os resultados forem despejados em células vizinhas que já possuem valores, a fórmula exibirá um `#SPILL!` erro.</span><span class="sxs-lookup"><span data-stu-id="4afa6-112">If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.</span></span>

<span data-ttu-id="4afa6-113">O exemplo a seguir mostra como retornar uma matriz dinâmica que derrama.</span><span class="sxs-lookup"><span data-stu-id="4afa6-113">The following example shows how to return a dynamic array that spills down.</span></span>

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

<span data-ttu-id="4afa6-114">O exemplo a seguir mostra como retornar uma matriz dinâmica que despeja à direita.</span><span class="sxs-lookup"><span data-stu-id="4afa6-114">The following example shows how to return a dynamic array that spills right.</span></span> 

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

<span data-ttu-id="4afa6-115">O exemplo a seguir mostra como retornar uma matriz dinâmica que é despejada para baixo e para a direita.</span><span class="sxs-lookup"><span data-stu-id="4afa6-115">The following example shows how to return a dynamic array that spills both down and right.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4afa6-116">Confira também</span><span class="sxs-lookup"><span data-stu-id="4afa6-116">See also</span></span>

- [<span data-ttu-id="4afa6-117">Matrizes dinâmicas e comportamento de matriz derramada</span><span class="sxs-lookup"><span data-stu-id="4afa6-117">Dynamic arrays and spilled array behavior</span></span>](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="4afa6-118">Opções para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="4afa6-118">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)