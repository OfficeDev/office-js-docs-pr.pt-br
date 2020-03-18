---
ms.date: 12/18/2019
description: Retornar vários resultados de sua função personalizada em um suplemento do Office Excel.
title: Retornar vários resultados de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: a2632c621071f0cbc55f545847d9e9392d884b90
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719291"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="b0b20-103">Retornar vários resultados de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="b0b20-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="b0b20-104">Você pode retornar vários resultados de sua função personalizada que serão retornadas às células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="b0b20-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="b0b20-105">Esse comportamento é chamado de despejo.</span><span class="sxs-lookup"><span data-stu-id="b0b20-105">This behavior is called spilling.</span></span> <span data-ttu-id="b0b20-106">Quando sua função personalizada retorna uma matriz de resultados, ela é conhecida como uma fórmula de matriz dinâmica.</span><span class="sxs-lookup"><span data-stu-id="b0b20-106">When your custom function returns an array of results, it is known as a dynamic array formula.</span></span> <span data-ttu-id="b0b20-107">Para obter mais informações sobre fórmulas de matriz dinâmicas no Excel, consulte [matrizes dinâmicas e comportamento de matriz despejada](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span><span class="sxs-lookup"><span data-stu-id="b0b20-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="b0b20-108">A imagem a seguir mostra como `SORT` a função é despejada para baixo nas células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="b0b20-108">The following image shows how the `SORT` function spills down into neighboring cells.</span></span> <span data-ttu-id="b0b20-109">Sua função personalizada também pode retornar vários resultados como este.</span><span class="sxs-lookup"><span data-stu-id="b0b20-109">Your custom function can also return multiple results like this.</span></span>

![Captura de tela da função "SORT" exibindo vários resultados em várias células.](../images/dynamic-array-spill.png)

<span data-ttu-id="b0b20-111">Para criar uma função personalizada que seja uma fórmula de matriz dinâmica, ela deve retornar uma matriz bidimensional de valores.</span><span class="sxs-lookup"><span data-stu-id="b0b20-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="b0b20-112">Se os resultados forem despejados em células vizinhas que já possuem valores, a fórmula exibirá um `#SPILL!` erro.</span><span class="sxs-lookup"><span data-stu-id="b0b20-112">If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.</span></span>

<span data-ttu-id="b0b20-113">O exemplo a seguir mostra como retornar uma matriz dinâmica que derrama.</span><span class="sxs-lookup"><span data-stu-id="b0b20-113">The following example shows how to return a dynamic array that spills down.</span></span>

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

<span data-ttu-id="b0b20-114">O exemplo a seguir mostra como retornar uma matriz dinâmica que despeja à direita.</span><span class="sxs-lookup"><span data-stu-id="b0b20-114">The following example shows how to return a dynamic array that spills right.</span></span> 

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

<span data-ttu-id="b0b20-115">O exemplo a seguir mostra como retornar uma matriz dinâmica que é despejada para baixo e para a direita.</span><span class="sxs-lookup"><span data-stu-id="b0b20-115">The following example shows how to return a dynamic array that spills both down and right.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="b0b20-116">Também confira</span><span class="sxs-lookup"><span data-stu-id="b0b20-116">See also</span></span>

- [<span data-ttu-id="b0b20-117">Matrizes dinâmicas e comportamento de matriz derramada</span><span class="sxs-lookup"><span data-stu-id="b0b20-117">Dynamic arrays and spilled array behavior</span></span>](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="b0b20-118">Opções para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="b0b20-118">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)