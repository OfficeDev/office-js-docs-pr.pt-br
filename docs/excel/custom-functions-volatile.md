---
ms.date: 01/14/2020
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis nas funções
localization_priority: Normal
ms.openlocfilehash: 0f530e9d67894ebbc13c8b8a13e6219571c96ff1
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071627"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="c6c44-103">Valores voláteis nas funções</span><span class="sxs-lookup"><span data-stu-id="c6c44-103">Volatile values in functions</span></span>

<span data-ttu-id="c6c44-104">Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada.</span><span class="sxs-lookup"><span data-stu-id="c6c44-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="c6c44-105">O valor pode ser alterado mesmo se nenhum argumento da função for alterado.</span><span class="sxs-lookup"><span data-stu-id="c6c44-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="c6c44-106">Essas funções são recalculadas sempre que o Excel recalcular.</span><span class="sxs-lookup"><span data-stu-id="c6c44-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="c6c44-107">Por exemplo, imagine uma célula que chame a função `NOW`.</span><span class="sxs-lookup"><span data-stu-id="c6c44-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="c6c44-108">Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.</span><span class="sxs-lookup"><span data-stu-id="c6c44-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="c6c44-109">O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="c6c44-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="c6c44-110">Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="c6c44-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="c6c44-111">As funções personalizadas permitem que você crie suas próprias funções voláteis, o que pode ser útil ao lidar com datas, horas, números aleatórios e modelagem.</span><span class="sxs-lookup"><span data-stu-id="c6c44-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="c6c44-112">Por exemplo, as [simulações do Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method) exigem a geração de entradas aleatórias para determinar uma solução ideal.</span><span class="sxs-lookup"><span data-stu-id="c6c44-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="c6c44-113">Se escolher gerar automaticamente o arquivo JSON, declare uma função volátil com a marca de comentário JSDoc `@volatile` .</span><span class="sxs-lookup"><span data-stu-id="c6c44-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="c6c44-114">Para obter mais informações sobre a autogeração, consulte [AutoGenerate metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="c6c44-114">From more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="c6c44-115">Um exemplo de uma função personalizada volátil segue, que simula a transferência de um ou mais de seis lados.</span><span class="sxs-lookup"><span data-stu-id="c6c44-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

![Um gif mostrando uma função personalizada, retornando um valor aleatório para simular a rolagem de um e seis lados](../images/six-sided-die.gif)

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

## <a name="next-steps"></a><span data-ttu-id="c6c44-117">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="c6c44-117">Next steps</span></span>
* <span data-ttu-id="c6c44-118">Saiba mais sobre [as opções de parâmetro de funções personalizadas](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="c6c44-118">Learn about [custom functions parameter options](custom-functions-parameter-options.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c6c44-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="c6c44-119">See also</span></span>

* [<span data-ttu-id="c6c44-120">Criar manualmente metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c6c44-120">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c6c44-121">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="c6c44-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
