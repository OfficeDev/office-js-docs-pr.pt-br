---
ms.date: 04/30/2019
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis em funções (visualização)
localization_priority: Normal
ms.openlocfilehash: 63618adecff57398e1630e6b5ab43c0dbc753b36
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527295"
---
## <a name="volatile-values-in-functions"></a><span data-ttu-id="9a8d0-103">Valores voláteis em funções</span><span class="sxs-lookup"><span data-stu-id="9a8d0-103">Volatile values in functions</span></span>

<span data-ttu-id="9a8d0-104">Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="9a8d0-105">O valor pode ser alterado mesmo se nenhum argumento da função for alterado.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="9a8d0-106">Essas funções são recalculadas sempre que o Excel recalcular.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="9a8d0-107">Por exemplo, imagine uma célula que chame a função `NOW`.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="9a8d0-108">Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="9a8d0-109">O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="9a8d0-110">Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="9a8d0-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="9a8d0-111">As funções personalizadas permitem que você crie suas próprias funções voláteis, o que pode ser útil ao lidar com datas, horas, números aleatórios e modelagem.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="9a8d0-112">Por exemplo, as simulações do [Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method
) exigem a geração de entradas aleatórias para determinar uma solução ideal.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="9a8d0-113">Se escolher gerar automaticamente o arquivo JSON, declare uma função volátil com a marca `@volatile`de comentário JSDOC.</span><span class="sxs-lookup"><span data-stu-id="9a8d0-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="9a8d0-114">Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="9a8d0-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9a8d0-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="9a8d0-115">See also</span></span>

* [<span data-ttu-id="9a8d0-116">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="9a8d0-116">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="9a8d0-117">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="9a8d0-117">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="9a8d0-118">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="9a8d0-118">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="9a8d0-119">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="9a8d0-119">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="9a8d0-120">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="9a8d0-120">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
