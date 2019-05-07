---
ms.date: 05/03/2019
description: Saiba como implementar funções personalizadas de streaming volátil e offline.
title: Valores voláteis em funções
localization_priority: Normal
ms.openlocfilehash: 1ca3edc3de2d9ac5f2171004f89466352c5cfa1e
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627994"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="2ee4c-103">Valores voláteis em funções</span><span class="sxs-lookup"><span data-stu-id="2ee4c-103">Volatile values in functions</span></span>

<span data-ttu-id="2ee4c-104">Funções voláteis são funções nas quais o valor muda sempre que a célula é calculada.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="2ee4c-105">O valor pode ser alterado mesmo se nenhum argumento da função for alterado.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="2ee4c-106">Essas funções são recalculadas sempre que o Excel recalcular.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="2ee4c-107">Por exemplo, imagine uma célula que chame a função `NOW`.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="2ee4c-108">Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="2ee4c-109">O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="2ee4c-110">Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="2ee4c-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="2ee4c-111">As funções personalizadas permitem que você crie suas próprias funções voláteis, o que pode ser útil ao lidar com datas, horas, números aleatórios e modelagem.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="2ee4c-112">Por exemplo, as simulações do [Monte Carlo](https://en.wikipedia.org/wiki/Monte_Carlo_method
) exigem a geração de entradas aleatórias para determinar uma solução ideal.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="2ee4c-113">Se escolher gerar automaticamente o arquivo JSON, declare uma função volátil com a marca `@volatile`de comentário JSDOC.</span><span class="sxs-lookup"><span data-stu-id="2ee4c-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="2ee4c-114">Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="2ee4c-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="2ee4c-115">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="2ee4c-115">Next steps</span></span>
<span data-ttu-id="2ee4c-116">Saiba como [salvar o estado em suas funções personalizadas](custom-functions-save-state.md).</span><span class="sxs-lookup"><span data-stu-id="2ee4c-116">Learn how to [save state in your custom functions](custom-functions-save-state.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2ee4c-117">Confira também</span><span class="sxs-lookup"><span data-stu-id="2ee4c-117">See also</span></span>

* [<span data-ttu-id="2ee4c-118">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="2ee4c-118">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="2ee4c-119">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="2ee4c-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2ee4c-120">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="2ee4c-120">Create custom functions in Excel</span></span>](custom-functions-overview.md)
