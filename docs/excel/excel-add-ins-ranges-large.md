---
title: Ler ou gravar em intervalos grandes usando a API JavaScript do Excel
description: Saiba como ler ou gravar em intervalos grandes com a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652767"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a><span data-ttu-id="7c6ea-103">Ler ou gravar em um intervalo grande usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7c6ea-103">Read or write to a large range using the Excel JavaScript API</span></span>

<span data-ttu-id="7c6ea-104">Este artigo descreve como lidar com a leitura e a escrita em intervalos grandes com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="7c6ea-104">This article describes how to handle reading and writing to large ranges with the Excel JavaScript API.</span></span>

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a><span data-ttu-id="7c6ea-105">Executar operações de leitura ou gravação separadas para intervalos grandes</span><span class="sxs-lookup"><span data-stu-id="7c6ea-105">Run separate read or write operations for large ranges</span></span>

<span data-ttu-id="7c6ea-106">Se um intervalo contiver um grande número de células, valores, formatos de número ou fórmulas, talvez não seja possível executar operações de API nesse intervalo.</span><span class="sxs-lookup"><span data-stu-id="7c6ea-106">If a range contains a large number of cells, values, number formats, or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="7c6ea-107">A API sempre fará a melhor tentativa de executar a operação solicitada em um intervalo (isto é, para recuperar ou gravar os dados especificados), mas tentar executar operações de leitura ou gravação para um intervalo grande pode resultar em um erro de API devido à utilização excessiva de recursos.</span><span class="sxs-lookup"><span data-stu-id="7c6ea-107">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="7c6ea-108">Para evitar tais erros, é recomendável executar operações de leitura ou gravação separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma única operação de leitura ou gravação em um intervalo grande.</span><span class="sxs-lookup"><span data-stu-id="7c6ea-108">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="7c6ea-109">Para obter detalhes sobre as limitações do sistema, consulte a seção "Complementos do Excel" de Limites de recursos e otimização de desempenho para [Os Complementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span><span class="sxs-lookup"><span data-stu-id="7c6ea-109">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="7c6ea-110">Formatação condicional de intervalos</span><span class="sxs-lookup"><span data-stu-id="7c6ea-110">Conditional formatting of ranges</span></span>

<span data-ttu-id="7c6ea-111">Os intervalos podem ter formatos aplicados a células individuais baseadas em condições.</span><span class="sxs-lookup"><span data-stu-id="7c6ea-111">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="7c6ea-112">Confira mais informações sobre isso em [Aplicar a formatação condicional a intervalos do Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="7c6ea-112">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7c6ea-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="7c6ea-113">See also</span></span>

- [<span data-ttu-id="7c6ea-114">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="7c6ea-114">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7c6ea-115">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7c6ea-115">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="7c6ea-116">Ler ou gravar em um intervalo não-rebote usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7c6ea-116">Read or write to an unbounded range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-unbounded.md)
- [<span data-ttu-id="7c6ea-117">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="7c6ea-117">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
