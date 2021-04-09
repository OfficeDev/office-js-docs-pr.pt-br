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
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a>Ler ou gravar em um intervalo grande usando a API JavaScript do Excel

Este artigo descreve como lidar com a leitura e a escrita em intervalos grandes com a API JavaScript do Excel.

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a>Executar operações de leitura ou gravação separadas para intervalos grandes

Se um intervalo contiver um grande número de células, valores, formatos de número ou fórmulas, talvez não seja possível executar operações de API nesse intervalo. A API sempre fará a melhor tentativa de executar a operação solicitada em um intervalo (isto é, para recuperar ou gravar os dados especificados), mas tentar executar operações de leitura ou gravação para um intervalo grande pode resultar em um erro de API devido à utilização excessiva de recursos. Para evitar tais erros, é recomendável executar operações de leitura ou gravação separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma única operação de leitura ou gravação em um intervalo grande.

Para obter detalhes sobre as limitações do sistema, consulte a seção "Complementos do Excel" de Limites de recursos e otimização de desempenho para [Os Complementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).

### <a name="conditional-formatting-of-ranges"></a>Formatação condicional de intervalos

Os intervalos podem ter formatos aplicados a células individuais baseadas em condições. Confira mais informações sobre isso em [Aplicar a formatação condicional a intervalos do Excel](excel-add-ins-conditional-formatting.md).

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Ler ou gravar em um intervalo não-rebote usando a API JavaScript do Excel](excel-add-ins-ranges-unbounded.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
