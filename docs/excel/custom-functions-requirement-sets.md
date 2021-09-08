---
title: Conjuntos de requisitos de Funções Personalizadas
description: Detalhes sobre os conjuntos de requisitos de Funções Personalizadas para Excel API JavaScript.
ms.date: 09/14/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0860dd2d1b55376a85eadf04898d288d83b0205d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938511"
---
# <a name="custom-functions-requirement-sets"></a>Conjuntos de requisitos de Funções Personalizadas

As [Funções Personalizadas](custom-functions-overview.md) usam conjuntos de requisitos separados das principais APIs JavaScript do Excel. A tabela a seguir lista os conjuntos de requisitos de Funções Personalizadas, os aplicativos Office cliente com suporte e as versões de com build ou número desses aplicativos.

|  Conjunto de requisitos  |  Office no Windows<br>(conectado a uma assinatura do Microsoft 365)  |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.13127.20296 ou posterior | Sem suporte | 16.40.20081000 ou posterior | Julho de 2020 |
| CustomFunctionsRuntime 1.2 | 16.0.12527.20194 ou posterior | Sem suporte | 16.34.20020900 ou posterior | Janeiro de 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 ou posterior | Sem suporte | 16.34 ou posterior | Maio de 2019 |

> [!NOTE]
> Excel funções personalizadas não são suportadas no Office 2019 ou anterior (compra única).

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1, 1.2 e 1.3

CustomFunctionsRuntime 1.1 é a primeira versão da API. O conjunto de requisitos 1.2 adiciona o `CustomFunctions.Error` objeto para dar suporte ao tratamento de erros. O conjunto de requisitos 1.3 adiciona suporte [a streaming XLL](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) e novas opções ao `ErrorCode` objeto [CustomFunctions.Error.](/javascript/api/custom-functions-runtime/customfunctions.error) 

## <a name="see-also"></a>Confira também

- [Documentação de referência de funções personalizadas](/javascript/api/custom-functions-runtime)
- [Conjuntos de requisitos da API JavaScript do Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
