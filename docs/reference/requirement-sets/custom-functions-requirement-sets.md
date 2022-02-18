---
title: Conjuntos de requisitos de Funções Personalizadas
description: Detalhes sobre os conjuntos de requisitos de Funções Personalizadas para Excel API JavaScript.
ms.date: 02/15/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7558035b6b151977e985ec04ed1fa84c116f0886
ms.sourcegitcommit: 789545a81bd61ec2e7adef2bc24c06b5be113b00
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/18/2022
ms.locfileid: "62892514"
---
# <a name="custom-functions-requirement-sets"></a>Conjuntos de requisitos de Funções Personalizadas

As [Funções Personalizadas](../../excel/custom-functions-overview.md) usam conjuntos de requisitos separados das principais APIs JavaScript do Excel. A tabela a seguir lista os conjuntos de requisitos de Funções Personalizadas, os aplicativos Office cliente com suporte e as versões de com build ou número desses aplicativos.

|  Conjunto de requisitos  |  Office 2021 ou posterior no Windows<br>(compra avulsa)  |  Office no Windows<br>(conectado a uma assinatura do Microsoft 365)  |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2021 e posterior)  | Office na Web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.14326.20454 ou posterior | 16.0.13127.20296 ou posterior | Incompatível | 16.40.20081000 ou posterior | Julho de 2020 |
| CustomFunctionsRuntime 1.2 | 16.0.14326.20454 ou posterior | 16.0.12527.20194 ou posterior | Sem suporte | 16.34.20020900 ou posterior | Janeiro de 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.14326.20454 ou posterior | 16.0.12527.20092 ou posterior | Sem suporte | 16.34 ou posterior | Maio de 2019 |

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1, 1.2 e 1.3

CustomFunctionsRuntime 1.1 é a primeira versão da API. O conjunto de requisitos 1.2 adiciona o objeto para `CustomFunctions.Error` dar suporte ao tratamento de erros. O conjunto de requisitos 1.3 adiciona suporte [a streaming XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) e novas `ErrorCode` opções ao [objeto CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) .

## <a name="see-also"></a>Confira também

- [Documentação de referência de funções personalizadas](/javascript/api/custom-functions-runtime)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
