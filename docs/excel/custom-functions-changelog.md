---
ms.date: 01/08/2019
description: Descubra as atualizações mais recentes para as funções personalizadas do Excel.
title: Log de alteração de funções personalizadas (visualização)
localization_priority: Normal
ms.openlocfilehash: 03e4dd922ac3895e11a508f97e7ac3fa3e7b1cb0
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477534"
---
# <a name="custom-functions-changelog-preview"></a>Log de alteração de funções personalizadas (visualização)

As funções personalizadas do Excel ainda estão em visualização e isso significa que há alterações frequentes para o produto, incluindo alterações e o lançamento de novos recursos. Esse Log de alteração oferece informações mais atualizadas sobre as alterações para o produto.

- **7 de novembro de 2017**: enviados exemplos e visualizações de funções personalizadas
- **20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores
- **28 de novembro de 2017**: enviado o suporte para cancelamento em funções assíncronas (requer a alteração de funções de streaming)
- **7 de maio de 2018**: Suporte enviado para Mac, Excel Online e funções síncronas em execução no processo
- **20 de setembro de 2018**: Suporte enviado para funções personalizadas de tempo de execução do JavaScript. Para saber mais, veja o [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md).
- **20 de outubro de 2018**: Com o [build do Insider de outubro](https://support.office.com/en-us/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), as funções personalizadas agora exigem o parâmetro "id" na suas [funções personalizadas metadados](custom-functions-json.md) para a área de trabalho do Windows e Online. No Mac, esse parâmetro deve ser ignorado. Funções personalizadas também suportam parâmetros opcionais e do `any` tipo retorno.
- **12 de dezembro de 2018**: As funções personalizadas agora incluem uma maneira de descobrir o endereço da célula. Para saber mais, confira [determinar quais célula chamada sua função personalizada](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function).
- **8 de janeiro de 2019**: método de associação `CustomFunctionMapping()` foi alterado para `CustomFunctions.associate()`. Para saber mais, confira [práticas recomendadas de funções personalizados (visualização)](custom-functions-best-practices.md).

No \* canal[Office Insider](https://products.office.com/office-insider), (anteriormente chamado de "Insider – modo rápido")

Para obter uma lista de problemas conhecidos com o produto, confira [problemas conhecidos](custom-functions-overview.md#known-issues). 

## <a name="see-also"></a>Confira também

* [Visão geral de funções personalizadas](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas.](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Depuração de funções personalizadas](custom-functions-debugging.md)