---
title: Tratamento de erros com a API JavaScript do Excel
description: Saiba mais sobre a lógica de tratamento de erros da API JavaScript do Excel para levar em conta os erros de tempo de execução.
ms.date: 01/06/2021
localization_priority: Normal
ms.openlocfilehash: fd863e9783336ba9121312ba06aae03330d57562
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789118"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Tratamento de erros com a API JavaScript do Excel

Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.

> [!NOTE]
> Para obter mais informações sobre o método e a natureza assíncrona da API JavaScript do Excel, confira o modelo de objeto JavaScript do Excel nos `sync()` [Complementos do Office.](excel-add-ins-core-concepts.md)

## <a name="best-practices"></a>Práticas recomendadas

Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a>Erros de API

Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:

- **code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.

- **message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada. A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.

- **debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.

> [!NOTE]
> Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do complemento ou em qualquer lugar no aplicativo do Office.

## <a name="error-messages"></a>Mensagens de erro

A tabela a seguir é uma lista de erros que a API pode retornar.

|Código de erro | Mensagem de erro |
|:----------|:--------------|
|`AccessDenied` |Você não pode realizar a operação solicitada.|
|`ActivityLimitReached`|O limite de atividades foi alcançado.|
|`ApiNotAvailable`|A API solicitada não está disponível.|
|`ApiNotFound`|Não foi possível encontrar a API que você está tentando usar. Ele pode estar disponível em uma versão mais recente do Excel. Confira o [artigo sobre conjuntos de requisitos](../reference/requirement-sets/excel-api-requirement-sets.md) da API JavaScript do Excel para saber mais.|
|`BadPassword`|A senha fornecida está incorreta.|
|`Conflict`|A solicitação não pôde ser processada devido a um conflito.|
|`ContentLengthRequired`|Um `Content-length` cabeçalho HTTP está ausente.|
|`GeneralException`|Ocorreu um erro interno ao processar a solicitação.|
|`InactiveWorkbook`|A operação falhou porque várias workbooks estão abertas e a área de trabalho chamada por essa API perdeu o foco.|
|`InsertDeleteConflict`|A tentativa de operação de exclusão ou inserção resultou em um conflito.|
|`InvalidArgument` |O argumento é inválido, está ausente ou tem um formato incorreto.|
|`InvalidBinding`  |Esta associação de objetos não é mais válida devido às atualizações anteriores.|
|`InvalidOperation`|A tentativa de operação é inválida no objeto.|
|`InvalidReference`|Esta referência não é válida para a operação atual.|
|`InvalidRequest`  |Não é possível processar a solicitação.|
|`InvalidSelection`|A seleção atual é inválida para esta operação.|
|`ItemAlreadyExists`|O recurso que está sendo criado já existe.|
|`ItemNotFound` |O recurso solicitado não existe.|
|`NonBlankCellOffSheet`|A solicitação para inserir novas células não pode ser concluída porque ela tiraria as células não vazias do final da planilha. Essas células não vazias podem aparecer vazias, mas têm valores em branco, alguma formatação ou uma fórmula. Exclua linhas ou colunas suficientes para dar espaço ao que você deseja inserir e tente novamente.|
|`NotImplemented`|O recurso solicitado não foi implementado.|
|`RangeExceedsLimit`|A contagem de células no intervalo excedeu o número máximo suportado. Confira o [artigo Limites de recursos e otimização](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) de desempenho para Os Complementos do Office para obter mais informações.|
|`RequestAborted`|A solicitação foi anulada durante o tempo de execução.|
|`RequestPayloadSizeLimitExceeded`|O tamanho da carga da solicitação excedeu o limite. Confira o [artigo Limites de recursos e otimização](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) de desempenho para Os Complementos do Office para obter mais informações. <br><br>Esse erro só ocorre no Excel na Web.|
|`ResponsePayloadSizeLimitExceeded`|O tamanho da carga de resposta excedeu o limite. Confira o [artigo Limites de recursos e otimização](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) de desempenho para Os Complementos do Office para obter mais informações.  <br><br>Esse erro só ocorre no Excel na Web.|
|`ServiceNotAvailable`|O serviço não está disponível.|
|`Unauthenticated` |Informações de autenticação necessárias estão ausentes ou inválidas.|
|`UnsupportedOperation`|Não há suporte para a operação que está sendo tentada.|
|`UnsupportedSheet`|Esse tipo de planilha não dá suporte a essa operação, pois é uma planilha de Macro ou Gráfico.|

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
