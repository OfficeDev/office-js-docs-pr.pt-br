---
title: Tratamento de erros com a API JavaScript do Excel
description: Saiba mais sobre a lógica de tratamento de erro da API JavaScript do Excel para considerar os erros de tempo de execução.
ms.date: 10/22/2020
localization_priority: Normal
ms.openlocfilehash: a3b1bbfa7daba1b856bce35aa075d5b625bd9769
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740816"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Tratamento de erros com a API JavaScript do Excel

Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.

> [!NOTE]
> Para obter mais informações sobre o `sync()` método e a natureza assíncrona da API JavaScript do Excel, consulte [modelo de objeto do Excel JavaScript em suplementos do Office](excel-add-ins-core-concepts.md).

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
> Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo do Office.

## <a name="error-messages"></a>Mensagens de erro

A tabela a seguir é uma lista de erros que a API pode retornar.

|Código de erro | Mensagem de erro |
|:----------|:--------------|
|`AccessDenied` |Você não pode realizar a operação solicitada.|
|`ActivityLimitReached`|O limite de atividades foi alcançado.|
|`ApiNotAvailable`|A API solicitada não está disponível.|
|`ApiNotFound`|A API que você está tentando usar não pôde ser encontrada. Ele pode estar disponível em uma versão mais recente do Excel. Confira o artigo [conjuntos de requisitos da API JavaScript do Excel](../reference/requirement-sets/excel-api-requirement-sets.md) para obter mais informações.|
|`BadPassword`|A senha que você forneceu está incorreta.|
|`Conflict`|A solicitação não pôde ser processada devido a um conflito.|
|`ContentLengthRequired`|Um `Content-length` cabeçalho HTTP está ausente.|
|`GeneralException`|Ocorreu um erro interno ao processar a solicitação.|
|`InsertDeleteConflict`|A tentativa de operação de exclusão ou inserção resultou em um conflito.|
|`InvalidArgument` |O argumento é inválido, está ausente ou tem um formato incorreto.|
|`InvalidBinding`  |Esta associação de objetos não é mais válida devido às atualizações anteriores.|
|`InvalidOperation`|A tentativa de operação é inválida no objeto.|
|`InvalidReference`|Esta referência não é válida para a operação atual.|
|`InvalidRequest`  |Não é possível processar a solicitação.|
|`InvalidSelection`|A seleção atual é inválida para esta operação.|
|`ItemAlreadyExists`|O recurso que está sendo criado já existe.|
|`ItemNotFound` |O recurso solicitado não existe.|
|`NonBlankCellOffSheet`|A solicitação para inserir novas células não pode ser concluída, pois ela enviaria células não vazias para fora do final da planilha. Essas células não vazias podem aparecer vazias, mas têm valores em branco, parte da formatação ou uma fórmula. Exclua linhas ou colunas suficientes para liberar espaço para o que você deseja inserir e tente novamente.|
|`NotImplemented`|O recurso solicitado não foi implementado.|
|`RangeExceedsLimit`|A contagem de células no intervalo excedeu o número máximo com suporte. Consulte o artigo sobre [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.|
|`RequestAborted`|A solicitação foi anulada durante o tempo de execução.|
|`RequestPayloadSizeLimitExceeded`|O tamanho do conteúdo da solicitação excedeu o limite. Consulte o artigo sobre [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações. <br><br>Esse erro ocorre apenas no Excel na Web.|
|`ResponsePayloadSizeLimitExceeded`|O tamanho do conteúdo da resposta excedeu o limite. Consulte o artigo sobre [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.  <br><br>Esse erro ocorre apenas no Excel na Web.|
|`ServiceNotAvailable`|O serviço não está disponível.|
|`Unauthenticated` |Informações de autenticação necessárias estão ausentes ou inválidas.|
|`UnsupportedOperation`|Não há suporte para a operação que está sendo tentada.|
|`UnsupportedSheet`|Este tipo de planilha não tem suporte para essa operação, pois é uma macro ou uma planilha de gráfico.|

## <a name="see-also"></a>Confira também

- [Modelo de objeto do JavaScript do Excel em suplementos do Office](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
