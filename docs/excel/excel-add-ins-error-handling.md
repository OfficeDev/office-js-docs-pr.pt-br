---
title: Tratamento de erros
description: Saiba mais sobre a lógica de tratamento de erro da API JavaScript do Excel para considerar os erros de tempo de execução.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 8d410ae7eea315e14383b5aa08373ede3768cace
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006441"
---
# <a name="error-handling"></a>Tratamento de erros

Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.

> [!NOTE]
> Para obter mais informações sobre o `sync()` método e a natureza assíncrona da API JavaScript do Excel, consulte [conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md).

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
> Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento nem em qualquer outro lugar do aplicativo host.

## <a name="error-messages"></a>Mensagens de erro

A tabela a seguir é uma lista de erros que a API pode retornar.

|error.code | error.message |
|:----------|:--------------|
|`AccessDenied` |Você não pode realizar a operação solicitada.|
|`ActivityLimitReached`|O limite de atividades foi alcançado.|
|`ApiNotAvailable`|A API solicitada não está disponível.|
|`Conflict`|A solicitação não pôde ser processada devido a um conflito.|
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
|`NotImplemented`  |O recurso solicitado não foi implementado.|
|`RequestAborted`|A solicitação foi anulada durante o tempo de execução.|
|`ServiceNotAvailable`|O serviço não está disponível.|
|`Unauthenticated` |Informações de autenticação necessárias estão ausentes ou inválidas.|
|`UnsupportedOperation`|Não há suporte para a operação que está sendo tentada.|
|`UnsupportedSheet`|Este tipo de planilha não tem suporte para essa operação, pois é uma macro ou uma planilha de gráfico.|

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview)
