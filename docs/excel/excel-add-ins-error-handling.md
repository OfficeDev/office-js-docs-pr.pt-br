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

When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.

> [!NOTE]
> Para obter mais informações sobre o `sync()` método e a natureza assíncrona da API JavaScript do Excel, consulte [conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Práticas recomendadas

Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.

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

- **code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.

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
