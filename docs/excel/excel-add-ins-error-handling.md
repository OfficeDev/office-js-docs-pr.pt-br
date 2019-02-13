---
title: Tratamento de erros
description: ''
ms.date: 10/16/2018
localization_priority: Normal
ms.openlocfilehash: 8c6de5d2a22fdb4614742ddfb7fbf566780c0f0f
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29388959"
---
# <a name="error-handling"></a>Tratamento de erros

Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.

> [!NOTE]
> Para saber mais sobre o método **sync()** e a natureza assíncrona da API JavaScript do Excel, consulte [Principais conceitos de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md).

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
|InvalidArgument |O argumento é inválido, está ausente ou tem um formato incorreto.|
|InvalidRequest  |Não é possível processar a solicitação.|
|InvalidReference|Esta referência não é válida para a operação atual.|
|InvalidBinding  |Esta associação de objetos não é mais válida devido às atualizações anteriores.|
|InvalidSelection|A seleção atual é inválida para esta operação.|
|Unauthenticated |Informações de autenticação necessárias estão ausentes ou inválidas.|
|AccessDenied |Você não pode realizar a operação solicitada.|
|ItemNotFound |O recurso solicitado não existe.|
|ActivityLimitReached|O limite de atividades foi alcançado.|
|GeneralException|Ocorreu um erro interno ao processar a solicitação.|
|NotImplemented  |O recurso solicitado não foi implementado.|
|ServiceNotAvailable|O serviço não está disponível.|
|Conflito|A solicitação não pôde ser processada devido a um conflito.|
|ItemAlreadyExists|O recurso que está sendo criado já existe.|
|UnsupportedOperation|Não há suporte para a operação que está sendo tentada.|
|RequestAborted|A solicitação foi anulada durante o tempo de execução.|
|ApiNotAvailable|A API solicitada não está disponível.|
|InsertDeleteConflict|A tentativa de operação de exclusão ou inserção resultou em um conflito.|
|InvalidOperation|A tentativa de operação é inválida no objeto.|

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error)
