---
title: Tratamento de erros
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: caba29f7d6949cc6d9df1498ac0a3d4f5de6c4ee
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579811"
---
# <a name="error-handling"></a>Tratamento de erros

Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de manipulação de erros para considerar os erros de tempo de execução. Fazer isso é fundamental, devido à natureza assíncrona da API.

> [!NOTE]
> Para obter mais informações sobre o método **sync()** e a natureza assíncrona do Excel API do JavaScript, consulte [conceitos fundamentais de programação com a API do JavaScript do Excel](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Práticas recomendadas

Ao longo dos exemplos de código nesta documentação, você notará que todas as chamadas para `Excel.run` são acompanhadas por uma instrução `catch` para detectar quaisquer erros que ocorram dentro de `Excel.run`. Recomendamos que você use o mesmo padrão ao criar um suplemento usando as APIs JavaScript do Excel.

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

- **código**: A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes` . Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Códigos de erro não são localizados. 

- **mensagem**: A propriedade `message` de uma mensagem de erro contém um resumo do erro na seqüência localizada. A mensagem de erro não é destinada ao consumo por usuários finais; você deve usar o código de erro e a lógica de negócios apropriada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.

- **debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para entender a causa raiz do erro. 

> [!NOTE]
> Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens só serão visíveis no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo host.

## <a name="error-messages"></a>Mensagens de erro

A tabela a seguir define uma lista de erros que a API pode retornar.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |O argumento é inválido, ausente ou tem um formato incorreto.|
|InvalidRequest  |Não é possível processar a solicitação.|
|InvalidReference|Essa referência não é válida para a operação atual.|
|InvalidBinding  |Essa associação de objetos não é mais válida devido a atualizações anteriores.|
|InvalidSelection|A seleção atual é inválida para esta operação.|
|Unauthenticated |Informações de autenticação necessárias estão ausentes ou inválidas.|
|AccessDenied |Você não pode realizar a operação solicitada.|
|ItemNotFound |O recurso solicitado não existe.|
|ActivityLimitReached|O limite de atividades foi alcançado.|
|GeneralException|Ocorreu um erro interno ao processar a solicitação.|
|NotImplemented  |O recurso solicitado não foi implementado.|
|ServiceNotAvailable|O serviço não está disponível.|
|Conflict|A solicitação não pôde ser processada devido a um conflito.|
|ItemAlreadyExists|O recurso que está sendo criado já existe.|
|UnsupportedOperation|Não há suporte para a operação.|
|RequestAborted|A solicitação foi anulada durante o tempo de execução.|
|ApiNotAvailable|A API solicitada não está disponível.|
|InsertDeleteConflict|A operação de exclusão ou inserção resultou em um conflito.|
|InvalidOperation|A operação é inválida no objeto.|

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
