---
title: Tratamento de erros
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348083"
---
# <a name="error-handling"></a>Tratamento de erros

Ao criar um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros em tempo de execução. Isso é fundamental devido à natureza assíncrona da API.

> [!NOTE]
> Para saber mais sobre o método **sync()** e a natureza assíncrona da API JavaScript do Excel, confira [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md).

## <a name="best-practices"></a>Práticas recomendadas

Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão ao criar um suplemento usando as APIs JavaScript do Excel.

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

- **message**: A propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada. A mensagem de erro não se destina ao usuário final; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento deve mostrar aos usuários finais.

- **debugInfo**: Se estiver presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro. 

> [!NOTE]
> Se você usar `console.log()` para exibir as mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento nem em nenhum outro lugar do aplicativo host.

## <a name="see-also"></a>Confira também

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto OfficeExtension.Error (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
