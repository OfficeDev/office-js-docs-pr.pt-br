---
title: Tratamento de erros
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b07012516cbe15374d0707c157738117a9c8fe96
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459228"
---
# <a name="error-handling"></a><span data-ttu-id="99971-102">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="99971-102">Error handling</span></span>

<span data-ttu-id="99971-p101">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de manipulação de erros para considerar os erros de tempo de execução. Fazer isso é fundamental, devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="99971-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="99971-105">Para obter mais informações sobre o método **sync()** e a natureza assíncrona do Excel API do JavaScript, consulte [conceitos fundamentais de programação com a API do JavaScript do Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="99971-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="99971-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="99971-106">Best practices</span></span>

<span data-ttu-id="99971-p102">Ao longo dos exemplos de código nesta documentação, você notará que todas as chamadas para `Excel.run` são acompanhadas por uma instrução `catch` para detectar quaisquer erros que ocorram dentro de `Excel.run`. Recomendamos que você use o mesmo padrão ao criar um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="99971-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="99971-109">Erros de API</span><span class="sxs-lookup"><span data-stu-id="99971-109">API errors</span></span> 

<span data-ttu-id="99971-110">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="99971-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="99971-p103">**código**: A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes` . Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="99971-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span> 

- <span data-ttu-id="99971-p104">**mensagem**: A propriedade `message` de uma mensagem de erro contém um resumo do erro na seqüência localizada. A mensagem de erro não é destinada ao consumo por usuários finais; você deve usar o código de erro e a lógica de negócios apropriada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="99971-p104">**message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="99971-116">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para entender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="99971-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="99971-p105">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens só serão visíveis no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="99971-p105">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server. End users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="99971-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="99971-119">See also</span></span>

- [<span data-ttu-id="99971-120">Conceitos de programação fundamentais com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="99971-120">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="99971-121">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="99971-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
