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
# <a name="error-handling"></a><span data-ttu-id="51f27-102">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="51f27-102">Error handling</span></span>

<span data-ttu-id="51f27-103">Ao criar um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="51f27-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="51f27-104">Isso é fundamental devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="51f27-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="51f27-105">Para saber mais sobre o método **sync()** e a natureza assíncrona da API JavaScript do Excel, confira [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="51f27-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="51f27-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="51f27-106">Best practices</span></span>

<span data-ttu-id="51f27-107">Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="51f27-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="51f27-108">É recomendável usar o mesmo padrão ao criar um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="51f27-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="51f27-109">Erros de API</span><span class="sxs-lookup"><span data-stu-id="51f27-109">API errors</span></span> 

<span data-ttu-id="51f27-110">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="51f27-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="51f27-111">**code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`.</span><span class="sxs-lookup"><span data-stu-id="51f27-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="51f27-112">Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada.</span><span class="sxs-lookup"><span data-stu-id="51f27-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="51f27-113">Os códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="51f27-113">Error codes are not localized.</span></span> 

- <span data-ttu-id="51f27-114">**message**: A propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada.</span><span class="sxs-lookup"><span data-stu-id="51f27-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="51f27-115">A mensagem de erro não se destina ao usuário final; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento deve mostrar aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="51f27-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="51f27-116">**debugInfo**: Se estiver presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="51f27-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="51f27-117">Se você usar `console.log()` para exibir as mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor.</span><span class="sxs-lookup"><span data-stu-id="51f27-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="51f27-118">Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento nem em nenhum outro lugar do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="51f27-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="51f27-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="51f27-119">See also</span></span>

- [<span data-ttu-id="51f27-120">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="51f27-120">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="51f27-121">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="51f27-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
