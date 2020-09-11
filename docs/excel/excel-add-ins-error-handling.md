---
title: Tratamento de erros
description: Saiba mais sobre a lógica de tratamento de erro da API JavaScript do Excel para considerar os erros de tempo de execução.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: cb06ab890fc10afd327bde0eb708ff97fabbcdca
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430552"
---
# <a name="error-handling"></a><span data-ttu-id="1a03a-103">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="1a03a-103">Error handling</span></span>

<span data-ttu-id="1a03a-p101">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="1a03a-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="1a03a-106">Para obter mais informações sobre o `sync()` método e a natureza assíncrona da API JavaScript do Excel, consulte [modelo de objeto do Excel JavaScript em suplementos do Office](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="1a03a-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="1a03a-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="1a03a-107">Best practices</span></span>

<span data-ttu-id="1a03a-p102">Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="1a03a-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="1a03a-110">Erros de API</span><span class="sxs-lookup"><span data-stu-id="1a03a-110">API errors</span></span>

<span data-ttu-id="1a03a-111">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="1a03a-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="1a03a-p103">**code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="1a03a-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="1a03a-115">**message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada.</span><span class="sxs-lookup"><span data-stu-id="1a03a-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="1a03a-116">A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="1a03a-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="1a03a-117">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="1a03a-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="1a03a-118">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor.</span><span class="sxs-lookup"><span data-stu-id="1a03a-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="1a03a-119">Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="1a03a-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="1a03a-120">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="1a03a-120">Error Messages</span></span>

<span data-ttu-id="1a03a-121">A tabela a seguir é uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="1a03a-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="1a03a-122">error.code</span><span class="sxs-lookup"><span data-stu-id="1a03a-122">error.code</span></span> | <span data-ttu-id="1a03a-123">error.message</span><span class="sxs-lookup"><span data-stu-id="1a03a-123">error.message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="1a03a-124">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="1a03a-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="1a03a-125">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="1a03a-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="1a03a-126">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="1a03a-126">The requested API is not available.</span></span>|
|`Conflict`|<span data-ttu-id="1a03a-127">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="1a03a-127">Request could not be processed because of a conflict.</span></span>|
|`GeneralException`|<span data-ttu-id="1a03a-128">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="1a03a-128">There was an internal error while processing the request.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="1a03a-129">A tentativa de operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="1a03a-129">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="1a03a-130">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="1a03a-130">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="1a03a-131">Esta associação de objetos não é mais válida devido às atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="1a03a-131">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="1a03a-132">A tentativa de operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="1a03a-132">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="1a03a-133">Esta referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="1a03a-133">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="1a03a-134">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="1a03a-134">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="1a03a-135">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="1a03a-135">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="1a03a-136">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="1a03a-136">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="1a03a-137">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="1a03a-137">The requested resource doesn't exist.</span></span>|
|`NotImplemented`  |<span data-ttu-id="1a03a-138">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="1a03a-138">The requested feature isn't implemented.</span></span>|
|`RequestAborted`|<span data-ttu-id="1a03a-139">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="1a03a-139">The request was aborted during run time.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="1a03a-140">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="1a03a-140">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="1a03a-141">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="1a03a-141">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="1a03a-142">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="1a03a-142">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="1a03a-143">Este tipo de planilha não tem suporte para essa operação, pois é uma macro ou uma planilha de gráfico.</span><span class="sxs-lookup"><span data-stu-id="1a03a-143">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="1a03a-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="1a03a-144">See also</span></span>

- [<span data-ttu-id="1a03a-145">Modelo de objeto do JavaScript do Excel em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1a03a-145">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="1a03a-146">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="1a03a-146">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
