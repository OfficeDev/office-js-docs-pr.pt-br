---
title: Tratamento de erros
description: Saiba mais sobre a lógica de tratamento de erro da API JavaScript do Excel para considerar os erros de tempo de execução.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bee5824d8854a55d5ac4041be1335ce239b31a9e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717163"
---
# <a name="error-handling"></a><span data-ttu-id="d6c27-103">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="d6c27-103">Error handling</span></span>

<span data-ttu-id="d6c27-p101">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="d6c27-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="d6c27-106">Para obter mais informações sobre `sync()` o método e a natureza assíncrona da API JavaScript do Excel, consulte [conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="d6c27-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="d6c27-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="d6c27-107">Best practices</span></span>

<span data-ttu-id="d6c27-p102">Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="d6c27-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="d6c27-110">Erros de API</span><span class="sxs-lookup"><span data-stu-id="d6c27-110">API errors</span></span>

<span data-ttu-id="d6c27-111">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="d6c27-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="d6c27-p103">**code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="d6c27-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="d6c27-115">**message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada.</span><span class="sxs-lookup"><span data-stu-id="d6c27-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="d6c27-116">A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="d6c27-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="d6c27-117">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="d6c27-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="d6c27-118">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor.</span><span class="sxs-lookup"><span data-stu-id="d6c27-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="d6c27-119">Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento nem em qualquer outro lugar do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="d6c27-119">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="d6c27-120">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="d6c27-120">Error Messages</span></span>

<span data-ttu-id="d6c27-121">A tabela a seguir é uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="d6c27-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="d6c27-122">error.code</span><span class="sxs-lookup"><span data-stu-id="d6c27-122">error.code</span></span> | <span data-ttu-id="d6c27-123">error.message</span><span class="sxs-lookup"><span data-stu-id="d6c27-123">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="d6c27-124">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="d6c27-124">InvalidArgument</span></span> |<span data-ttu-id="d6c27-125">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="d6c27-125">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="d6c27-126">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="d6c27-126">InvalidRequest</span></span>  |<span data-ttu-id="d6c27-127">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="d6c27-127">Cannot process the request.</span></span>|
|<span data-ttu-id="d6c27-128">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="d6c27-128">InvalidReference</span></span>|<span data-ttu-id="d6c27-129">Esta referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="d6c27-129">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="d6c27-130">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="d6c27-130">InvalidBinding</span></span>  |<span data-ttu-id="d6c27-131">Esta associação de objetos não é mais válida devido às atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="d6c27-131">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="d6c27-132">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="d6c27-132">InvalidSelection</span></span>|<span data-ttu-id="d6c27-133">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="d6c27-133">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="d6c27-134">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="d6c27-134">Unauthenticated</span></span> |<span data-ttu-id="d6c27-135">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="d6c27-135">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="d6c27-136">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="d6c27-136">AccessDenied</span></span> |<span data-ttu-id="d6c27-137">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="d6c27-137">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="d6c27-138">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="d6c27-138">ItemNotFound</span></span> |<span data-ttu-id="d6c27-139">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="d6c27-139">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="d6c27-140">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="d6c27-140">ActivityLimitReached</span></span>|<span data-ttu-id="d6c27-141">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="d6c27-141">Activity limit has been reached.</span></span>|
|<span data-ttu-id="d6c27-142">GeneralException</span><span class="sxs-lookup"><span data-stu-id="d6c27-142">GeneralException</span></span>|<span data-ttu-id="d6c27-143">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="d6c27-143">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="d6c27-144">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="d6c27-144">NotImplemented</span></span>  |<span data-ttu-id="d6c27-145">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="d6c27-145">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="d6c27-146">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="d6c27-146">ServiceNotAvailable</span></span>|<span data-ttu-id="d6c27-147">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="d6c27-147">The service is unavailable.</span></span>|
|<span data-ttu-id="d6c27-148">Conflito</span><span class="sxs-lookup"><span data-stu-id="d6c27-148">Conflict</span></span>|<span data-ttu-id="d6c27-149">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="d6c27-149">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="d6c27-150">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="d6c27-150">ItemAlreadyExists</span></span>|<span data-ttu-id="d6c27-151">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="d6c27-151">The resource being created already exists.</span></span>|
|<span data-ttu-id="d6c27-152">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="d6c27-152">UnsupportedOperation</span></span>|<span data-ttu-id="d6c27-153">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="d6c27-153">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="d6c27-154">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="d6c27-154">RequestAborted</span></span>|<span data-ttu-id="d6c27-155">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d6c27-155">The request was aborted during run time.</span></span>|
|<span data-ttu-id="d6c27-156">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="d6c27-156">ApiNotAvailable</span></span>|<span data-ttu-id="d6c27-157">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="d6c27-157">The requested API is not available.</span></span>|
|<span data-ttu-id="d6c27-158">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="d6c27-158">InsertDeleteConflict</span></span>|<span data-ttu-id="d6c27-159">A tentativa de operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="d6c27-159">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="d6c27-160">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="d6c27-160">InvalidOperation</span></span>|<span data-ttu-id="d6c27-161">A tentativa de operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="d6c27-161">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="d6c27-162">Confira também</span><span class="sxs-lookup"><span data-stu-id="d6c27-162">See also</span></span>

- [<span data-ttu-id="d6c27-163">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="d6c27-163">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d6c27-164">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="d6c27-164">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
