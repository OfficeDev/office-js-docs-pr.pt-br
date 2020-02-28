---
title: Tratamento de erros
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e3732af26aeaa6129a4b98d6cbb8e3caf501141f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325105"
---
# <a name="error-handling"></a><span data-ttu-id="66504-102">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="66504-102">Error handling</span></span>

<span data-ttu-id="66504-p101">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="66504-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="66504-105">Para obter mais informações sobre `sync()` o método e a natureza assíncrona da API JavaScript do Excel, consulte [conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="66504-105">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="66504-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="66504-106">Best practices</span></span>

<span data-ttu-id="66504-p102">Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="66504-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="66504-109">Erros de API</span><span class="sxs-lookup"><span data-stu-id="66504-109">API errors</span></span>

<span data-ttu-id="66504-110">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="66504-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="66504-p103">**code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="66504-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="66504-114">**message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada.</span><span class="sxs-lookup"><span data-stu-id="66504-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="66504-115">A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="66504-115">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="66504-116">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="66504-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="66504-117">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor.</span><span class="sxs-lookup"><span data-stu-id="66504-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="66504-118">Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento nem em qualquer outro lugar do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="66504-118">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="66504-119">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="66504-119">Error Messages</span></span>

<span data-ttu-id="66504-120">A tabela a seguir é uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="66504-120">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="66504-121">error.code</span><span class="sxs-lookup"><span data-stu-id="66504-121">error.code</span></span> | <span data-ttu-id="66504-122">error.message</span><span class="sxs-lookup"><span data-stu-id="66504-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="66504-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="66504-123">InvalidArgument</span></span> |<span data-ttu-id="66504-124">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="66504-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="66504-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="66504-125">InvalidRequest</span></span>  |<span data-ttu-id="66504-126">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="66504-126">Cannot process the request.</span></span>|
|<span data-ttu-id="66504-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="66504-127">InvalidReference</span></span>|<span data-ttu-id="66504-128">Esta referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="66504-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="66504-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="66504-129">InvalidBinding</span></span>  |<span data-ttu-id="66504-130">Esta associação de objetos não é mais válida devido às atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="66504-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="66504-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="66504-131">InvalidSelection</span></span>|<span data-ttu-id="66504-132">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="66504-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="66504-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="66504-133">Unauthenticated</span></span> |<span data-ttu-id="66504-134">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="66504-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="66504-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="66504-135">AccessDenied</span></span> |<span data-ttu-id="66504-136">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="66504-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="66504-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="66504-137">ItemNotFound</span></span> |<span data-ttu-id="66504-138">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="66504-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="66504-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="66504-139">ActivityLimitReached</span></span>|<span data-ttu-id="66504-140">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="66504-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="66504-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="66504-141">GeneralException</span></span>|<span data-ttu-id="66504-142">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="66504-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="66504-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="66504-143">NotImplemented</span></span>  |<span data-ttu-id="66504-144">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="66504-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="66504-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="66504-145">ServiceNotAvailable</span></span>|<span data-ttu-id="66504-146">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="66504-146">The service is unavailable.</span></span>|
|<span data-ttu-id="66504-147">Conflito</span><span class="sxs-lookup"><span data-stu-id="66504-147">Conflict</span></span>|<span data-ttu-id="66504-148">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="66504-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="66504-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="66504-149">ItemAlreadyExists</span></span>|<span data-ttu-id="66504-150">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="66504-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="66504-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="66504-151">UnsupportedOperation</span></span>|<span data-ttu-id="66504-152">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="66504-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="66504-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="66504-153">RequestAborted</span></span>|<span data-ttu-id="66504-154">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="66504-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="66504-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="66504-155">ApiNotAvailable</span></span>|<span data-ttu-id="66504-156">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="66504-156">The requested API is not available.</span></span>|
|<span data-ttu-id="66504-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="66504-157">InsertDeleteConflict</span></span>|<span data-ttu-id="66504-158">A tentativa de operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="66504-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="66504-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="66504-159">InvalidOperation</span></span>|<span data-ttu-id="66504-160">A tentativa de operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="66504-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="66504-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="66504-161">See also</span></span>

- [<span data-ttu-id="66504-162">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="66504-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="66504-163">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="66504-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
