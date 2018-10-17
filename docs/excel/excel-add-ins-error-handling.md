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
# <a name="error-handling"></a><span data-ttu-id="64428-102">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="64428-102">Error handling</span></span>

<span data-ttu-id="64428-p101">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de manipulação de erros para considerar os erros de tempo de execução. Fazer isso é fundamental, devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="64428-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="64428-105">Para obter mais informações sobre o método **sync()** e a natureza assíncrona do Excel API do JavaScript, consulte [conceitos fundamentais de programação com a API do JavaScript do Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="64428-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="64428-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="64428-106">Best practices</span></span>

<span data-ttu-id="64428-p102">Ao longo dos exemplos de código nesta documentação, você notará que todas as chamadas para `Excel.run` são acompanhadas por uma instrução `catch` para detectar quaisquer erros que ocorram dentro de `Excel.run`. Recomendamos que você use o mesmo padrão ao criar um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="64428-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="64428-109">Erros de API</span><span class="sxs-lookup"><span data-stu-id="64428-109">API errors</span></span> 

<span data-ttu-id="64428-110">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="64428-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="64428-p103">**código**: A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes` . Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="64428-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span> 

- <span data-ttu-id="64428-p104">**mensagem**: A propriedade `message` de uma mensagem de erro contém um resumo do erro na seqüência localizada. A mensagem de erro não é destinada ao consumo por usuários finais; você deve usar o código de erro e a lógica de negócios apropriada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="64428-p104">**message**: The `message` property of an error message contains a summary of the error in the localized string. The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="64428-116">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para entender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="64428-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="64428-p105">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens só serão visíveis no servidor. Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="64428-p105">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server. End users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="64428-119">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="64428-119">Error Messages</span></span>

<span data-ttu-id="64428-120">A tabela a seguir define uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="64428-120">The following table defines a list of errors that the API may return.</span></span>

|<span data-ttu-id="64428-121">error.code</span><span class="sxs-lookup"><span data-stu-id="64428-121">error.code</span></span> | <span data-ttu-id="64428-122">error.message</span><span class="sxs-lookup"><span data-stu-id="64428-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="64428-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="64428-123">InvalidArgument</span></span> |<span data-ttu-id="64428-124">O argumento é inválido, ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="64428-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="64428-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="64428-125">InvalidRequest</span></span>  |<span data-ttu-id="64428-126">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="64428-126">Cannot process the request.</span></span>|
|<span data-ttu-id="64428-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="64428-127">InvalidReference</span></span>|<span data-ttu-id="64428-128">Essa referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="64428-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="64428-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="64428-129">InvalidBinding</span></span>  |<span data-ttu-id="64428-130">Essa associação de objetos não é mais válida devido a atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="64428-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="64428-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="64428-131">InvalidSelection</span></span>|<span data-ttu-id="64428-132">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="64428-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="64428-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="64428-133">Unauthenticated</span></span> |<span data-ttu-id="64428-134">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="64428-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="64428-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="64428-135">AccessDenied</span></span> |<span data-ttu-id="64428-136">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="64428-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="64428-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="64428-137">ItemNotFound</span></span> |<span data-ttu-id="64428-138">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="64428-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="64428-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="64428-139">ActivityLimitReached</span></span>|<span data-ttu-id="64428-140">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="64428-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="64428-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="64428-141">GeneralException</span></span>|<span data-ttu-id="64428-142">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="64428-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="64428-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="64428-143">NotImplemented</span></span>  |<span data-ttu-id="64428-144">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="64428-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="64428-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="64428-145">ServiceNotAvailable</span></span>|<span data-ttu-id="64428-146">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="64428-146">The service is unavailable.</span></span>|
|<span data-ttu-id="64428-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="64428-147">Conflict</span></span>|<span data-ttu-id="64428-148">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="64428-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="64428-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="64428-149">ItemAlreadyExists</span></span>|<span data-ttu-id="64428-150">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="64428-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="64428-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="64428-151">UnsupportedOperation</span></span>|<span data-ttu-id="64428-152">Não há suporte para a operação.</span><span class="sxs-lookup"><span data-stu-id="64428-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="64428-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="64428-153">RequestAborted</span></span>|<span data-ttu-id="64428-154">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="64428-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="64428-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="64428-155">ApiNotAvailable</span></span>|<span data-ttu-id="64428-156">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="64428-156">The requested API is not available.</span></span>|
|<span data-ttu-id="64428-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="64428-157">InsertDeleteConflict</span></span>|<span data-ttu-id="64428-158">A operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="64428-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="64428-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="64428-159">InvalidOperation</span></span>|<span data-ttu-id="64428-160">A operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="64428-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="64428-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="64428-161">See also</span></span>

- [<span data-ttu-id="64428-162">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="64428-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="64428-163">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="64428-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
