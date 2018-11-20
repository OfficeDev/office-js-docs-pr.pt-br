---
title: Tratamento de erros
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 52d1c88fef0f4e3af075ed625f856b029353a963
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533711"
---
# <a name="error-handling"></a><span data-ttu-id="11df9-102">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="11df9-102">Error handling</span></span>

<span data-ttu-id="11df9-103">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="11df9-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="11df9-104">Isso é fundamental devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="11df9-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="11df9-105">Para saber mais sobre o método **sync()** e a natureza assíncrona da API JavaScript do Excel, consulte [Principais conceitos de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="11df9-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="11df9-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="11df9-106">Best practices</span></span>

<span data-ttu-id="11df9-107">Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="11df9-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="11df9-108">É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="11df9-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="11df9-109">Erros de API</span><span class="sxs-lookup"><span data-stu-id="11df9-109">API errors</span></span>

<span data-ttu-id="11df9-110">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="11df9-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="11df9-111">**code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`.</span><span class="sxs-lookup"><span data-stu-id="11df9-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="11df9-112">Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada.</span><span class="sxs-lookup"><span data-stu-id="11df9-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="11df9-113">Os códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="11df9-113">Error codes are not localized.</span></span>

- <span data-ttu-id="11df9-114">**message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada.</span><span class="sxs-lookup"><span data-stu-id="11df9-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="11df9-115">A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="11df9-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="11df9-116">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="11df9-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="11df9-117">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor.</span><span class="sxs-lookup"><span data-stu-id="11df9-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="11df9-118">Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento nem em qualquer outro lugar do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="11df9-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="11df9-119">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="11df9-119">Error Messages</span></span>

<span data-ttu-id="11df9-120">A tabela a seguir é uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="11df9-120">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="11df9-121">error.code</span><span class="sxs-lookup"><span data-stu-id="11df9-121">ErrorCode</span></span> | <span data-ttu-id="11df9-122">error.message</span><span class="sxs-lookup"><span data-stu-id="11df9-122">ErrorMessage</span></span> |
|:----------|:--------------|
|<span data-ttu-id="11df9-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="11df9-123">InvalidArgument</span></span> |<span data-ttu-id="11df9-124">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="11df9-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="11df9-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="11df9-125">invalidRequest</span></span>  |<span data-ttu-id="11df9-126">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="11df9-126">Cannot process the request.</span></span>|
|<span data-ttu-id="11df9-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="11df9-127">InvalidReference</span></span>|<span data-ttu-id="11df9-128">Esta referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="11df9-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="11df9-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="11df9-129">InvalidBinding</span></span>  |<span data-ttu-id="11df9-130">Esta associação de objetos não é mais válida devido às atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="11df9-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="11df9-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="11df9-131">InvalidSelection</span></span>|<span data-ttu-id="11df9-132">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="11df9-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="11df9-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="11df9-133">unauthenticated</span></span> |<span data-ttu-id="11df9-134">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="11df9-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="11df9-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="11df9-135">accessDenied</span></span> |<span data-ttu-id="11df9-136">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="11df9-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="11df9-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="11df9-137">itemNotFound</span></span> |<span data-ttu-id="11df9-138">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="11df9-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="11df9-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="11df9-139">activityLimitReached</span></span>|<span data-ttu-id="11df9-140">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="11df9-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="11df9-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="11df9-141">generalException</span></span>|<span data-ttu-id="11df9-142">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="11df9-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="11df9-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="11df9-143">NotImplemented</span></span>  |<span data-ttu-id="11df9-144">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="11df9-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="11df9-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="11df9-145">serviceNotAvailable</span></span>|<span data-ttu-id="11df9-146">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="11df9-146">The service is unavailable.</span></span>|
|<span data-ttu-id="11df9-147">Conflito</span><span class="sxs-lookup"><span data-stu-id="11df9-147">Conflict</span></span>|<span data-ttu-id="11df9-148">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="11df9-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="11df9-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="11df9-149">ItemAlreadyExists</span></span>|<span data-ttu-id="11df9-150">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="11df9-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="11df9-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="11df9-151">UnsupportedOperation</span></span>|<span data-ttu-id="11df9-152">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="11df9-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="11df9-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="11df9-153">RequestAborted</span></span>|<span data-ttu-id="11df9-154">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="11df9-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="11df9-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="11df9-155">ApiNotAvailable</span></span>|<span data-ttu-id="11df9-156">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="11df9-156">The requested API is not available.</span></span>|
|<span data-ttu-id="11df9-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="11df9-157">InsertDeleteConflict</span></span>|<span data-ttu-id="11df9-158">A tentativa de operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="11df9-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="11df9-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="11df9-159">InvalidOperation</span></span>|<span data-ttu-id="11df9-160">A tentativa de operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="11df9-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="11df9-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="11df9-161">See also</span></span>

- [<span data-ttu-id="11df9-162">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="11df9-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="11df9-163">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="11df9-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
