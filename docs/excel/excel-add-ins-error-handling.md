---
title: Tratamento de erros com a API JavaScript do Excel
description: Saiba mais sobre a lógica de tratamento de erro da API JavaScript do Excel para considerar os erros de tempo de execução.
ms.date: 10/22/2020
localization_priority: Normal
ms.openlocfilehash: a3b1bbfa7daba1b856bce35aa075d5b625bd9769
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740816"
---
# <a name="error-handling-with-the-excel-javascript-api"></a><span data-ttu-id="6ad3d-103">Tratamento de erros com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="6ad3d-103">Error handling with the Excel JavaScript API</span></span>

<span data-ttu-id="6ad3d-p101">Quando você cria um suplemento usando a API JavaScript do Excel, certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. Isso é fundamental devido à natureza assíncrona da API.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="6ad3d-106">Para obter mais informações sobre o `sync()` método e a natureza assíncrona da API JavaScript do Excel, consulte [modelo de objeto do Excel JavaScript em suplementos do Office](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="6ad3d-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="6ad3d-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="6ad3d-107">Best practices</span></span>

<span data-ttu-id="6ad3d-p102">Em todos os exemplos de código desta documentação, você notará que cada chamada a `Excel.run` é acompanhada de uma instrução `catch` para capturar todos os erros que ocorrem no `Excel.run`. É recomendável usar o mesmo padrão quando você cria um suplemento usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="6ad3d-110">Erros de API</span><span class="sxs-lookup"><span data-stu-id="6ad3d-110">API errors</span></span>

<span data-ttu-id="6ad3d-111">Quando uma solicitação da API JavaScript do Excel não é bem-sucedida, a API retorna um objeto de erro que contém as seguintes propriedades:</span><span class="sxs-lookup"><span data-stu-id="6ad3d-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="6ad3d-p103">**code**:  A propriedade `code` de uma mensagem de erro contém uma cadeia de caracteres que faz parte da lista `OfficeExtension.ErrorCodes` ou `Excel.ErrorCodes`. Por exemplo, o código de erro "InvalidReference" indica que a referência não é válida para a operação especificada. Os códigos de erro não são localizados.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="6ad3d-115">**message**: a propriedade `message` de uma mensagem de erro contém um resumo do erro na cadeia de caracteres localizada.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="6ad3d-116">A mensagem de erro não se destina aos usuários finais; você deve usar o código de erro e a lógica de negócios adequada para determinar a mensagem de erro que seu suplemento mostra aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="6ad3d-117">**debugInfo**: Quando presente, a propriedade `debugInfo` da mensagem de erro fornece informações adicionais que você pode usar para compreender a causa raiz do erro.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="6ad3d-118">Se você usar `console.log()` para imprimir mensagens de erro no console, essas mensagens ficarão visíveis apenas no servidor.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="6ad3d-119">Os usuários finais não verão essas mensagens de erro no painel de tarefas do suplemento ou em qualquer lugar no aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="6ad3d-120">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="6ad3d-120">Error Messages</span></span>

<span data-ttu-id="6ad3d-121">A tabela a seguir é uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="6ad3d-122">Código de erro</span><span class="sxs-lookup"><span data-stu-id="6ad3d-122">Error code</span></span> | <span data-ttu-id="6ad3d-123">Mensagem de erro</span><span class="sxs-lookup"><span data-stu-id="6ad3d-123">Error message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="6ad3d-124">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="6ad3d-125">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="6ad3d-126">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-126">The requested API is not available.</span></span>|
|`ApiNotFound`|<span data-ttu-id="6ad3d-127">A API que você está tentando usar não pôde ser encontrada.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-127">The API you are trying to use could not be found.</span></span> <span data-ttu-id="6ad3d-128">Ele pode estar disponível em uma versão mais recente do Excel.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-128">It may be available in a newer version of Excel.</span></span> <span data-ttu-id="6ad3d-129">Confira o artigo [conjuntos de requisitos da API JavaScript do Excel](../reference/requirement-sets/excel-api-requirement-sets.md) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-129">See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.</span></span>|
|`BadPassword`|<span data-ttu-id="6ad3d-130">A senha que você forneceu está incorreta.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-130">The password you supplied is incorrect.</span></span>|
|`Conflict`|<span data-ttu-id="6ad3d-131">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-131">Request could not be processed because of a conflict.</span></span>|
|`ContentLengthRequired`|<span data-ttu-id="6ad3d-132">Um `Content-length` cabeçalho HTTP está ausente.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-132">A `Content-length` HTTP header is missing.</span></span>|
|`GeneralException`|<span data-ttu-id="6ad3d-133">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-133">There was an internal error while processing the request.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="6ad3d-134">A tentativa de operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-134">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="6ad3d-135">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-135">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="6ad3d-136">Esta associação de objetos não é mais válida devido às atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-136">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="6ad3d-137">A tentativa de operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-137">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="6ad3d-138">Esta referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-138">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="6ad3d-139">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-139">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="6ad3d-140">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-140">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="6ad3d-141">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-141">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="6ad3d-142">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-142">The requested resource doesn't exist.</span></span>|
|`NonBlankCellOffSheet`|<span data-ttu-id="6ad3d-143">A solicitação para inserir novas células não pode ser concluída, pois ela enviaria células não vazias para fora do final da planilha.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-143">The request to insert new cells can't be completed because it would push non-empty cells off the end of the worksheet.</span></span> <span data-ttu-id="6ad3d-144">Essas células não vazias podem aparecer vazias, mas têm valores em branco, parte da formatação ou uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-144">These non-empty cells might appear empty but have blank values, some formatting, or a formula.</span></span> <span data-ttu-id="6ad3d-145">Exclua linhas ou colunas suficientes para liberar espaço para o que você deseja inserir e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-145">Delete enough rows or columns to make room for what you want to insert and then try again.</span></span>|
|`NotImplemented`|<span data-ttu-id="6ad3d-146">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-146">The requested feature isn't implemented.</span></span>|
|`RangeExceedsLimit`|<span data-ttu-id="6ad3d-147">A contagem de células no intervalo excedeu o número máximo com suporte.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-147">The cell count in the range has exceeded the maximum supported number.</span></span> <span data-ttu-id="6ad3d-148">Consulte o artigo sobre [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-148">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>|
|`RequestAborted`|<span data-ttu-id="6ad3d-149">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-149">The request was aborted during run time.</span></span>|
|`RequestPayloadSizeLimitExceeded`|<span data-ttu-id="6ad3d-150">O tamanho do conteúdo da solicitação excedeu o limite.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-150">The request payload size has exceeded the limit.</span></span> <span data-ttu-id="6ad3d-151">Consulte o artigo sobre [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-151">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span> <br><br><span data-ttu-id="6ad3d-152">Esse erro ocorre apenas no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-152">This error only occurs in Excel on the web.</span></span>|
|`ResponsePayloadSizeLimitExceeded`|<span data-ttu-id="6ad3d-153">O tamanho do conteúdo da resposta excedeu o limite.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-153">The response payload size has exceeded the limit.</span></span> <span data-ttu-id="6ad3d-154">Consulte o artigo sobre [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-154">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>  <br><br><span data-ttu-id="6ad3d-155">Esse erro ocorre apenas no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-155">This error only occurs in Excel on the web.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="6ad3d-156">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-156">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="6ad3d-157">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-157">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="6ad3d-158">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-158">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="6ad3d-159">Este tipo de planilha não tem suporte para essa operação, pois é uma macro ou uma planilha de gráfico.</span><span class="sxs-lookup"><span data-stu-id="6ad3d-159">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="6ad3d-160">Confira também</span><span class="sxs-lookup"><span data-stu-id="6ad3d-160">See also</span></span>

- [<span data-ttu-id="6ad3d-161">Modelo de objeto do JavaScript do Excel em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6ad3d-161">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6ad3d-162">Objeto OfficeExtension.Error (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="6ad3d-162">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
