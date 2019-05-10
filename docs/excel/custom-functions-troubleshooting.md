---
ms.date: 05/03/2019
description: Solução de problemas comuns em funções personalizadas do Excel.
title: Solução de problemas das funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 04da6d58c2610130961a1b89d2b9a1101b54bcb2
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628008"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="04283-103">Solução de problemas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="04283-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="04283-104">Ao desenvolver funções personalizadas, você poderá encontrar erros no produto durante a criação e testes das funções.</span><span class="sxs-lookup"><span data-stu-id="04283-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="04283-105">Para resolver problemas, você pode [habilitar o log de tempo de execução para capturar erros](#enable-runtime-logging) e consultar as [mensagens de erro nativas do Excel](#check-for-excel-error-messages).</span><span class="sxs-lookup"><span data-stu-id="04283-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="04283-106">Além disso, verifique se há erros comuns, como [deixar promessas não resolvidas](#ensure-promises-return) e esquecer de [associar as funções](#my-functions-wont-load-associate-functions).</span><span class="sxs-lookup"><span data-stu-id="04283-106">Also, check for common mistakes such as not [verifying ssl certificates](#ensure-promises-return) properly, [leaving promises unresolved](#my-functions-wont-load-associate-functions), and forgetting to associate your functions.</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="04283-107">Habilitar o log de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="04283-107">Enable runtime logging</span></span>

<span data-ttu-id="04283-108">Se estiver testando o suplemento do Office no Windows, você deverá [habilitar o log de tempo de execução](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="04283-108">If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span> <span data-ttu-id="04283-109">O log de tempo de execução entrega instruções `console.log` a um arquivo de log separado criado para ajudar você a descobrir problemas.</span><span class="sxs-lookup"><span data-stu-id="04283-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="04283-110">As instruções abrangem vários erros, incluindo os relacionados ao arquivo de manifesto XML do suplemento, condições do tempo de execução ou a instalação de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="04283-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span>  <span data-ttu-id="04283-111">Saiba mais sobre o log de tempo de execução em [Usar o log de tempo de execução para depurar seu suplemento](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="04283-111">For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span>  

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="04283-112">Verificar se há mensagens de erro do Excel</span><span class="sxs-lookup"><span data-stu-id="04283-112">Check for Excel error messages</span></span>

<span data-ttu-id="04283-113">O Excel tem diversas mensagens de erro internas que serão retornadas para uma célula se houver um erro de cálculo.</span><span class="sxs-lookup"><span data-stu-id="04283-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="04283-114">As funções personalizadas usam apenas as seguintes mensagens de erro: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` e `#BUSY!`.</span><span class="sxs-lookup"><span data-stu-id="04283-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

<span data-ttu-id="04283-115">Geralmente, estes erros correspondem aos erros que você já deve estar familiarizado no Excel.</span><span class="sxs-lookup"><span data-stu-id="04283-115">Generally, these errors correspond to the errors you might already be familiar with in Excel.</span></span> <span data-ttu-id="04283-116">Existem apenas algumas exceções específicas para as funções personalizadas, listadas aqui:</span><span class="sxs-lookup"><span data-stu-id="04283-116">The are only a few exceptions specific to custom functions, listed here:</span></span>

- <span data-ttu-id="04283-117">Um erro `#NAME` geralmente significa que houve um problema ao registrar as suas funções.</span><span class="sxs-lookup"><span data-stu-id="04283-117">A `#NAME` error generally means there has been an issue registering your functions.</span></span>
- <span data-ttu-id="04283-118">Um erro `#VALUE` normalmente indica um erro no arquivo de script das funções.</span><span class="sxs-lookup"><span data-stu-id="04283-118">A `#VALUE` error typically indicates an error in the functions' script file.</span></span>
- <span data-ttu-id="04283-119">Um erro `#N/A` também pode ser um sinal de que esta função, embora registrada, não pode ser executada.</span><span class="sxs-lookup"><span data-stu-id="04283-119">A `#N/A` error is also maybe a sign that that function while registered could not be run.</span></span> <span data-ttu-id="04283-120">Isto é normalmente devido à um comando `CustomFunctions.associate` em falta.</span><span class="sxs-lookup"><span data-stu-id="04283-120">This is typically due to a missing `CustomFunctions.associate` command.</span></span>
- <span data-ttu-id="04283-121">Um erro `#REF!` pode indicar que o nome da sua função é o mesmo nome de uma função em um suplemento já existente.</span><span class="sxs-lookup"><span data-stu-id="04283-121">A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.</span></span>

## <a name="clear-the-office-cache"></a><span data-ttu-id="04283-122">Limpar o cache do Office</span><span class="sxs-lookup"><span data-stu-id="04283-122">Clear the Office cache</span></span>

<span data-ttu-id="04283-123">Informações sobre funções personalizadas são armazenadas em cache pelo Office.</span><span class="sxs-lookup"><span data-stu-id="04283-123">Information about custom functions is cached by Office.</span></span> <span data-ttu-id="04283-124">Às vezes, ao desenvolver e recarregar repetidamente um suplemento com funções personalizadas, as suas alterações podem não aparecer.</span><span class="sxs-lookup"><span data-stu-id="04283-124">Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear.</span></span> <span data-ttu-id="04283-125">Isso pode ser corrigido limpando o cache do Office.</span><span class="sxs-lookup"><span data-stu-id="04283-125">You can fix this by clearing the Office cache.</span></span> <span data-ttu-id="04283-126">Para mais informações, consulte a seção «Limpar o Cache do Office» no artigo [Validar e solucionar problemas com seu manifesto](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)</span><span class="sxs-lookup"><span data-stu-id="04283-126">For more information, see the "Clear the Office cache" section in the article [Validate and troubleshoot issues with your manifest](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)</span></span>

## <a name="common-issues"></a><span data-ttu-id="04283-127">Problemas comuns</span><span class="sxs-lookup"><span data-stu-id="04283-127">Common issues</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="04283-128">Minhas funções não carregam: associar funções</span><span class="sxs-lookup"><span data-stu-id="04283-128">My functions won't load: associate functions</span></span>

<span data-ttu-id="04283-129">No arquivo de script das funções personalizadas, você precisa associar cada função personalizada à respectiva ID especificada no [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="04283-129">In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="04283-130">Isso é feito usando o método `CustomFunctions.associate()`.</span><span class="sxs-lookup"><span data-stu-id="04283-130">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="04283-131">Normalmente, essa chamada de método é feita após cada função ou no final do arquivo de script.</span><span class="sxs-lookup"><span data-stu-id="04283-131">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="04283-132">Se uma função personalizada não estiver associada, ele não funcionará.</span><span class="sxs-lookup"><span data-stu-id="04283-132">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="04283-133">O exemplo a seguir mostra uma função add, seguida pelo nome `add` da função que está sendo associada a `ADD` da id JSON correspondente.</span><span class="sxs-lookup"><span data-stu-id="04283-133">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="04283-134">Saiba mais sobre esse processo em [Associar os nomes de função com metadados JSON](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="04283-134">For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a><span data-ttu-id="04283-135">Não é possível abrir um suplemento de um localhost: utilize uma exceção de loopback local</span><span class="sxs-lookup"><span data-stu-id="04283-135">Can't open add-in from localhost: use a local loopback exception</span></span>

<span data-ttu-id="04283-136">Se você vir o erro "Não é possível abrir este suplemento de um localhost", será necessário habilitar uma exceção de loopback local.</span><span class="sxs-lookup"><span data-stu-id="04283-136">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception.</span></span> <span data-ttu-id="04283-137">Para obter detalhes sobre como fazer isso, confira [este artigo de suporte da Microsoft](https://support.microsoft.com/pt-BR/help/4490419/local-loopback-exemption-does-not-work).</span><span class="sxs-lookup"><span data-stu-id="04283-137">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/pt-BR/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="04283-138">Garantir que as promessas retornem resultados</span><span class="sxs-lookup"><span data-stu-id="04283-138">Ensure promises return</span></span>

<span data-ttu-id="04283-139">Quando o Excel está aguardando a conclusão de uma função personalizada, ele exibe #BUSY!</span><span class="sxs-lookup"><span data-stu-id="04283-139">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="04283-140">na célula.</span><span class="sxs-lookup"><span data-stu-id="04283-140">in the cell.</span></span> <span data-ttu-id="04283-141">Se o código da função personalizada retornar uma promessa, mas a promessa não retornar um resultado, o Excel continuará exibindo #BUSY!.</span><span class="sxs-lookup"><span data-stu-id="04283-141">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!.</span></span> <span data-ttu-id="04283-142">Verifique suas funções para garantir que as promessas estejam retornando corretamente um resultado para uma célula.</span><span class="sxs-lookup"><span data-stu-id="04283-142">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a><span data-ttu-id="04283-143">Erro: O servidor de desenvolvimento já está em execução na porta 3000</span><span class="sxs-lookup"><span data-stu-id="04283-143">Error: The dev server is already running on port 3000</span></span>

<span data-ttu-id="04283-144">Às vezes, ao executar `npm start` você poderá ver um erro que o servidor de desenvolvimento já está executando na porta 3000 (ou qualquer outra porta que o seu suplemento use).</span><span class="sxs-lookup"><span data-stu-id="04283-144">Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses).</span></span> <span data-ttu-id="04283-145">Você pode parar o servidor de desenvolvimento executando `npm stop` ou fechando a janela Node.js.</span><span class="sxs-lookup"><span data-stu-id="04283-145">You can stop the dev server by running `npm stop` or by closing the Node.js window.</span></span> <span data-ttu-id="04283-146">Mas em alguns casos, poderá levar alguns minutos para que o servidor de desenvolvimento realmente pare de executar.</span><span class="sxs-lookup"><span data-stu-id="04283-146">But in some cases in can take a few minutes for the dev server to actually stop running.</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="04283-147">Fornecer comentários</span><span class="sxs-lookup"><span data-stu-id="04283-147">Reporting Feedback</span></span>

<span data-ttu-id="04283-148">Se você tiver problemas que não estão descritos aqui, fale conosco.</span><span class="sxs-lookup"><span data-stu-id="04283-148">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="04283-149">Há duas maneiras de relatar problemas.</span><span class="sxs-lookup"><span data-stu-id="04283-149">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="04283-150">No Excel para Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="04283-150">In Excel on Windows or Mac</span></span>

<span data-ttu-id="04283-151">Se estiver usando o Excel para Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel.</span><span class="sxs-lookup"><span data-stu-id="04283-151">If using Excel for Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="04283-152">Para fazer isso, selecione **Arquivo -> Comentários -> Enviar um Rosto Triste**.</span><span class="sxs-lookup"><span data-stu-id="04283-152">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="04283-153">Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.</span><span class="sxs-lookup"><span data-stu-id="04283-153">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="04283-154">No Github</span><span class="sxs-lookup"><span data-stu-id="04283-154">In Github</span></span>

<span data-ttu-id="04283-155">Sinta-se à vontade para enviar problemas encontrados através do recurso "Comentários do conteúdo" na parte inferior de todas as páginas de documentação ou [informe um novo problema diretamente no repositório de funções personalizadas](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="04283-155">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="04283-156">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="04283-156">Next steps</span></span>
<span data-ttu-id="04283-157">Saiba como [depurar as suas funções personalizadas](custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="04283-157">Learn how to [debug your custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="04283-158">Confira também</span><span class="sxs-lookup"><span data-stu-id="04283-158">See also</span></span>

* [<span data-ttu-id="04283-159">Geração automática de metadados das funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="04283-159">Custom functions metadata</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="04283-160">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="04283-160">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="04283-161">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="04283-161">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="04283-162">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="04283-162">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="04283-163">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="04283-163">Create custom functions in Excel</span></span>](custom-functions-overview.md)
