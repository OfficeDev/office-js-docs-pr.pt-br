---
ms.date: 05/03/2019
description: Use marcações JSDOC para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 67026e7c19580c3420638b4f37e333e50fce1b44
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589129"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="7c3f7-103">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="7c3f7-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="7c3f7-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="7c3f7-105">As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="7c3f7-106">O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="7c3f7-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="7c3f7-108">Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="7c3f7-109">Para mais informações, confira a marcação [@param](#param) e a seção [Tipos](#types).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="7c3f7-110">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="7c3f7-110">JSDoc Tags</span></span>
<span data-ttu-id="7c3f7-111">As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:</span><span class="sxs-lookup"><span data-stu-id="7c3f7-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="7c3f7-112">@cancelable</span><span class="sxs-lookup"><span data-stu-id="7c3f7-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="7c3f7-113">[@customfunction](#customfunction) nome de identificação</span><span class="sxs-lookup"><span data-stu-id="7c3f7-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="7c3f7-114">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="7c3f7-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="7c3f7-115">[@param](#param) _{type}_ nome e descrição</span><span class="sxs-lookup"><span data-stu-id="7c3f7-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="7c3f7-116">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="7c3f7-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="7c3f7-117">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="7c3f7-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="7c3f7-118">@streaming</span><span class="sxs-lookup"><span data-stu-id="7c3f7-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="7c3f7-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="7c3f7-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="7c3f7-120">@cancelable</span><span class="sxs-lookup"><span data-stu-id="7c3f7-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="7c3f7-121">Indica que uma função personalizada deseja executar uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="7c3f7-122">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="7c3f7-123">A função pode atribuir uma função à propriedade `oncanceled` para denotar a ação a ser executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="7c3f7-124">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, será considerado `@cancelable` mesmo se a marcação não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="7c3f7-125">Uma função não pode ter ao mesmo tempo as marcações `@cancelable` e `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="7c3f7-126">@customfunction</span><span class="sxs-lookup"><span data-stu-id="7c3f7-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="7c3f7-127">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="7c3f7-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="7c3f7-128">Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="7c3f7-129">Essa marcação é necessária para criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="7c3f7-130">Também deve haver uma chamada para `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="7c3f7-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="7c3f7-131">id</span><span class="sxs-lookup"><span data-stu-id="7c3f7-131">id</span></span>

<span data-ttu-id="7c3f7-132">O id é usado como o identificador invariável da função personalizada armazenada no documento.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="7c3f7-133">Ele não deve mudar.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-133">It should not change.</span></span>

* <span data-ttu-id="7c3f7-134">Se o ID não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="7c3f7-135">O id deve ser exclusivo, para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="7c3f7-136">Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="7c3f7-137">nome</span><span class="sxs-lookup"><span data-stu-id="7c3f7-137">name</span></span>

<span data-ttu-id="7c3f7-138">Fornece o nome de exibição para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-138">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="7c3f7-139">Se nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="7c3f7-140">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="7c3f7-141">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-141">Must start with a letter.</span></span>
* <span data-ttu-id="7c3f7-142">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="7c3f7-143">@helpurl</span><span class="sxs-lookup"><span data-stu-id="7c3f7-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="7c3f7-144">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="7c3f7-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="7c3f7-145">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="7c3f7-146">@param</span><span class="sxs-lookup"><span data-stu-id="7c3f7-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="7c3f7-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="7c3f7-147">JavaScript</span></span>

<span data-ttu-id="7c3f7-148">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="7c3f7-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="7c3f7-149">`{type}` deve especificar a informação de tipo entre chaves.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="7c3f7-150">Confira [Tipos](##types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="7c3f7-151">Opcional: se não especificado, o tipo `any` será usado.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="7c3f7-152">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="7c3f7-153">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-153">Required.</span></span>
* <span data-ttu-id="7c3f7-154">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="7c3f7-155">Opcional.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-155">Optional.</span></span>

<span data-ttu-id="7c3f7-156">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="7c3f7-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="7c3f7-157">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="7c3f7-158">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-158">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="7c3f7-159">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-159">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="7c3f7-160">TypeScript</span><span class="sxs-lookup"><span data-stu-id="7c3f7-160">TypeScript</span></span>

<span data-ttu-id="7c3f7-161">Sintaxe de TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="7c3f7-161">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="7c3f7-162">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-162">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="7c3f7-163">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-163">Required.</span></span>
* <span data-ttu-id="7c3f7-164">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-164">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="7c3f7-165">Opcional.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-165">Optional.</span></span>

<span data-ttu-id="7c3f7-166">Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-166">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="7c3f7-167">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="7c3f7-167">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="7c3f7-168">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-168">Use an optional parameter.</span></span> <span data-ttu-id="7c3f7-169">Por exemplo: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="7c3f7-169">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="7c3f7-170">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-170">Give the parameter a default value.</span></span> <span data-ttu-id="7c3f7-171">Por exemplo: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="7c3f7-171">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="7c3f7-172">Para uma descrição detalhada do @param confira: [JSDoc](https://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="7c3f7-172">For detailed description of the @param see: [JSDoc](https://usejsdoc.org/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="7c3f7-173">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-173">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="7c3f7-174">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="7c3f7-174">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="7c3f7-175">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-175">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="7c3f7-176">O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-176">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="7c3f7-177">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-177">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="7c3f7-178">@returns</span><span class="sxs-lookup"><span data-stu-id="7c3f7-178">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="7c3f7-179">Sintaxe: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="7c3f7-179">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="7c3f7-180">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-180">Provides the type for the return value.</span></span>

<span data-ttu-id="7c3f7-181">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-181">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="7c3f7-182">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-182">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="7c3f7-183">@streaming</span><span class="sxs-lookup"><span data-stu-id="7c3f7-183">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="7c3f7-184">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-184">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="7c3f7-185">O último parâmetro deve ser do tipo `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-185">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="7c3f7-186">A função deve retornar `void`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-186">The function should return `void`.</span></span>

<span data-ttu-id="7c3f7-187">As funções de streaming não retornam valores diretamente, mas em vez disso devem chamar `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-187">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="7c3f7-188">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-188">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="7c3f7-189">`setResult()` pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-189">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="7c3f7-190">Funções de transmissão não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-190">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="7c3f7-191">@volatile</span><span class="sxs-lookup"><span data-stu-id="7c3f7-191">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="7c3f7-192">Uma função volátil é aquela cujo resultado não pode ser assumido como sendo o mesmo de um momento para o outro, mesmo que não receba argumentos ou que os argumentos não sejam alterados.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-192">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="7c3f7-193">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-193">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="7c3f7-194">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-194">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="7c3f7-195">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-195">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="7c3f7-196">Tipos</span><span class="sxs-lookup"><span data-stu-id="7c3f7-196">Types</span></span>

<span data-ttu-id="7c3f7-197">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-197">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="7c3f7-198">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-198">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="7c3f7-199">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="7c3f7-199">Value types</span></span>

<span data-ttu-id="7c3f7-200">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-200">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="7c3f7-201">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="7c3f7-201">Matrix type</span></span>

<span data-ttu-id="7c3f7-202">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-202">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="7c3f7-203">Por exemplo, o tipo `number[][]` indica uma matriz de números.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-203">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="7c3f7-204">`string[][]` indica uma matriz de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-204">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="7c3f7-205">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="7c3f7-205">Error type</span></span>

<span data-ttu-id="7c3f7-206">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-206">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="7c3f7-207">Uma função de streaming pode indicar um erro chamando setResult () com um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-207">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="7c3f7-208">Promessa</span><span class="sxs-lookup"><span data-stu-id="7c3f7-208">Promise</span></span>

<span data-ttu-id="7c3f7-209">Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-209">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="7c3f7-210">Se a promessa for rejeitada, então é um erro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-210">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="7c3f7-211">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="7c3f7-211">Other types</span></span>

<span data-ttu-id="7c3f7-212">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="7c3f7-212">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7c3f7-213">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="7c3f7-213">Next steps</span></span>
<span data-ttu-id="7c3f7-214">Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-214">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="7c3f7-215">Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-215">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7c3f7-216">Confira também</span><span class="sxs-lookup"><span data-stu-id="7c3f7-216">See also</span></span>

* [<span data-ttu-id="7c3f7-217">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="7c3f7-217">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="7c3f7-218">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="7c3f7-218">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="7c3f7-219">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="7c3f7-219">Create custom functions in Excel</span></span>](custom-functions-overview.md)
