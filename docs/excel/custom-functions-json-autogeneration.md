---
ms.date: 04/03/2019
description: Use marcações JSDOC para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Criar metadados JSON para funções personalizadas (visualização)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478947"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="cab11-103">Criar metadados JSON para funções personalizadas (visualização)</span><span class="sxs-lookup"><span data-stu-id="cab11-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="cab11-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cab11-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="cab11-105">As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="cab11-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="cab11-106">O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="cab11-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="cab11-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cab11-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="cab11-108">Os tipos de parâmetro de função podem ser fornecidos usando a marcação [@param](#param) no JavaScript, ou o [Tipo de função](http://www.typescriptlang.org/docs/handbook/functions.html) no TypeScript.</span><span class="sxs-lookup"><span data-stu-id="cab11-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](http://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="cab11-109">Para mais informações, consulte a marcação [@param](#param) e a seção [Tipos](#Types).</span><span class="sxs-lookup"><span data-stu-id="cab11-109">For more information, see the [@param](#param) tag and [Types](#Types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="cab11-110">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="cab11-110">JSDoc Tags</span></span>
<span data-ttu-id="cab11-111">As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:</span><span class="sxs-lookup"><span data-stu-id="cab11-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [@cancelable](#cancelable)
* <span data-ttu-id="cab11-112">[@customfunction](#customfunction) nome id</span><span class="sxs-lookup"><span data-stu-id="cab11-112">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="cab11-113">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="cab11-113">URL</span></span>
* <span data-ttu-id="cab11-114">[@param](#param) _{type}_ descrição do nome</span><span class="sxs-lookup"><span data-stu-id="cab11-114">[@param](#param) _{type}_ name description</span></span>
* [@requiresAddress](#requiresAddress)
* <span data-ttu-id="cab11-115">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="cab11-115">Type</span></span>
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

<span data-ttu-id="cab11-116">Indica que uma função personalizada deseja executar uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="cab11-116">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="cab11-117">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="cab11-117">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="cab11-118">A função pode atribuir uma função à propriedade `oncanceled` para denotar a ação a ser executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="cab11-118">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="cab11-119">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, será considerado `@cancelable` mesmo se a marcação não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="cab11-119">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="cab11-120">Uma função não pode ter ao mesmo tempo as marcações `@cancelable` e `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="cab11-120">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

<span data-ttu-id="cab11-121">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="cab11-121">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="cab11-122">Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="cab11-122">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="cab11-123">Essa marcação é necessária para criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cab11-123">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="cab11-124">Também deve haver uma chamada para</span><span class="sxs-lookup"><span data-stu-id="cab11-124">There should also be a call to</span></span> `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a><span data-ttu-id="cab11-125">id</span><span class="sxs-lookup"><span data-stu-id="cab11-125">id</span></span> 

<span data-ttu-id="cab11-126">O id é usado como o identificador invariável da função personalizada armazenada no documento.</span><span class="sxs-lookup"><span data-stu-id="cab11-126">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="cab11-127">Ele não deve mudar.</span><span class="sxs-lookup"><span data-stu-id="cab11-127">It should not change.</span></span>

* <span data-ttu-id="cab11-128">Se o ID não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="cab11-128">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="cab11-129">O id deve ser exclusivo, para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cab11-129">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="cab11-130">Os caracteres permitidos são limitados a: A-Z, a-z, 0-9 e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="cab11-130">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="cab11-131">nome</span><span class="sxs-lookup"><span data-stu-id="cab11-131">name</span></span>

<span data-ttu-id="cab11-132">Fornece o nome de exibição para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cab11-132">Provides the display name for the custom function.</span></span> 

* <span data-ttu-id="cab11-133">Se nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="cab11-133">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="cab11-134">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="cab11-134">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="cab11-135">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="cab11-135">Must begin with a letter.</span></span>
* <span data-ttu-id="cab11-136">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="cab11-136">Maximum length is 255 characters.</span></span>

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

<span data-ttu-id="cab11-137">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="cab11-137">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="cab11-138">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="cab11-138">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="cab11-139">JavaScript</span><span class="sxs-lookup"><span data-stu-id="cab11-139">JavaScript</span></span>

<span data-ttu-id="cab11-140">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="cab11-140">JavaScript Syntax: @param {type} name _description_</span></span>

* `{type}` <span data-ttu-id="cab11-141">deve especificar as informações de tipo entre chaves.</span><span class="sxs-lookup"><span data-stu-id="cab11-141">should specify the type info within curly braces.</span></span> <span data-ttu-id="cab11-142">Confira [Tipos](##types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="cab11-142">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="cab11-143">Opcional: se não especificado, o tipo `any` será usado.</span><span class="sxs-lookup"><span data-stu-id="cab11-143">Optional: if not specified, the type `any` will be used.</span></span>
* `name` <span data-ttu-id="cab11-144">especifica para qual parâmetro a etiqueta @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="cab11-144">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="cab11-145">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="cab11-145">Required.</span></span>
* `description` <span data-ttu-id="cab11-146">fornece a descrição que aparece no Excel para o parâmetro da função.</span><span class="sxs-lookup"><span data-stu-id="cab11-146">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="cab11-147">Opcional.</span><span class="sxs-lookup"><span data-stu-id="cab11-147">Optional.</span></span>

<span data-ttu-id="cab11-148">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="cab11-148">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="cab11-149">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cab11-149">Put square brackets around the parameter name.</span></span> <span data-ttu-id="cab11-150">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="cab11-150">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="cab11-151">TypeScript</span><span class="sxs-lookup"><span data-stu-id="cab11-151">TypeScript</span></span>

<span data-ttu-id="cab11-152">Sintaxe do TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="cab11-152">TypeScript Syntax: @param name _description_</span></span>

* `name` <span data-ttu-id="cab11-153">especifica para qual parâmetro a etiqueta @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="cab11-153">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="cab11-154">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="cab11-154">Required.</span></span>
* `description` <span data-ttu-id="cab11-155">fornece a descrição que aparece no Excel para o parâmetro da função.</span><span class="sxs-lookup"><span data-stu-id="cab11-155">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="cab11-156">Opcional.</span><span class="sxs-lookup"><span data-stu-id="cab11-156">Optional.</span></span>

<span data-ttu-id="cab11-157">Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="cab11-157">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="cab11-158">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="cab11-158">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="cab11-159">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="cab11-159">Use an optional parameter.</span></span> <span data-ttu-id="cab11-160">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="cab11-160">For example:</span></span> `function f(text?: string)`
* <span data-ttu-id="cab11-161">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="cab11-161">Give the parameter a default value.</span></span> <span data-ttu-id="cab11-162">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="cab11-162">For example:</span></span> `function f(text: string = "abc")`

<span data-ttu-id="cab11-163">Para uma descrição detalhada do @param, confira: [JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="cab11-163">For a detailed description of the code, see "HelloData Details."</span></span>

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

<span data-ttu-id="cab11-164">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="cab11-164">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="cab11-165">O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado.</span><span class="sxs-lookup"><span data-stu-id="cab11-165">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="cab11-166">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="cab11-166">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a>@returns
<a id="returns"/>

<span data-ttu-id="cab11-167">Sintaxe: @returns {_tipo_}</span><span class="sxs-lookup"><span data-stu-id="cab11-167">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="cab11-168">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="cab11-168">Provides the type for the return value.</span></span>

<span data-ttu-id="cab11-169">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="cab11-169">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="cab11-170">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="cab11-170">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

<span data-ttu-id="cab11-171">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="cab11-171">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="cab11-172">O último parâmetro deve ser do tipo `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="cab11-172">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="cab11-173">A função deve retornar `void`.</span><span class="sxs-lookup"><span data-stu-id="cab11-173">The function should return `void`.</span></span>

<span data-ttu-id="cab11-174">As funções de streaming não retornam valores diretamente, mas em vez disso devem chamar `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cab11-174">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="cab11-175">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="cab11-175">Exceptions thrown by a streaming function are ignored.</span></span> `setResult()` <span data-ttu-id="cab11-176">pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="cab11-176">may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="cab11-177">As funções de streaming não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="cab11-177">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

<span data-ttu-id="cab11-178">Uma função volátil é aquela cujo resultado não pode ser assumido como sendo o mesmo de um momento para o outro, mesmo que não receba argumentos ou que os argumentos não sejam alterados.</span><span class="sxs-lookup"><span data-stu-id="cab11-178">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="cab11-179">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="cab11-179">Excel reevaluates cells that contain volatile functions, together with all dependents, every time that it recalculates.</span></span> <span data-ttu-id="cab11-180">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="cab11-180">For this reason, too much reliance on volatile functions can make recalculation times slow.</span></span>

<span data-ttu-id="cab11-181">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="cab11-181">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="cab11-182">Tipos</span><span class="sxs-lookup"><span data-stu-id="cab11-182">Types</span></span>

<span data-ttu-id="cab11-183">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="cab11-183">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="cab11-184">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="cab11-184">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="cab11-185">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="cab11-185">Value types</span></span>

<span data-ttu-id="cab11-186">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="cab11-186">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="cab11-187">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="cab11-187">Matrix type</span></span>

<span data-ttu-id="cab11-188">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="cab11-188">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="cab11-189">Por exemplo, o tipo `number[][]` indica uma matriz de números.</span><span class="sxs-lookup"><span data-stu-id="cab11-189">For example, the type `number[][]` indicates a matrix of numbers.</span></span> `string[][]` <span data-ttu-id="cab11-190">indica uma matriz de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="cab11-190">indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="cab11-191">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="cab11-191">Error Type</span></span>

<span data-ttu-id="cab11-192">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="cab11-192">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="cab11-193">Uma função de streaming pode indicar um erro chamando setResult () com um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="cab11-193">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="cab11-194">Promessa</span><span class="sxs-lookup"><span data-stu-id="cab11-194">Promise object.</span></span>

<span data-ttu-id="cab11-195">Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida.</span><span class="sxs-lookup"><span data-stu-id="cab11-195">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="cab11-196">Se a promessa for rejeitada, então é um erro.</span><span class="sxs-lookup"><span data-stu-id="cab11-196">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="cab11-197">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="cab11-197">Other types of groups</span></span>

<span data-ttu-id="cab11-198">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="cab11-198">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="cab11-199">Confira também</span><span class="sxs-lookup"><span data-stu-id="cab11-199">See also</span></span>

* [<span data-ttu-id="cab11-200">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cab11-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="cab11-201">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="cab11-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="cab11-202">Práticas recomendadas para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cab11-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cab11-203">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cab11-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="cab11-204">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="cab11-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="cab11-205">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cab11-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
