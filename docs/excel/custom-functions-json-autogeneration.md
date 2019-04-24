---
ms.date: 04/03/2019
description: Use marcações JSDOC para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Criar metadados JSON para funções personalizadas (visualização)
localization_priority: Priority
ms.openlocfilehash: 2efe2a9a5a83ba60ef327273d5bd599f82916d48
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449250"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="eda30-103">Criar metadados JSON para funções personalizadas (visualização)</span><span class="sxs-lookup"><span data-stu-id="eda30-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="eda30-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="eda30-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="eda30-105">As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="eda30-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="eda30-106">O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="eda30-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="eda30-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="eda30-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="eda30-108">Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript.</span><span class="sxs-lookup"><span data-stu-id="eda30-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="eda30-109">Para mais informações, confira a marcação [@param](#param) e a seção [Tipos](#types).</span><span class="sxs-lookup"><span data-stu-id="eda30-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="eda30-110">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="eda30-110">JSDoc Tags</span></span>
<span data-ttu-id="eda30-111">As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:</span><span class="sxs-lookup"><span data-stu-id="eda30-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="eda30-112">@cancelable</span><span class="sxs-lookup"><span data-stu-id="eda30-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="eda30-113">[@customfunction](#customfunction) nome de identificação</span><span class="sxs-lookup"><span data-stu-id="eda30-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="eda30-114">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="eda30-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="eda30-115">[@param](#param) _{type}_ nome e descrição</span><span class="sxs-lookup"><span data-stu-id="eda30-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="eda30-116">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="eda30-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="eda30-117">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="eda30-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="eda30-118">@streaming</span><span class="sxs-lookup"><span data-stu-id="eda30-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="eda30-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="eda30-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="eda30-120">@cancelable</span><span class="sxs-lookup"><span data-stu-id="eda30-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="eda30-121">Indica que uma função personalizada deseja executar uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="eda30-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="eda30-122">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="eda30-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="eda30-123">A função pode atribuir uma função à propriedade `oncanceled` para denotar a ação a ser executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="eda30-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="eda30-124">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, será considerado `@cancelable` mesmo se a marcação não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="eda30-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="eda30-125">Uma função não pode ter ao mesmo tempo as marcações `@cancelable` e `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="eda30-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="eda30-126">@customfunction</span><span class="sxs-lookup"><span data-stu-id="eda30-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="eda30-127">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="eda30-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="eda30-128">Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="eda30-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="eda30-129">Essa marcação é necessária para criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="eda30-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="eda30-130">Também deve haver uma chamada para `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="eda30-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="eda30-131">id</span><span class="sxs-lookup"><span data-stu-id="eda30-131">id</span></span> 

<span data-ttu-id="eda30-132">O id é usado como o identificador invariável da função personalizada armazenada no documento.</span><span class="sxs-lookup"><span data-stu-id="eda30-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="eda30-133">Ele não deve mudar.</span><span class="sxs-lookup"><span data-stu-id="eda30-133">It should not change.</span></span>

* <span data-ttu-id="eda30-134">Se o ID não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="eda30-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="eda30-135">O id deve ser exclusivo, para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="eda30-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="eda30-136">Os caracteres permitidos são limitados a: A-Z, a-z, 0-9 e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="eda30-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="eda30-137">nome</span><span class="sxs-lookup"><span data-stu-id="eda30-137">name</span></span>

<span data-ttu-id="eda30-138">Fornece o nome de exibição para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="eda30-138">Provides the display name for the custom function.</span></span> 

* <span data-ttu-id="eda30-139">Se nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="eda30-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="eda30-140">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="eda30-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="eda30-141">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="eda30-141">Must start with a letter.</span></span>
* <span data-ttu-id="eda30-142">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="eda30-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="eda30-143">@helpurl</span><span class="sxs-lookup"><span data-stu-id="eda30-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="eda30-144">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="eda30-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="eda30-145">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="eda30-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="eda30-146">@param</span><span class="sxs-lookup"><span data-stu-id="eda30-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="eda30-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="eda30-147">JavaScript</span></span>

<span data-ttu-id="eda30-148">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="eda30-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="eda30-149">`{type}` deve especificar a informação de tipo entre chaves.</span><span class="sxs-lookup"><span data-stu-id="eda30-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="eda30-150">Confira [Tipos](##types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="eda30-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="eda30-151">Opcional: se não especificado, o tipo `any` será usado.</span><span class="sxs-lookup"><span data-stu-id="eda30-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="eda30-152">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="eda30-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="eda30-153">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="eda30-153">Required.</span></span>
* <span data-ttu-id="eda30-154">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="eda30-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="eda30-155">Opcional.</span><span class="sxs-lookup"><span data-stu-id="eda30-155">Optional.</span></span>

<span data-ttu-id="eda30-156">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="eda30-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="eda30-157">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="eda30-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="eda30-158">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="eda30-158">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="eda30-159">TypeScript</span><span class="sxs-lookup"><span data-stu-id="eda30-159">TypeScript</span></span>

<span data-ttu-id="eda30-160">Sintaxe de TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="eda30-160">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="eda30-161">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="eda30-161">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="eda30-162">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="eda30-162">Required.</span></span>
* <span data-ttu-id="eda30-163">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="eda30-163">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="eda30-164">Opcional.</span><span class="sxs-lookup"><span data-stu-id="eda30-164">Optional.</span></span>

<span data-ttu-id="eda30-165">Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="eda30-165">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="eda30-166">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="eda30-166">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="eda30-167">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="eda30-167">Use an optional parameter.</span></span> <span data-ttu-id="eda30-168">Por exemplo: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="eda30-168">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="eda30-169">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="eda30-169">Give the parameter a default value.</span></span> <span data-ttu-id="eda30-170">Por exemplo: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="eda30-170">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="eda30-171">Para uma descrição detalhada do @param confira: [JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="eda30-171">For detailed description of the @param see: [JSDoc](http://usejsdoc.org/tags-param.html)</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="eda30-172">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="eda30-172">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="eda30-173">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="eda30-173">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="eda30-174">O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado.</span><span class="sxs-lookup"><span data-stu-id="eda30-174">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="eda30-175">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="eda30-175">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="eda30-176">@returns</span><span class="sxs-lookup"><span data-stu-id="eda30-176">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="eda30-177">Sintaxe: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="eda30-177">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="eda30-178">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="eda30-178">Provides the type for the return value.</span></span>

<span data-ttu-id="eda30-179">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="eda30-179">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="eda30-180">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="eda30-180">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="eda30-181">@streaming</span><span class="sxs-lookup"><span data-stu-id="eda30-181">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="eda30-182">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="eda30-182">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="eda30-183">O último parâmetro deve ser do tipo `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="eda30-183">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="eda30-184">A função deve retornar `void`.</span><span class="sxs-lookup"><span data-stu-id="eda30-184">The function should return `void`.</span></span>

<span data-ttu-id="eda30-185">As funções de streaming não retornam valores diretamente, mas em vez disso devem chamar `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="eda30-185">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="eda30-186">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="eda30-186">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="eda30-187">`setResult()` pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="eda30-187">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="eda30-188">Funções de transmissão não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="eda30-188">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="eda30-189">@volatile</span><span class="sxs-lookup"><span data-stu-id="eda30-189">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="eda30-190">Uma função volátil é aquela cujo resultado não pode ser assumido como sendo o mesmo de um momento para o outro, mesmo que não receba argumentos ou que os argumentos não sejam alterados.</span><span class="sxs-lookup"><span data-stu-id="eda30-190">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="eda30-191">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="eda30-191">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="eda30-192">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="eda30-192">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="eda30-193">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="eda30-193">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="eda30-194">Tipos</span><span class="sxs-lookup"><span data-stu-id="eda30-194">Types</span></span>

<span data-ttu-id="eda30-195">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="eda30-195">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="eda30-196">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="eda30-196">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="eda30-197">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="eda30-197">Value types</span></span>

<span data-ttu-id="eda30-198">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="eda30-198">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="eda30-199">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="eda30-199">Matrix type</span></span>

<span data-ttu-id="eda30-200">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="eda30-200">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="eda30-201">Por exemplo, o tipo `number[][]` indica uma matriz de números.</span><span class="sxs-lookup"><span data-stu-id="eda30-201">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="eda30-202">`string[][]` indica uma matriz de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="eda30-202">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="eda30-203">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="eda30-203">Error type</span></span>

<span data-ttu-id="eda30-204">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="eda30-204">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="eda30-205">Uma função de streaming pode indicar um erro chamando setResult () com um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="eda30-205">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="eda30-206">Promessa</span><span class="sxs-lookup"><span data-stu-id="eda30-206">Promise</span></span>

<span data-ttu-id="eda30-207">Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida.</span><span class="sxs-lookup"><span data-stu-id="eda30-207">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="eda30-208">Se a promessa for rejeitada, então é um erro.</span><span class="sxs-lookup"><span data-stu-id="eda30-208">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="eda30-209">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="eda30-209">Other types</span></span>

<span data-ttu-id="eda30-210">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="eda30-210">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="eda30-211">Confira também</span><span class="sxs-lookup"><span data-stu-id="eda30-211">See also</span></span>

* [<span data-ttu-id="eda30-212">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="eda30-212">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="eda30-213">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="eda30-213">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="eda30-214">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="eda30-214">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="eda30-215">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="eda30-215">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="eda30-216">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="eda30-216">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="eda30-217">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="eda30-217">Custom functions debugging</span></span>](custom-functions-debugging.md)
