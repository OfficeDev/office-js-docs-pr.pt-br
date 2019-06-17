---
ms.date: 06/10/2019
description: Use tags JSDoc para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 960e1eca1e01aec21967733d802a5fdd48122cbc
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910298"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="930ca-103">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="930ca-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="930ca-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="930ca-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="930ca-105">As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="930ca-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="930ca-106">O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="930ca-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="930ca-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="930ca-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="930ca-108">Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript.</span><span class="sxs-lookup"><span data-stu-id="930ca-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="930ca-109">Para mais informações, confira a marcação [@param](#param) e a seção [Tipos](#types).</span><span class="sxs-lookup"><span data-stu-id="930ca-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="930ca-110">Adicionando uma descrição a uma função</span><span class="sxs-lookup"><span data-stu-id="930ca-110">Adding a description to a function</span></span>

<span data-ttu-id="930ca-111">A descrição é exibida para o usuário como texto de ajuda quando eles precisam de ajuda para entender o que a função personalizada executa.</span><span class="sxs-lookup"><span data-stu-id="930ca-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="930ca-112">A descrição não requer nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="930ca-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="930ca-113">Basta digitar uma breve descrição de texto no comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="930ca-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="930ca-114">Em geral, a descrição é colocada no início da seção de comentários do JSDoc, mas funcionará independentemente de onde seja colocada.</span><span class="sxs-lookup"><span data-stu-id="930ca-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="930ca-115">Para ver exemplos das descrições de funções internas, abra o Excel, vá para a guia **Fórmulas** e escolha **Inserir função**.</span><span class="sxs-lookup"><span data-stu-id="930ca-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="930ca-116">Você pode navegar por todas as descrições de funções e também ver suas próprias funções personalizadas listadas.</span><span class="sxs-lookup"><span data-stu-id="930ca-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="930ca-117">No exemplo a seguir, a frase "Calcula o volume de uma esfera."</span><span class="sxs-lookup"><span data-stu-id="930ca-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="930ca-118">é a descrição da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="930ca-118">is the description for the custom function.</span></span>

```JS
/**
/* Calculates the volume of a sphere
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="930ca-119">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="930ca-119">JSDoc Tags</span></span>
<span data-ttu-id="930ca-120">As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:</span><span class="sxs-lookup"><span data-stu-id="930ca-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="930ca-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="930ca-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="930ca-122">[@customfunction](#customfunction) nome de identificação</span><span class="sxs-lookup"><span data-stu-id="930ca-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="930ca-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="930ca-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="930ca-124">[@param](#param) _{type}_ nome e descrição</span><span class="sxs-lookup"><span data-stu-id="930ca-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="930ca-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="930ca-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="930ca-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="930ca-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="930ca-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="930ca-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="930ca-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="930ca-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="930ca-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="930ca-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="930ca-130">Indica que uma função personalizada deseja executar uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="930ca-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="930ca-131">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="930ca-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="930ca-132">A função pode atribuir uma função à propriedade `oncanceled` para denotar a ação a ser executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="930ca-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="930ca-133">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, ela será considerada `@cancelable`, mesmo se a tag não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="930ca-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="930ca-134">Uma função não pode ter as tags `@cancelable` e `@streaming` ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="930ca-134">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="930ca-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="930ca-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="930ca-136">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="930ca-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="930ca-137">Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="930ca-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="930ca-138">Essa marcação é necessária para criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="930ca-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="930ca-139">Também deve haver uma chamada para `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="930ca-139">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="930ca-140">id</span><span class="sxs-lookup"><span data-stu-id="930ca-140">id</span></span>

<span data-ttu-id="930ca-141">O `id` é um identificador invariável para a função customizada.</span><span class="sxs-lookup"><span data-stu-id="930ca-141">The id is used as the invariant identifier for the custom function stored in the document.</span></span>

* <span data-ttu-id="930ca-142">Se `id` não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="930ca-142">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="930ca-143">O `id` deve ser exclusivo para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="930ca-143">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="930ca-144">Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="930ca-144">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="930ca-145">nome</span><span class="sxs-lookup"><span data-stu-id="930ca-145">name</span></span>

<span data-ttu-id="930ca-146">Fornece a exibição `name` da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="930ca-146">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="930ca-147">Se o nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="930ca-147">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="930ca-148">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="930ca-148">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="930ca-149">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="930ca-149">Must start with a letter.</span></span>
* <span data-ttu-id="930ca-150">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="930ca-150">Maximum length is 128 characters.</span></span>

### <a name="description"></a><span data-ttu-id="930ca-151">description</span><span class="sxs-lookup"><span data-stu-id="930ca-151">description</span></span>

<span data-ttu-id="930ca-152">Uma descrição não exige nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="930ca-152">A description doesn't require any specific tag.</span></span> <span data-ttu-id="930ca-153">Adicione uma descrição a uma função personalizada acrescentando uma frase para descrever o que a função realiza dentro do comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="930ca-153">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="930ca-154">Por padrão, qualquer texto sem tags na seção de comentários JSDoc será a descrição da função.</span><span class="sxs-lookup"><span data-stu-id="930ca-154">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="930ca-155">A descrição aparece para os usuários no Excel quando eles entram na função.</span><span class="sxs-lookup"><span data-stu-id="930ca-155">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="930ca-156">No exemplo a seguir, a frase "Uma função que soma dois números" é a descrição da função personalizada com a propriedade id de `SUM`.</span><span class="sxs-lookup"><span data-stu-id="930ca-156">In the following example, the phrase "A function that sums two numbers" is the description for the custom function with the id property of `SUM`.</span></span>

```JS
/**
/* @customfunction SUM
/* A function that sums two numbers
...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="930ca-157">@helpurl</span><span class="sxs-lookup"><span data-stu-id="930ca-157">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="930ca-158">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="930ca-158">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="930ca-159">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="930ca-159">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="930ca-160">@param</span><span class="sxs-lookup"><span data-stu-id="930ca-160">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="930ca-161">JavaScript</span><span class="sxs-lookup"><span data-stu-id="930ca-161">JavaScript</span></span>

<span data-ttu-id="930ca-162">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="930ca-162">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="930ca-163">`{type}` deve especificar a informação de tipo entre chaves.</span><span class="sxs-lookup"><span data-stu-id="930ca-163">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="930ca-164">Confira [Tipos](##types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="930ca-164">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="930ca-165">Opcional: se não especificado, o tipo `any` será usado.</span><span class="sxs-lookup"><span data-stu-id="930ca-165">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="930ca-166">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="930ca-166">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="930ca-167">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="930ca-167">Required.</span></span>
* <span data-ttu-id="930ca-168">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="930ca-168">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="930ca-169">Opcional.</span><span class="sxs-lookup"><span data-stu-id="930ca-169">Optional.</span></span>

<span data-ttu-id="930ca-170">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="930ca-170">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="930ca-171">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="930ca-171">Put square brackets around the parameter name.</span></span> <span data-ttu-id="930ca-172">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="930ca-172">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="930ca-173">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="930ca-173">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="930ca-174">TypeScript</span><span class="sxs-lookup"><span data-stu-id="930ca-174">TypeScript</span></span>

<span data-ttu-id="930ca-175">Sintaxe de TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="930ca-175">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="930ca-176">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="930ca-176">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="930ca-177">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="930ca-177">Required.</span></span>
* <span data-ttu-id="930ca-178">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="930ca-178">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="930ca-179">Opcional.</span><span class="sxs-lookup"><span data-stu-id="930ca-179">Optional.</span></span>

<span data-ttu-id="930ca-180">Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="930ca-180">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="930ca-181">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="930ca-181">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="930ca-182">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="930ca-182">Use an optional parameter.</span></span> <span data-ttu-id="930ca-183">Por exemplo: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="930ca-183">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="930ca-184">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="930ca-184">Give the parameter a default value.</span></span> <span data-ttu-id="930ca-185">Por exemplo: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="930ca-185">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="930ca-186">Para uma descrição detalhada do @param confira: [JSDoc](https://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="930ca-186">For detailed description of the @param see: [JSDoc](https://usejsdoc.org/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="930ca-187">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="930ca-187">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="930ca-188">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="930ca-188">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="930ca-189">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="930ca-189">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="930ca-190">O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado.</span><span class="sxs-lookup"><span data-stu-id="930ca-190">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="930ca-191">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="930ca-191">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="930ca-192">@returns</span><span class="sxs-lookup"><span data-stu-id="930ca-192">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="930ca-193">Sintaxe: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="930ca-193">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="930ca-194">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="930ca-194">Provides the type for the return value.</span></span>

<span data-ttu-id="930ca-195">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="930ca-195">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="930ca-196">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="930ca-196">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="930ca-197">@streaming</span><span class="sxs-lookup"><span data-stu-id="930ca-197">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="930ca-198">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="930ca-198">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="930ca-199">O último parâmetro deve ser do tipo `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="930ca-199">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="930ca-200">A função deve retornar `void`.</span><span class="sxs-lookup"><span data-stu-id="930ca-200">The function should return `void`.</span></span>

<span data-ttu-id="930ca-201">As funções de streaming não retornam valores diretamente, mas devem chamar `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="930ca-201">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="930ca-202">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="930ca-202">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="930ca-203">`setResult()` pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="930ca-203">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="930ca-204">As funções de streaming não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="930ca-204">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="930ca-205">@volatile</span><span class="sxs-lookup"><span data-stu-id="930ca-205">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="930ca-206">Uma função volátil é aquela cujo resultado não é o mesmo de um momento para o outro, mesmo que não receba argumentos ou os argumentos não mudem.</span><span class="sxs-lookup"><span data-stu-id="930ca-206">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="930ca-207">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="930ca-207">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="930ca-208">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="930ca-208">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="930ca-209">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="930ca-209">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="930ca-210">Tipos</span><span class="sxs-lookup"><span data-stu-id="930ca-210">Types</span></span>

<span data-ttu-id="930ca-211">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="930ca-211">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="930ca-212">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="930ca-212">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="930ca-213">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="930ca-213">Value types</span></span>

<span data-ttu-id="930ca-214">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="930ca-214">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="930ca-215">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="930ca-215">Matrix type</span></span>

<span data-ttu-id="930ca-216">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="930ca-216">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="930ca-217">Por exemplo, o tipo `number[][]` indica uma matriz de números.</span><span class="sxs-lookup"><span data-stu-id="930ca-217">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="930ca-218">`string[][]` indica uma matriz de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="930ca-218">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="930ca-219">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="930ca-219">Error type</span></span>

<span data-ttu-id="930ca-220">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="930ca-220">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="930ca-221">Uma função de streaming pode indicar um erro chamando `setResult()` com um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="930ca-221">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="930ca-222">Promessa</span><span class="sxs-lookup"><span data-stu-id="930ca-222">Promise</span></span>

<span data-ttu-id="930ca-223">Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida.</span><span class="sxs-lookup"><span data-stu-id="930ca-223">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="930ca-224">Se a promessa for rejeitada, então é um erro.</span><span class="sxs-lookup"><span data-stu-id="930ca-224">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="930ca-225">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="930ca-225">Other types</span></span>

<span data-ttu-id="930ca-226">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="930ca-226">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="930ca-227">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="930ca-227">Next steps</span></span>
<span data-ttu-id="930ca-228">Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="930ca-228">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="930ca-229">Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="930ca-229">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="930ca-230">Confira também</span><span class="sxs-lookup"><span data-stu-id="930ca-230">See also</span></span>

* [<span data-ttu-id="930ca-231">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="930ca-231">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="930ca-232">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="930ca-232">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="930ca-233">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="930ca-233">Create custom functions in Excel</span></span>](custom-functions-overview.md)
