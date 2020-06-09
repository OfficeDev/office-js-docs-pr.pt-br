---
ms.date: 05/06/2020
description: Use tags JSDoc para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Normal
ms.openlocfilehash: f09fbbfcd028d773b9e9e25eb5eb43eb1d5a93cd
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609307"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="66b75-103">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="66b75-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="66b75-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as [marcações JSDoc](https://jsdoc.app/) são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="66b75-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="66b75-105">As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="66b75-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="66b75-106">O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="66b75-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="66b75-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="66b75-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="66b75-108">Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript.</span><span class="sxs-lookup"><span data-stu-id="66b75-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="66b75-109">Para saber mais, confira a marcação [@param](#param) e as seções [Tipos](#types).</span><span class="sxs-lookup"><span data-stu-id="66b75-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="66b75-110">Adicionando uma descrição a uma função</span><span class="sxs-lookup"><span data-stu-id="66b75-110">Adding a description to a function</span></span>

<span data-ttu-id="66b75-111">A descrição é exibida para o usuário como texto de ajuda quando eles precisam de ajuda para entender o que a função personalizada executa.</span><span class="sxs-lookup"><span data-stu-id="66b75-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="66b75-112">A descrição não requer nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="66b75-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="66b75-113">Basta digitar uma breve descrição de texto no comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="66b75-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="66b75-114">Em geral, a descrição é colocada no início da seção de comentários do JSDoc, mas funcionará independentemente de onde seja colocada.</span><span class="sxs-lookup"><span data-stu-id="66b75-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="66b75-115">Para ver exemplos das descrições de funções internas, abra o Excel, vá para a guia **Fórmulas** e escolha **Inserir função**.</span><span class="sxs-lookup"><span data-stu-id="66b75-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="66b75-116">Você pode navegar por todas as descrições de funções e também ver suas próprias funções personalizadas listadas.</span><span class="sxs-lookup"><span data-stu-id="66b75-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="66b75-117">No exemplo a seguir, a frase "Calcula o volume de uma esfera."</span><span class="sxs-lookup"><span data-stu-id="66b75-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="66b75-118">é a descrição da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="66b75-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="66b75-119">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="66b75-119">JSDoc Tags</span></span>
<span data-ttu-id="66b75-120">As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:</span><span class="sxs-lookup"><span data-stu-id="66b75-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="66b75-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="66b75-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="66b75-122">[@customfunction](#customfunction) nome de identificação</span><span class="sxs-lookup"><span data-stu-id="66b75-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="66b75-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="66b75-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="66b75-124">[@param](#param) _{type}_ nome e descrição</span><span class="sxs-lookup"><span data-stu-id="66b75-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="66b75-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="66b75-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="66b75-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="66b75-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="66b75-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="66b75-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="66b75-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="66b75-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="66b75-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="66b75-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="66b75-130">Indica que uma função personalizada executa uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="66b75-130">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="66b75-131">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="66b75-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="66b75-132">A função pode atribuir uma função à `oncanceled` propriedade para indicar o resultado quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="66b75-132">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="66b75-133">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, ela será considerada `@cancelable`, mesmo se a tag não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="66b75-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="66b75-134">Uma função não pode ter as tags `@cancelable` e `@streaming` ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="66b75-134">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="66b75-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="66b75-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="66b75-136">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="66b75-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="66b75-137">Essa marca indica que a função JavaScript/TypeScript é uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="66b75-137">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="66b75-138">É necessário criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="66b75-138">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="66b75-139">Veja a seguir um exemplo dessa marca.</span><span class="sxs-lookup"><span data-stu-id="66b75-139">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="66b75-140">id</span><span class="sxs-lookup"><span data-stu-id="66b75-140">id</span></span>

<span data-ttu-id="66b75-141">O `id` identifica uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="66b75-141">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="66b75-142">Se `id` não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="66b75-142">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="66b75-143">O `id` deve ser exclusivo para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="66b75-143">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="66b75-144">Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="66b75-144">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="66b75-145">No exemplo a seguir, incremento é o `id` e o `name` da função.</span><span class="sxs-lookup"><span data-stu-id="66b75-145">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="66b75-146">nome</span><span class="sxs-lookup"><span data-stu-id="66b75-146">name</span></span>

<span data-ttu-id="66b75-147">Fornece a exibição `name` da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="66b75-147">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="66b75-148">Se o nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="66b75-148">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="66b75-149">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="66b75-149">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="66b75-150">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="66b75-150">Must start with a letter.</span></span>
* <span data-ttu-id="66b75-151">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="66b75-151">Maximum length is 128 characters.</span></span>

<span data-ttu-id="66b75-152">No exemplo a seguir, Inc é a `id`da função e `increment` é o `name`.</span><span class="sxs-lookup"><span data-stu-id="66b75-152">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="66b75-153">descrição</span><span class="sxs-lookup"><span data-stu-id="66b75-153">description</span></span>

<span data-ttu-id="66b75-154">Uma descrição aparece para os usuários no Excel à medida que estão inserindo a função e especifica o que a função faz.</span><span class="sxs-lookup"><span data-stu-id="66b75-154">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="66b75-155">Uma descrição não exige nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="66b75-155">A description doesn't require any specific tag.</span></span> <span data-ttu-id="66b75-156">Adicione uma descrição a uma função personalizada acrescentando uma frase para descrever o que a função realiza dentro do comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="66b75-156">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="66b75-157">Por padrão, qualquer texto sem tags na seção de comentários JSDoc será a descrição da função.</span><span class="sxs-lookup"><span data-stu-id="66b75-157">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="66b75-158">No exemplo a seguir, a frase "Uma função que soma dois números" é a descrição da função personalizada com a propriedade id de `ADD`.</span><span class="sxs-lookup"><span data-stu-id="66b75-158">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="66b75-159">@helpurl</span><span class="sxs-lookup"><span data-stu-id="66b75-159">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="66b75-160">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="66b75-160">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="66b75-161">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="66b75-161">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="66b75-162">No exemplo a seguir, o `helpurl` é `www.contoso.com/weatherhelp` .</span><span class="sxs-lookup"><span data-stu-id="66b75-162">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a><span data-ttu-id="66b75-163">@param</span><span class="sxs-lookup"><span data-stu-id="66b75-163">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="66b75-164">JavaScript</span><span class="sxs-lookup"><span data-stu-id="66b75-164">JavaScript</span></span>

<span data-ttu-id="66b75-165">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="66b75-165">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="66b75-166">`{type}`Especifica as informações de tipo nas chaves.</span><span class="sxs-lookup"><span data-stu-id="66b75-166">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="66b75-167">Confira a seção [Tipos](#types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="66b75-167">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="66b75-168">Se nenhum tipo for especificado, o tipo padrão `any` será usado.</span><span class="sxs-lookup"><span data-stu-id="66b75-168">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="66b75-169">`name`Especifica o parâmetro ao qual a marca @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="66b75-169">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="66b75-170">É necessário.</span><span class="sxs-lookup"><span data-stu-id="66b75-170">It is required.</span></span>
* <span data-ttu-id="66b75-171">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="66b75-171">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="66b75-172">É opcional.</span><span class="sxs-lookup"><span data-stu-id="66b75-172">It is optional.</span></span>

<span data-ttu-id="66b75-173">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="66b75-173">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="66b75-174">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="66b75-174">Put square brackets around the parameter name.</span></span> <span data-ttu-id="66b75-175">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="66b75-175">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="66b75-176">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="66b75-176">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="66b75-177">O exemplo a seguir mostra uma função ADD que adiciona dois ou três números, com o terceiro número como um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="66b75-177">The following example shows a ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a><span data-ttu-id="66b75-178">TypeScript</span><span class="sxs-lookup"><span data-stu-id="66b75-178">TypeScript</span></span>

<span data-ttu-id="66b75-179">Sintaxe de TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="66b75-179">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="66b75-180">`name`Especifica o parâmetro ao qual a marca @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="66b75-180">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="66b75-181">É necessário.</span><span class="sxs-lookup"><span data-stu-id="66b75-181">It is required.</span></span>
* <span data-ttu-id="66b75-182">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="66b75-182">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="66b75-183">É opcional.</span><span class="sxs-lookup"><span data-stu-id="66b75-183">It is optional.</span></span>

<span data-ttu-id="66b75-184">Confira a seção [Tipos](#types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="66b75-184">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="66b75-185">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="66b75-185">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="66b75-186">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="66b75-186">Use an optional parameter.</span></span> <span data-ttu-id="66b75-187">Por exemplo: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="66b75-187">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="66b75-188">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="66b75-188">Give the parameter a default value.</span></span> <span data-ttu-id="66b75-189">Por exemplo: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="66b75-189">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="66b75-190">Para uma descrição detalhada do @param confira: [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="66b75-190">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="66b75-191">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="66b75-191">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="66b75-192">O exemplo a seguir mostra `add` a função que adiciona dois números.</span><span class="sxs-lookup"><span data-stu-id="66b75-192">The following example shows the `add` function that adds two numbers.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="requiresaddress"></a><span data-ttu-id="66b75-193">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="66b75-193">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="66b75-194">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="66b75-194">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="66b75-195">O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado.</span><span class="sxs-lookup"><span data-stu-id="66b75-195">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="66b75-196">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="66b75-196">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="66b75-197">@returns</span><span class="sxs-lookup"><span data-stu-id="66b75-197">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="66b75-198">Sintaxe: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="66b75-198">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="66b75-199">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="66b75-199">Provides the type for the return value.</span></span>

<span data-ttu-id="66b75-200">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="66b75-200">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="66b75-201">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="66b75-201">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="66b75-202">O exemplo a seguir mostra a `add` função que usa `@returns` marca.</span><span class="sxs-lookup"><span data-stu-id="66b75-202">The following example shows the `add` function that uses the `@returns` tag.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="streaming"></a><span data-ttu-id="66b75-203">@streaming</span><span class="sxs-lookup"><span data-stu-id="66b75-203">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="66b75-204">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="66b75-204">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="66b75-205">O último parâmetro é do tipo `CustomFunctions.StreamingInvocation<ResultType>` .</span><span class="sxs-lookup"><span data-stu-id="66b75-205">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="66b75-206">A função retorna `void` .</span><span class="sxs-lookup"><span data-stu-id="66b75-206">The function returns `void`.</span></span>

<span data-ttu-id="66b75-207">As funções de streaming não retornam valores diretamente, em vez de chamarem `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="66b75-207">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="66b75-208">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="66b75-208">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="66b75-209">`setResult()` pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="66b75-209">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="66b75-210">Para obter um exemplo de uma função de streaming e mais informações, confira [, criar uma função de streaming](./custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="66b75-210">For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="66b75-211">As funções de streaming não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="66b75-211">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="66b75-212">@volatile</span><span class="sxs-lookup"><span data-stu-id="66b75-212">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="66b75-213">Uma função volátil é aquela cujo resultado não é o mesmo de um momento para o outro, mesmo que não receba argumentos ou os argumentos não mudem.</span><span class="sxs-lookup"><span data-stu-id="66b75-213">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="66b75-214">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="66b75-214">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="66b75-215">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="66b75-215">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="66b75-216">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="66b75-216">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="66b75-217">A função a seguir é volátil e usa `@volatile` a marca.</span><span class="sxs-lookup"><span data-stu-id="66b75-217">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="66b75-218">Tipos</span><span class="sxs-lookup"><span data-stu-id="66b75-218">Types</span></span>

<span data-ttu-id="66b75-219">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="66b75-219">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="66b75-220">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="66b75-220">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="66b75-221">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="66b75-221">Value types</span></span>

<span data-ttu-id="66b75-222">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="66b75-222">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="66b75-223">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="66b75-223">Matrix type</span></span>

<span data-ttu-id="66b75-224">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="66b75-224">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="66b75-225">Por exemplo, o tipo `number[][]` indica uma matriz de números.</span><span class="sxs-lookup"><span data-stu-id="66b75-225">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="66b75-226">`string[][]` indica uma matriz de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="66b75-226">`string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="66b75-227">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="66b75-227">Error type</span></span>

<span data-ttu-id="66b75-228">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="66b75-228">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="66b75-229">Uma função de streaming pode indicar um erro chamando `setResult()` com um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="66b75-229">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="66b75-230">Promessa</span><span class="sxs-lookup"><span data-stu-id="66b75-230">Promise</span></span>

<span data-ttu-id="66b75-231">Uma função pode retornar uma promessa, que fornece o valor quando a promessa é resolvida.</span><span class="sxs-lookup"><span data-stu-id="66b75-231">A function can return a Promise, that provides the value when the promise is resolved.</span></span> <span data-ttu-id="66b75-232">Se a promessa for rejeitada, será gerado um erro.</span><span class="sxs-lookup"><span data-stu-id="66b75-232">If the promise is rejected, then it will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="66b75-233">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="66b75-233">Other types</span></span>

<span data-ttu-id="66b75-234">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="66b75-234">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="66b75-235">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="66b75-235">Next steps</span></span>
<span data-ttu-id="66b75-236">Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="66b75-236">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="66b75-237">Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="66b75-237">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="66b75-238">Confira também</span><span class="sxs-lookup"><span data-stu-id="66b75-238">See also</span></span>

* [<span data-ttu-id="66b75-239">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="66b75-239">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="66b75-240">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="66b75-240">Create custom functions in Excel</span></span>](custom-functions-overview.md)
