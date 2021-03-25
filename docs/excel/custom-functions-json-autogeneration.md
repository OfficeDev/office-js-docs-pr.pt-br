---
ms.date: 03/15/2021
description: Use tags JSDoc para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Normal
ms.openlocfilehash: e31059de78e9daedc31c9b0a8605b5352fd0ed94
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178045"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="5eca7-103">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eca7-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="5eca7-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as [marcações JSDoc](https://jsdoc.app/) são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="5eca7-105">As marcações JSDoc são usadas no momento da criação para criar o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="5eca7-105">The JSDoc tags are then used at build time to create the JSON metadata file.</span></span> <span data-ttu-id="5eca7-106">O uso de marcas JSDoc salva você do esforço de editar manualmente o arquivo de [metadados JSON.](custom-functions-json.md)</span><span class="sxs-lookup"><span data-stu-id="5eca7-106">Using JSDoc tags saves you from the effort of [manually editing the JSON metadata file](custom-functions-json.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="5eca7-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="5eca7-108">Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript.</span><span class="sxs-lookup"><span data-stu-id="5eca7-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="5eca7-109">Para saber mais, confira a marcação [@param](#param) e as seções [Tipos](#types).</span><span class="sxs-lookup"><span data-stu-id="5eca7-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="5eca7-110">Adicionando uma descrição a uma função</span><span class="sxs-lookup"><span data-stu-id="5eca7-110">Adding a description to a function</span></span>

<span data-ttu-id="5eca7-111">A descrição é exibida para o usuário como texto de ajuda quando eles precisam de ajuda para entender o que a função personalizada executa.</span><span class="sxs-lookup"><span data-stu-id="5eca7-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="5eca7-112">A descrição não requer nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="5eca7-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="5eca7-113">Basta digitar uma breve descrição de texto no comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="5eca7-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="5eca7-114">Em geral, a descrição é colocada no início da seção de comentários do JSDoc, mas funcionará independentemente de onde seja colocada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="5eca7-115">Para ver exemplos das descrições de funções internas, abra o Excel, vá para a guia **Fórmulas** e escolha **Inserir função**.</span><span class="sxs-lookup"><span data-stu-id="5eca7-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="5eca7-116">Você pode navegar por todas as descrições de funções e também ver suas próprias funções personalizadas listadas.</span><span class="sxs-lookup"><span data-stu-id="5eca7-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="5eca7-117">No exemplo a seguir, a frase "Calcula o volume de uma esfera."</span><span class="sxs-lookup"><span data-stu-id="5eca7-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="5eca7-118">é a descrição da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="5eca7-119">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="5eca7-119">JSDoc Tags</span></span>

<span data-ttu-id="5eca7-120">As seguintes marcas JSDoc são suportadas em funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="5eca7-120">The following JSDoc tags are supported in Excel custom functions.</span></span>

* [<span data-ttu-id="5eca7-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="5eca7-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="5eca7-122">[@customfunction](#customfunction) nome de identificação</span><span class="sxs-lookup"><span data-stu-id="5eca7-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="5eca7-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="5eca7-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="5eca7-124">[@param](#param) _{type}_ nome e descrição</span><span class="sxs-lookup"><span data-stu-id="5eca7-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="5eca7-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="5eca7-125">@requiresAddress</span></span>](#requiresAddress)
* [<span data-ttu-id="5eca7-126">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="5eca7-126">@requiresParameterAddresses</span></span>](#requiresParameterAddresses)
* <span data-ttu-id="5eca7-127">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="5eca7-127">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="5eca7-128">@streaming</span><span class="sxs-lookup"><span data-stu-id="5eca7-128">@streaming</span></span>](#streaming)
* [<span data-ttu-id="5eca7-129">@volatile</span><span class="sxs-lookup"><span data-stu-id="5eca7-129">@volatile</span></span>](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a><span data-ttu-id="5eca7-130">@cancelable</span><span class="sxs-lookup"><span data-stu-id="5eca7-130">@cancelable</span></span>

<span data-ttu-id="5eca7-131">Indica que uma função personalizada executa uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-131">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="5eca7-132">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-132">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="5eca7-133">A função pode atribuir uma função à `oncanceled` propriedade para indicar o resultado quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-133">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="5eca7-134">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, ela será considerada `@cancelable`, mesmo se a tag não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="5eca7-134">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="5eca7-135">Uma função não pode ter as tags `@cancelable` e `@streaming` ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="5eca7-135">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

<a id="customfunction"></a>

### <a name="customfunction"></a><span data-ttu-id="5eca7-136">@customfunction</span><span class="sxs-lookup"><span data-stu-id="5eca7-136">@customfunction</span></span>

<span data-ttu-id="5eca7-137">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="5eca7-137">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="5eca7-138">Essa marca indica que a função JavaScript/TypeScript é uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="5eca7-138">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="5eca7-139">É necessário criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-139">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="5eca7-140">O exemplo a seguir mostra um exemplo dessa marca.</span><span class="sxs-lookup"><span data-stu-id="5eca7-140">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="5eca7-141">id</span><span class="sxs-lookup"><span data-stu-id="5eca7-141">id</span></span>

<span data-ttu-id="5eca7-142">O `id` identifica uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-142">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="5eca7-143">Se `id` não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="5eca7-143">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="5eca7-144">O `id` deve ser exclusivo para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5eca7-144">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="5eca7-145">Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="5eca7-145">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="5eca7-146">No exemplo a seguir, incremento é o `id` e o `name` da função.</span><span class="sxs-lookup"><span data-stu-id="5eca7-146">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="5eca7-147">nome</span><span class="sxs-lookup"><span data-stu-id="5eca7-147">name</span></span>

<span data-ttu-id="5eca7-148">Fornece a exibição `name` da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-148">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="5eca7-149">Se o nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="5eca7-149">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="5eca7-150">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="5eca7-150">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="5eca7-151">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="5eca7-151">Must start with a letter.</span></span>
* <span data-ttu-id="5eca7-152">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5eca7-152">Maximum length is 128 characters.</span></span>

<span data-ttu-id="5eca7-153">No exemplo a seguir, Inc é a `id`da função e `increment` é o `name`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-153">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="5eca7-154">descrição</span><span class="sxs-lookup"><span data-stu-id="5eca7-154">description</span></span>

<span data-ttu-id="5eca7-155">Uma descrição aparece para os usuários no Excel enquanto eles ingressam na função e especifica o que a função faz.</span><span class="sxs-lookup"><span data-stu-id="5eca7-155">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="5eca7-156">Uma descrição não exige nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="5eca7-156">A description doesn't require any specific tag.</span></span> <span data-ttu-id="5eca7-157">Adicione uma descrição a uma função personalizada acrescentando uma frase para descrever o que a função realiza dentro do comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="5eca7-157">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="5eca7-158">Por padrão, qualquer texto sem tags na seção de comentários JSDoc será a descrição da função.</span><span class="sxs-lookup"><span data-stu-id="5eca7-158">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="5eca7-159">No exemplo a seguir, a frase "Uma função que soma dois números" é a descrição da função personalizada com a propriedade id de `ADD`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-159">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>
### <a name="helpurl"></a><span data-ttu-id="5eca7-160">@helpurl</span><span class="sxs-lookup"><span data-stu-id="5eca7-160">@helpurl</span></span>

<span data-ttu-id="5eca7-161">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="5eca7-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="5eca7-162">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="5eca7-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="5eca7-163">No exemplo a seguir, `helpurl` o é `www.contoso.com/weatherhelp` .</span><span class="sxs-lookup"><span data-stu-id="5eca7-163">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>
### <a name="param"></a><span data-ttu-id="5eca7-164">@param</span><span class="sxs-lookup"><span data-stu-id="5eca7-164">@param</span></span>

#### <a name="javascript"></a><span data-ttu-id="5eca7-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="5eca7-165">JavaScript</span></span>

<span data-ttu-id="5eca7-166">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="5eca7-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="5eca7-167">`{type}` especifica as informações de tipo em chaves.</span><span class="sxs-lookup"><span data-stu-id="5eca7-167">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="5eca7-168">Confira a seção [Tipos](#types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="5eca7-168">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="5eca7-169">Se nenhum tipo for especificado, o tipo `any` padrão será usado.</span><span class="sxs-lookup"><span data-stu-id="5eca7-169">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="5eca7-170">`name` especifica o parâmetro ao @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="5eca7-170">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="5eca7-171">É necessário.</span><span class="sxs-lookup"><span data-stu-id="5eca7-171">It is required.</span></span>
* <span data-ttu-id="5eca7-172">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="5eca7-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="5eca7-173">É opcional.</span><span class="sxs-lookup"><span data-stu-id="5eca7-173">It is optional.</span></span>

<span data-ttu-id="5eca7-174">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="5eca7-174">To denote a custom function parameter as optional:</span></span>

* <span data-ttu-id="5eca7-175">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="5eca7-176">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="5eca7-177">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="5eca7-178">O exemplo a seguir mostra uma função ADD que adiciona dois ou três números, com o terceiro número como um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="5eca7-178">The following example shows an ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="5eca7-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="5eca7-179">TypeScript</span></span>

<span data-ttu-id="5eca7-180">Sintaxe de TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="5eca7-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="5eca7-181">`name` especifica o parâmetro ao @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="5eca7-181">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="5eca7-182">É necessário.</span><span class="sxs-lookup"><span data-stu-id="5eca7-182">It is required.</span></span>
* <span data-ttu-id="5eca7-183">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="5eca7-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="5eca7-184">É opcional.</span><span class="sxs-lookup"><span data-stu-id="5eca7-184">It is optional.</span></span>

<span data-ttu-id="5eca7-185">Confira a seção [Tipos](#types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="5eca7-185">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="5eca7-186">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="5eca7-186">To denote a custom function parameter as optional, do one of the following:</span></span>

* <span data-ttu-id="5eca7-187">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="5eca7-187">Use an optional parameter.</span></span> <span data-ttu-id="5eca7-188">Por exemplo: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="5eca7-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="5eca7-189">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="5eca7-189">Give the parameter a default value.</span></span> <span data-ttu-id="5eca7-190">Por exemplo: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="5eca7-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="5eca7-191">Para uma descrição detalhada do @param confira: [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="5eca7-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="5eca7-192">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="5eca7-193">O exemplo a seguir mostra `add` a função que adiciona dois números.</span><span class="sxs-lookup"><span data-stu-id="5eca7-193">The following example shows the `add` function that adds two numbers.</span></span>

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

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a><span data-ttu-id="5eca7-194">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="5eca7-194">@requiresAddress</span></span>

<span data-ttu-id="5eca7-195">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="5eca7-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="5eca7-196">O último parâmetro de função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado para usar `@requiresAddress` .</span><span class="sxs-lookup"><span data-stu-id="5eca7-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use `@requiresAddress`.</span></span> <span data-ttu-id="5eca7-197">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="5eca7-197">When the function is called, the `address` property will contain the address.</span></span>

<span data-ttu-id="5eca7-198">O exemplo a seguir mostra como usar o parâmetro em combinação com para retornar o endereço da célula que `invocation` `@requiresAddress` invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-198">The following sample shows how to use the `invocation` parameter in combination with `@requiresAddress` to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="5eca7-199">Consulte [o parâmetro Invocation](custom-functions-parameter-options.md#invocation-parameter) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="5eca7-199">See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) for more information.</span></span>

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a><span data-ttu-id="5eca7-200">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="5eca7-200">@requiresParameterAddresses</span></span>

<span data-ttu-id="5eca7-201">Indica que a função deve retornar os endereços dos parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-201">Indicates that the function should return the addresses of input parameters.</span></span> 

<span data-ttu-id="5eca7-202">O último parâmetro de função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado para usar  `@requiresParameterAddresses` .</span><span class="sxs-lookup"><span data-stu-id="5eca7-202">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use  `@requiresParameterAddresses`.</span></span> <span data-ttu-id="5eca7-203">O comentário JSDoc também deve incluir uma marca especificando que o valor de retorno seja uma `@returns` matriz, como `@returns {string[][]}` ou `@returns {number[][]}` .</span><span class="sxs-lookup"><span data-stu-id="5eca7-203">The JSDoc comment must also include an `@returns` tag specifying that the return value be a matrix, such as `@returns {string[][]}` or `@returns {number[][]}`.</span></span> <span data-ttu-id="5eca7-204">Consulte [Tipos de matriz](#matrix-type) para obter informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="5eca7-204">See [Matrix types](#matrix-type) for additional information.</span></span> 

<span data-ttu-id="5eca7-205">Quando a função for chamada, a `parameterAddresses` propriedade conterá os endereços dos parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-205">When the function is called, the `parameterAddresses` property will contain the addresses of the input parameters.</span></span>

<span data-ttu-id="5eca7-206">O exemplo a seguir mostra como usar o parâmetro em combinação com para `invocation` `@requiresParameterAddresses` retornar os endereços de três parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-206">The following sample shows how to use the `invocation` parameter in combination with `@requiresParameterAddresses` to return the addresses of three input parameters.</span></span> <span data-ttu-id="5eca7-207">Consulte [Detectar o endereço de um parâmetro para](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="5eca7-207">See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.</span></span> 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>
### <a name="returns"></a><span data-ttu-id="5eca7-208">@returns</span><span class="sxs-lookup"><span data-stu-id="5eca7-208">@returns</span></span>

<span data-ttu-id="5eca7-209">Sintaxe: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="5eca7-209">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="5eca7-210">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="5eca7-210">Provides the type for the return value.</span></span>

<span data-ttu-id="5eca7-211">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="5eca7-211">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="5eca7-212">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-212">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="5eca7-213">O exemplo a seguir mostra a `add` função que usa `@returns` marca.</span><span class="sxs-lookup"><span data-stu-id="5eca7-213">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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

<a id="streaming"></a>
### <a name="streaming"></a><span data-ttu-id="5eca7-214">@streaming</span><span class="sxs-lookup"><span data-stu-id="5eca7-214">@streaming</span></span>

<span data-ttu-id="5eca7-215">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="5eca7-215">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="5eca7-216">O último parâmetro é do tipo `CustomFunctions.StreamingInvocation<ResultType>` .</span><span class="sxs-lookup"><span data-stu-id="5eca7-216">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="5eca7-217">A função retorna `void` .</span><span class="sxs-lookup"><span data-stu-id="5eca7-217">The function returns `void`.</span></span>

<span data-ttu-id="5eca7-218">As funções de streaming não retornam valores diretamente, em vez disso, elas chamam `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-218">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="5eca7-219">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="5eca7-219">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="5eca7-220">`setResult()` pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-220">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="5eca7-221">Para obter um exemplo de uma função de streaming e mais informações, confira [, criar uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="5eca7-221">For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="5eca7-222">As funções de streaming não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="5eca7-222">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

<a id="volatile"></a>
### <a name="volatile"></a><span data-ttu-id="5eca7-223">@volatile</span><span class="sxs-lookup"><span data-stu-id="5eca7-223">@volatile</span></span>

<span data-ttu-id="5eca7-224">Uma função volátil é aquela cujo resultado não é o mesmo de um momento para o outro, mesmo que não receba argumentos ou os argumentos não mudem.</span><span class="sxs-lookup"><span data-stu-id="5eca7-224">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="5eca7-225">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="5eca7-225">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="5eca7-226">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="5eca7-226">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="5eca7-227">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="5eca7-227">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="5eca7-228">A função a seguir é volátil e usa `@volatile` a marca.</span><span class="sxs-lookup"><span data-stu-id="5eca7-228">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="5eca7-229">Tipos</span><span class="sxs-lookup"><span data-stu-id="5eca7-229">Types</span></span>

<span data-ttu-id="5eca7-230">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="5eca7-230">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="5eca7-231">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="5eca7-231">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="5eca7-232">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="5eca7-232">Value types</span></span>

<span data-ttu-id="5eca7-233">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="5eca7-233">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="5eca7-234">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="5eca7-234">Matrix type</span></span>

<span data-ttu-id="5eca7-235">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="5eca7-235">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="5eca7-236">Por exemplo, o tipo `number[][]` indica uma matriz de números e indica uma matriz de `string[][]` cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="5eca7-236">For example, the type `number[][]` indicates a matrix of numbers and `string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="5eca7-237">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="5eca7-237">Error type</span></span>

<span data-ttu-id="5eca7-238">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-238">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="5eca7-239">Uma função de streaming pode indicar um erro chamando `setResult()` com um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-239">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="5eca7-240">Promessa</span><span class="sxs-lookup"><span data-stu-id="5eca7-240">Promise</span></span>

<span data-ttu-id="5eca7-241">Uma função personalizada pode retornar uma promessa que fornece o valor quando a promessa é resolvida.</span><span class="sxs-lookup"><span data-stu-id="5eca7-241">A custom function can return a promise that provides the value when the promise is resolved.</span></span> <span data-ttu-id="5eca7-242">Se a promessa for rejeitada, a função personalizada lançará um erro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-242">If the promise is rejected, then the custom function will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="5eca7-243">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="5eca7-243">Other types</span></span>

<span data-ttu-id="5eca7-244">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="5eca7-244">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="5eca7-245">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="5eca7-245">Next steps</span></span>

<span data-ttu-id="5eca7-246">Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="5eca7-246">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="5eca7-247">Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="5eca7-247">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="5eca7-248">Confira também</span><span class="sxs-lookup"><span data-stu-id="5eca7-248">See also</span></span>

* [<span data-ttu-id="5eca7-249">Criar metadados JSON manualmente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eca7-249">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="5eca7-250">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="5eca7-250">Create custom functions in Excel</span></span>](custom-functions-overview.md)
