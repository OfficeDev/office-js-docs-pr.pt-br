---
ms.date: 06/21/2019
description: Use tags JSDoc para criar dinamicamente seus metadados JSON de funções personalizadas.
title: Gerar metadados JSON automaticamente para funções personalizadas
localization_priority: Priority
ms.openlocfilehash: cc28eca4e1ab1a03186983c81380a00bcf5eb85a
ms.sourcegitcommit: 6d1cb188c76c09d320025abfcc99db1b16b7e37b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2019
ms.locfileid: "35226759"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="0ffd0-103">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0ffd0-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="0ffd0-104">Quando uma função personalizada do Excel é gravada em JavaScript ou em TypeScript, as marcações JSDoc são usadas para fornecer informações adicionais sobre a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="0ffd0-105">As marcações JSDoc são usadas no momento da criação para criar o [arquivo de metadados JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="0ffd0-106">O uso de marcações JSDoc poupa você do esforço de editar manualmente o arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0ffd0-107">Adicione a marcação `@customfunction` nos comentários de código de uma função JavaScript ou TypeScript para marcá-la como uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="0ffd0-108">Os tipos de parâmetros da função podem ser fornecidos usando a marcação [@param](#param) em JavaScript ou do [Tipo de função](https://www.typescriptlang.org/docs/handbook/functions.html) em TypeScript.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="0ffd0-109">Para mais informações, confira a marcação [@param](#param) e a seção [Tipos](#types).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="0ffd0-110">Adicionando uma descrição a uma função</span><span class="sxs-lookup"><span data-stu-id="0ffd0-110">Adding a description to a function</span></span>

<span data-ttu-id="0ffd0-111">A descrição é exibida para o usuário como texto de ajuda quando eles precisam de ajuda para entender o que a função personalizada executa.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="0ffd0-112">A descrição não requer nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="0ffd0-113">Basta digitar uma breve descrição de texto no comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="0ffd0-114">Em geral, a descrição é colocada no início da seção de comentários do JSDoc, mas funcionará independentemente de onde seja colocada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="0ffd0-115">Para ver exemplos das descrições de funções internas, abra o Excel, vá para a guia **Fórmulas** e escolha **Inserir função**.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="0ffd0-116">Você pode navegar por todas as descrições de funções e também ver suas próprias funções personalizadas listadas.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="0ffd0-117">No exemplo a seguir, a frase "Calcula o volume de uma esfera."</span><span class="sxs-lookup"><span data-stu-id="0ffd0-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="0ffd0-118">é a descrição da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="0ffd0-119">Marcações JSDoc</span><span class="sxs-lookup"><span data-stu-id="0ffd0-119">JSDoc Tags</span></span>
<span data-ttu-id="0ffd0-120">As seguintes marcações JSDoc possuem suporte em funções personalizadas do Excel:</span><span class="sxs-lookup"><span data-stu-id="0ffd0-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="0ffd0-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="0ffd0-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="0ffd0-122">[@customfunction](#customfunction) nome de identificação</span><span class="sxs-lookup"><span data-stu-id="0ffd0-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="0ffd0-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="0ffd0-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="0ffd0-124">[@param](#param) _{type}_ nome e descrição</span><span class="sxs-lookup"><span data-stu-id="0ffd0-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="0ffd0-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="0ffd0-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="0ffd0-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="0ffd0-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="0ffd0-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="0ffd0-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="0ffd0-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="0ffd0-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="0ffd0-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="0ffd0-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="0ffd0-130">Indica que uma função personalizada deseja executar uma ação quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="0ffd0-131">O último parâmetro da função deve ser do tipo `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="0ffd0-132">A função pode atribuir uma função à propriedade `oncanceled` para denotar a ação a ser executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="0ffd0-133">Se o último parâmetro da função for do tipo `CustomFunctions.CancelableInvocation`, ela será considerada `@cancelable`, mesmo se a tag não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="0ffd0-134">Uma função não pode ter as tags `@cancelable` e `@streaming` ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-134">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="0ffd0-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="0ffd0-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="0ffd0-136">Sintaxe: @customfunction _id_ _nome_</span><span class="sxs-lookup"><span data-stu-id="0ffd0-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="0ffd0-137">Especifique esta marcação para tratar a função JavaScript/TypeScript como uma função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span> 

<span data-ttu-id="0ffd0-138">Essa marcação é necessária para criar metadados para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="0ffd0-139">Também deve haver uma chamada para `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="0ffd0-139">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

<span data-ttu-id="0ffd0-140">O exemplo a seguir mostra a maneira mais simples de declarar uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-140">The following example shows the simplest way to declare a custom function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="0ffd0-141">id</span><span class="sxs-lookup"><span data-stu-id="0ffd0-141">id</span></span>

<span data-ttu-id="0ffd0-142">O `id` é um identificador invariável para a função customizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-142">The id is used as the invariant identifier for the custom function stored in the document.</span></span>

* <span data-ttu-id="0ffd0-143">Se `id` não for fornecido, o nome da função JavaScript/TypeScript será convertido em maiúsculas e os caracteres não permitidos serão removidos.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-143">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="0ffd0-144">O `id` deve ser exclusivo para todas as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-144">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="0ffd0-145">Os caracteres permitidos estão limitados a: A-Z, a-z, 0-9, sublinhados (\_) e ponto (.).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-145">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

<span data-ttu-id="0ffd0-146">No exemplo a seguir, incremento é o `id` e o `name` da função.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-146">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="0ffd0-147">nome</span><span class="sxs-lookup"><span data-stu-id="0ffd0-147">name</span></span>

<span data-ttu-id="0ffd0-148">Fornece a exibição `name` da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-148">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="0ffd0-149">Se o nome não for fornecido, o id também será usado como nome.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-149">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="0ffd0-150">Caracteres permitidos: Letras de [caractere Alfabético Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), números, ponto (.) e sublinhado (\_).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-150">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="0ffd0-151">Deve começar com uma letra.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-151">Must start with a letter.</span></span>
* <span data-ttu-id="0ffd0-152">O comprimento máximo é de 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-152">Maximum length is 128 characters.</span></span>

<span data-ttu-id="0ffd0-153">No exemplo a seguir, Inc é a `id`da função e `increment` é o `name`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-153">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="0ffd0-154">descrição</span><span class="sxs-lookup"><span data-stu-id="0ffd0-154">description</span></span>

<span data-ttu-id="0ffd0-155">Uma descrição não exige nenhuma tag específica.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-155">A description doesn't require any specific tag.</span></span> <span data-ttu-id="0ffd0-156">Adicione uma descrição a uma função personalizada acrescentando uma frase para descrever o que a função realiza dentro do comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-156">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="0ffd0-157">Por padrão, qualquer texto sem tags na seção de comentários JSDoc será a descrição da função.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-157">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="0ffd0-158">A descrição aparece para os usuários no Excel quando eles entram na função.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-158">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="0ffd0-159">No exemplo a seguir, a frase "Uma função que soma dois números" é a descrição da função personalizada com a propriedade id de `ADD`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-159">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

<span data-ttu-id="0ffd0-160">No exemplo a seguir, adicionar é o `id` e `name` da função e uma descrição é fornecida.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-160">In the following example, ADD is the `id` and `name` of the function and a description is given.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="0ffd0-161">@helpurl</span><span class="sxs-lookup"><span data-stu-id="0ffd0-161">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="0ffd0-162">Sintaxe: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="0ffd0-162">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="0ffd0-163">A _url_ fornecida é exibida no Excel.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-163">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="0ffd0-164">No exemplo a seguir, o `helpurl` é www.contoso.com/weatherhelp.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-164">In the following example, the `helpurl` is www.contoso.com/weatherhelp.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a><span data-ttu-id="0ffd0-165">@param</span><span class="sxs-lookup"><span data-stu-id="0ffd0-165">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="0ffd0-166">JavaScript</span><span class="sxs-lookup"><span data-stu-id="0ffd0-166">JavaScript</span></span>

<span data-ttu-id="0ffd0-167">Sintaxe de JavaScript: @param {type} nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="0ffd0-167">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="0ffd0-168">`{type}` deve especificar a informação de tipo entre chaves.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-168">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="0ffd0-169">Confira [Tipos](##types) para mais informações sobre os tipos que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-169">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="0ffd0-170">Opcional: se não especificado, o tipo `any` será usado.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-170">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="0ffd0-171">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-171">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="0ffd0-172">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-172">Required.</span></span>
* <span data-ttu-id="0ffd0-173">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-173">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="0ffd0-174">Opcional.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-174">Optional.</span></span>

<span data-ttu-id="0ffd0-175">Para denotar um parâmetro de função personalizado como opcional:</span><span class="sxs-lookup"><span data-stu-id="0ffd0-175">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="0ffd0-176">Coloque colchetes ao redor do nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-176">Put square brackets around the parameter name.</span></span> <span data-ttu-id="0ffd0-177">Por exemplo: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-177">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="0ffd0-178">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-178">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="0ffd0-179">O exemplo a seguir mostra uma função adicionar que adiciona dois ou três números, com o terceiro número como um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-179">The following example shows a ADD function which adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="0ffd0-180">TypeScript</span><span class="sxs-lookup"><span data-stu-id="0ffd0-180">TypeScript</span></span>

<span data-ttu-id="0ffd0-181">Sintaxe de TypeScript: @param nome _descrição_</span><span class="sxs-lookup"><span data-stu-id="0ffd0-181">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="0ffd0-182">`name` especifica a qual parâmetro a marcação @param se aplica.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-182">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="0ffd0-183">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-183">Required.</span></span>
* <span data-ttu-id="0ffd0-184">`description` fornece a descrição que aparece no Excel para o parâmetro de função.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-184">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="0ffd0-185">Opcional.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-185">Optional.</span></span>

<span data-ttu-id="0ffd0-186">Confira [Tipos](##types) para mais informações sobre os tipos de parâmetros de função que podem ser usados.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-186">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="0ffd0-187">Para denotar um parâmetro de função personalizado como opcional, siga um destes procedimentos:</span><span class="sxs-lookup"><span data-stu-id="0ffd0-187">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="0ffd0-188">Use um parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-188">Use an optional parameter.</span></span> <span data-ttu-id="0ffd0-189">Por exemplo: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="0ffd0-189">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="0ffd0-190">Dê ao parâmetro um valor padrão.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-190">Give the parameter a default value.</span></span> <span data-ttu-id="0ffd0-191">Por exemplo: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="0ffd0-191">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="0ffd0-192">Para uma descrição detalhada do @param confira: [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="0ffd0-192">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="0ffd0-193">O valor padrão para parâmetros opcionais é `null`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-193">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="0ffd0-194">O exemplo a seguir mostra `add` a função que adiciona dois números.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-194">The following example shows the `add` function that adds two numbers.</span></span>

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
### <a name="requiresaddress"></a><span data-ttu-id="0ffd0-195">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="0ffd0-195">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="0ffd0-196">Indica que o endereço da célula onde a função está sendo avaliada deve ser fornecido.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-196">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="0ffd0-197">O último parâmetro da função deve ser do tipo `CustomFunctions.Invocation` ou de um tipo derivado.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-197">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="0ffd0-198">Quando a função é chamada, a propriedade `address` conterá o endereço.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-198">When the function is called, the `address` property will contain the address.</span></span> <span data-ttu-id="0ffd0-199">Para obter um exemplo de uma função que usa `@requiresAddress` a marca, [Confira o parâmetro contexto](./custom-functions-parameter-options.md#addressing-cells-context-parameter)da célula de endereçamento.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-199">For an example of a function that uses the `@requiresAddress` tag, see [Addressing cell's context parameter](./custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span>

---
### <a name="returns"></a><span data-ttu-id="0ffd0-200">@returns</span><span class="sxs-lookup"><span data-stu-id="0ffd0-200">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="0ffd0-201">Sintaxe: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="0ffd0-201">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="0ffd0-202">Fornece o tipo para o valor de retorno.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-202">Provides the type for the return value.</span></span>

<span data-ttu-id="0ffd0-203">Se `{type}` for omitido, as informações do tipo TypeScript serão usadas.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-203">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="0ffd0-204">Se não houver informações de tipo, o tipo será `any`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-204">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="0ffd0-205">O exemplo a seguir mostra a `add` função que usa `@returns` marca.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-205">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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
### <a name="streaming"></a><span data-ttu-id="0ffd0-206">@streaming</span><span class="sxs-lookup"><span data-stu-id="0ffd0-206">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="0ffd0-207">Usado para indicar que uma função personalizada é uma função de streaming.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-207">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="0ffd0-208">O último parâmetro deve ser do tipo `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-208">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="0ffd0-209">A função deve retornar `void`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-209">The function should return `void`.</span></span>

<span data-ttu-id="0ffd0-210">As funções de streaming não retornam valores diretamente, mas devem chamar `setResult(result: ResultType)` usando o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-210">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="0ffd0-211">Exceções lançadas por uma função de streaming são ignoradas.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-211">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="0ffd0-212">`setResult()` pode ser chamado com Erro para indicar um resultado de erro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-212">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="0ffd0-213">Para obter um exemplo de uma função de streaming e mais informações, confira [, criar uma função de streaming](./custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-213">For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="0ffd0-214">As funções de streaming não podem ser marcadas como [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-214">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="0ffd0-215">@volatile</span><span class="sxs-lookup"><span data-stu-id="0ffd0-215">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="0ffd0-216">Uma função volátil é aquela cujo resultado não é o mesmo de um momento para o outro, mesmo que não receba argumentos ou os argumentos não mudem.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-216">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="0ffd0-217">O Excel reavalia células que contenham funções voláteis, juntamente com todos os dependentes, sempre que um cálculo é feito.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-217">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="0ffd0-218">Por esse motivo, confiar demais em funções voláteis pode retardar o tempo de recálculo; portanto, use-as com moderação.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-218">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="0ffd0-219">Funções de streaming não podem ser voláteis.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-219">Streaming functions cannot be volatile.</span></span>

<span data-ttu-id="0ffd0-220">A função a seguir é volátil e usa `@volatile` a marca.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-220">The following function is volatile and uses the `@volatile` tag.</span></span>

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

## <a name="types"></a><span data-ttu-id="0ffd0-221">Tipos</span><span class="sxs-lookup"><span data-stu-id="0ffd0-221">Types</span></span>

<span data-ttu-id="0ffd0-222">Especificando um tipo de parâmetro, o Excel converterá valores nesse tipo antes de chamar a função.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-222">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="0ffd0-223">Se o tipo for `any`, nenhuma conversão será executada.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-223">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="0ffd0-224">Tipos de valor</span><span class="sxs-lookup"><span data-stu-id="0ffd0-224">Value types</span></span>

<span data-ttu-id="0ffd0-225">Um valor pode ser representado usando um dos seguintes tipos: `boolean``number``string`.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-225">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="0ffd0-226">Tipo de matriz</span><span class="sxs-lookup"><span data-stu-id="0ffd0-226">Matrix type</span></span>

<span data-ttu-id="0ffd0-227">Use um tipo de matriz bidimensional para que o parâmetro ou valor de retorno seja uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-227">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="0ffd0-228">Por exemplo, o tipo `number[][]` indica uma matriz de números.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-228">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="0ffd0-229">`string[][]` indica uma matriz de cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-229">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="0ffd0-230">Tipo de erro</span><span class="sxs-lookup"><span data-stu-id="0ffd0-230">Error type</span></span>

<span data-ttu-id="0ffd0-231">Uma função que não seja de streaming pode indicar um erro retornando um tipo de Erro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-231">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="0ffd0-232">Uma função de streaming pode indicar um erro chamando `setResult()` com um tipo de erro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-232">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="0ffd0-233">Promessa</span><span class="sxs-lookup"><span data-stu-id="0ffd0-233">Promise</span></span>

<span data-ttu-id="0ffd0-234">Uma função pode retornar uma Promessa, que fornecerá o valor quando a promessa for resolvida.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-234">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="0ffd0-235">Se a promessa for rejeitada, então é um erro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-235">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="0ffd0-236">Outros tipos</span><span class="sxs-lookup"><span data-stu-id="0ffd0-236">Other types</span></span>

<span data-ttu-id="0ffd0-237">Qualquer outro tipo será tratado como um erro.</span><span class="sxs-lookup"><span data-stu-id="0ffd0-237">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0ffd0-238">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="0ffd0-238">Next steps</span></span>
<span data-ttu-id="0ffd0-239">Saiba mais sobre [convenções de nomenclatura para funções personalizadas](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-239">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="0ffd0-240">Como alternativa, saiba como [localizar as funções](custom-functions-localize.md) que requerem a [gravação do arquivo JSON à mão](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-240">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0ffd0-241">Confira também</span><span class="sxs-lookup"><span data-stu-id="0ffd0-241">See also</span></span>

* [<span data-ttu-id="0ffd0-242">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0ffd0-242">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="0ffd0-243">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="0ffd0-243">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="0ffd0-244">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="0ffd0-244">Create custom functions in Excel</span></span>](custom-functions-overview.md)
