---
ms.date: 09/21/2020
description: 'Manipular e retornar erros como #NULL! a partir de sua função personalizada.'
title: Manipular e retornar erros de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175525"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="278f1-104">Manipular e retornar erros de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="278f1-104">Handle and return errors from your custom function</span></span>

<span data-ttu-id="278f1-105">Se algo der errado enquanto sua função personalizada é executada, retorne um erro para informar ao usuário.</span><span class="sxs-lookup"><span data-stu-id="278f1-105">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="278f1-106">Se você tiver requisitos de parâmetros específicos, como apenas números positivos, teste os parâmetros e acione um erro se eles não estiverem corretos.</span><span class="sxs-lookup"><span data-stu-id="278f1-106">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="278f1-107">Você também pode usar um bloco `try`-`catch` para detectar quaisquer erros que ocorram enquanto sua função personalizada é executada.</span><span class="sxs-lookup"><span data-stu-id="278f1-107">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="278f1-108">Detectar e lançar um erro</span><span class="sxs-lookup"><span data-stu-id="278f1-108">Detect and throw an error</span></span>

<span data-ttu-id="278f1-109">Vamos examinar um caso em que você precisa garantir que um parâmetro de CEP esteja no formato correto para que a função personalizada funcione.</span><span class="sxs-lookup"><span data-stu-id="278f1-109">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="278f1-110">A função personalizada a seguir usa uma expressão regular para verificar o CEP.</span><span class="sxs-lookup"><span data-stu-id="278f1-110">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="278f1-111">Se o formato do CEP estiver correto, ele pesquisará a cidade usando outra função e retornará o valor.</span><span class="sxs-lookup"><span data-stu-id="278f1-111">If the zip code format is correct, then it will look up the city using another function and return the value.</span></span> <span data-ttu-id="278f1-112">Se o formato não for válido, a função retornará um `#VALUE!` erro à célula.</span><span class="sxs-lookup"><span data-stu-id="278f1-112">If the format isn't valid, the function returns a `#VALUE!` error to the cell.</span></span>

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="278f1-113">O objeto CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="278f1-113">The CustomFunctions.Error object</span></span>

<span data-ttu-id="278f1-114">O objeto [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) é usado para retornar um erro de volta para a célula.</span><span class="sxs-lookup"><span data-stu-id="278f1-114">The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object is used to return an error back to the cell.</span></span> <span data-ttu-id="278f1-115">Ao criar o objeto, especifique o erro que você deseja usar, escolhendo um dos seguintes valores de `ErrorCode` enumeração.</span><span class="sxs-lookup"><span data-stu-id="278f1-115">When you create the object, specify which error you want to use by choosing one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="278f1-116">Valor de enumeração ErrorCode</span><span class="sxs-lookup"><span data-stu-id="278f1-116">ErrorCode enum value</span></span>  |<span data-ttu-id="278f1-117">Valor da célula do Excel</span><span class="sxs-lookup"><span data-stu-id="278f1-117">Excel cell value</span></span>  |<span data-ttu-id="278f1-118">Significado</span><span class="sxs-lookup"><span data-stu-id="278f1-118">Meaning</span></span>  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="278f1-119">Esteja ciente de que o JavaScript permite a divisão por zero, portanto, você precisa escrever um manipulador de erros com cuidado para detectar essa condição.</span><span class="sxs-lookup"><span data-stu-id="278f1-119">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidName`    | `#NAME?`  | <span data-ttu-id="278f1-120">Há um erro de digitação no nome da função.</span><span class="sxs-lookup"><span data-stu-id="278f1-120">There is a typo in the function name.</span></span> <span data-ttu-id="278f1-121">Observe que esse erro é suportado como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada.</span><span class="sxs-lookup"><span data-stu-id="278f1-121">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span> | 
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="278f1-122">Há um problema com um número na fórmula.</span><span class="sxs-lookup"><span data-stu-id="278f1-122">There is a problem with a number in the formula.</span></span> |
|`invalidReference` | `#REF!` | <span data-ttu-id="278f1-123">A função se refere a uma célula inválida.</span><span class="sxs-lookup"><span data-stu-id="278f1-123">The function refers to an invalid cell.</span></span> <span data-ttu-id="278f1-124">Observe que esse erro é suportado como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada.</span><span class="sxs-lookup"><span data-stu-id="278f1-124">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span>|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="278f1-125">Um valor na fórmula é do tipo incorreto.</span><span class="sxs-lookup"><span data-stu-id="278f1-125">A value in the formula is of the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="278f1-126">A função ou o serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="278f1-126">The function or service isn't available.</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="278f1-127">Os intervalos na fórmula não fazem interseção.</span><span class="sxs-lookup"><span data-stu-id="278f1-127">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="278f1-128">O exemplo de código a seguir mostra como criar e retornar um erro para um número inválido (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="278f1-128">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="278f1-129">O `#VALUE!` e os `#N/A` erros também dão suporte a mensagens de erro personalizadas.</span><span class="sxs-lookup"><span data-stu-id="278f1-129">The `#VALUE!` and `#N/A` errors also support custom error messages.</span></span> <span data-ttu-id="278f1-130">As mensagens de erro personalizadas são exibidas no menu indicador de erro, que é acessado ao passar o sinalizador de erro em cada célula com um erro.</span><span class="sxs-lookup"><span data-stu-id="278f1-130">Custom error messages are displayed in the error indicator menu, which is accessed by hovering over the error flag on each cell with an error.</span></span> <span data-ttu-id="278f1-131">O exemplo a seguir mostra como retornar uma mensagem de erro personalizada com o `#VALUE!` erro.</span><span class="sxs-lookup"><span data-stu-id="278f1-131">The following example shows how to return a custom error message with the `#VALUE!` error.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="278f1-132">Use blocos try-catch</span><span class="sxs-lookup"><span data-stu-id="278f1-132">Use try-catch blocks</span></span>

<span data-ttu-id="278f1-133">Em geral, use `try` - `catch` blocos em sua função personalizada para detectar quaisquer possíveis erros que ocorram.</span><span class="sxs-lookup"><span data-stu-id="278f1-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="278f1-134">Se você não tratar exceções no seu código, elas serão retornadas ao Excel.</span><span class="sxs-lookup"><span data-stu-id="278f1-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="278f1-135">Por padrão, o Excel retorna `#VALUE!` para erros ou exceções não manipuladas.</span><span class="sxs-lookup"><span data-stu-id="278f1-135">By default, Excel returns `#VALUE!` for unhandled errors or exceptions.</span></span>

<span data-ttu-id="278f1-136">No exemplo de código a seguir, a função personalizada faz uma chamada de busca para um serviço REST.</span><span class="sxs-lookup"><span data-stu-id="278f1-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="278f1-137">É possível que a chamada falhe, por exemplo, se o serviço REST retornar um erro ou a rede cair.</span><span class="sxs-lookup"><span data-stu-id="278f1-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="278f1-138">Se isso acontecer, a função personalizada retornará `#N/A` para indicar que a chamada da Web falhou.</span><span class="sxs-lookup"><span data-stu-id="278f1-138">If this happens, the custom function will return `#N/A` to indicate that the web call failed.</span></span>


```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="278f1-139">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="278f1-139">Next steps</span></span>

<span data-ttu-id="278f1-140">Saiba como [solucionar problemas com as suas funções personalizadas](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="278f1-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="278f1-141">Confira também</span><span class="sxs-lookup"><span data-stu-id="278f1-141">See also</span></span>

* [<span data-ttu-id="278f1-142">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="278f1-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="278f1-143">Requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="278f1-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="278f1-144">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="278f1-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
