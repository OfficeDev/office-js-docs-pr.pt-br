---
ms.date: 05/06/2020
description: 'Manipular e retornar erros como #NULL! da sua função personalizada'
title: Manipular e retornar erros da sua função personalizada (visualização)
localization_priority: Normal
ms.openlocfilehash: 5b1efcdc22a4efc59304bbe76f8d3f2d09979bc1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093466"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a><span data-ttu-id="89475-104">Manipular e retornar erros da sua função personalizada (visualização)</span><span class="sxs-lookup"><span data-stu-id="89475-104">Handle and return errors from your custom function (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="89475-105">Os recursos descritos neste artigo estão atualmente em visualização, estando sujeitos a alterações.</span><span class="sxs-lookup"><span data-stu-id="89475-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="89475-106">No momento, eles não têm suporte para utilização em ambientes de produção.</span><span class="sxs-lookup"><span data-stu-id="89475-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="89475-107">Você precisará participar do programa [Office Insider](https://insider.office.com/join) para experimentar os recursos de visualização.</span><span class="sxs-lookup"><span data-stu-id="89475-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="89475-108">Uma boa maneira de experimentar recursos de visualização é usar uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="89475-108">A good way to try out preview features is by using a Microsoft 365 subscription.</span></span> <span data-ttu-id="89475-109">Se você ainda não tem uma assinatura do Microsoft 365, é possível obter uma assinatura gratuita, de 90 dias, redimensionada pela Microsoft 365, participando do [programa de desenvolvedor do microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="89475-109">If you don't already have a Microsoft 365 subscription, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="89475-110">Se algo der errado enquanto sua função personalizada é executada, retorne um erro para informar ao usuário.</span><span class="sxs-lookup"><span data-stu-id="89475-110">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="89475-111">Se você tiver requisitos de parâmetros específicos, como apenas números positivos, teste os parâmetros e acione um erro se eles não estiverem corretos.</span><span class="sxs-lookup"><span data-stu-id="89475-111">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="89475-112">Você também pode usar um bloco `try`-`catch` para detectar quaisquer erros que ocorram enquanto sua função personalizada é executada.</span><span class="sxs-lookup"><span data-stu-id="89475-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="89475-113">Detectar e lançar um erro</span><span class="sxs-lookup"><span data-stu-id="89475-113">Detect and throw an error</span></span>

<span data-ttu-id="89475-114">Vamos examinar um caso em que você precisa garantir que um parâmetro de CEP esteja no formato correto para que a função personalizada funcione.</span><span class="sxs-lookup"><span data-stu-id="89475-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="89475-115">A função personalizada a seguir usa uma expressão regular para verificar o CEP.</span><span class="sxs-lookup"><span data-stu-id="89475-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="89475-116">Se estiver correto, ele pesquisará a cidade usando outra função e retornará o valor.</span><span class="sxs-lookup"><span data-stu-id="89475-116">If it is correct, then it will look up the city using another function, and return the value.</span></span> <span data-ttu-id="89475-117">Se ele não estiver correto, retornará um `#VALUE!` erro à célula.</span><span class="sxs-lookup"><span data-stu-id="89475-117">If it isn't correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="89475-118">O objeto CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="89475-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="89475-119">O objeto `CustomFunctions.Error` é usado para retornar um erro de volta à célula.</span><span class="sxs-lookup"><span data-stu-id="89475-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="89475-120">Ao criar o objeto, especifique qual erro você deseja usar usando um dos seguintes valores de enumeração `ErrorCode`.</span><span class="sxs-lookup"><span data-stu-id="89475-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="89475-121">Valor de enumeração ErrorCode</span><span class="sxs-lookup"><span data-stu-id="89475-121">ErrorCode enum value</span></span>  |<span data-ttu-id="89475-122">Valor da célula do Excel</span><span class="sxs-lookup"><span data-stu-id="89475-122">Excel cell value</span></span>  |<span data-ttu-id="89475-123">Significado</span><span class="sxs-lookup"><span data-stu-id="89475-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="89475-124">Um valor usado na fórmula é de tipo incorreto.</span><span class="sxs-lookup"><span data-stu-id="89475-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="89475-125">A função ou o serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="89475-125">The function or service isn't available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="89475-126">Esteja ciente de que o JavaScript permite a divisão por zero, portanto, você precisa escrever um manipulador de erros com cuidado para detectar essa condição.</span><span class="sxs-lookup"><span data-stu-id="89475-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="89475-127">Há um problema com o número usado na fórmula</span><span class="sxs-lookup"><span data-stu-id="89475-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="89475-128">Os intervalos na fórmula não fazem interseção.</span><span class="sxs-lookup"><span data-stu-id="89475-128">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="89475-129">O exemplo de código a seguir mostra como criar e retornar um erro para um número inválido (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="89475-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="89475-130">Quando você retorna um erro `#VALUE!`, também pode incluir uma mensagem personalizada que será mostrada em um pop-up quando o usuário passar o mouse sobre a célula.</span><span class="sxs-lookup"><span data-stu-id="89475-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="89475-131">O exemplo a seguir mostra como retornar uma mensagem de erro personalizada.</span><span class="sxs-lookup"><span data-stu-id="89475-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="89475-132">Use blocos try-catch</span><span class="sxs-lookup"><span data-stu-id="89475-132">Use try-catch blocks</span></span>

<span data-ttu-id="89475-133">Em geral, use `try` - `catch` blocos em sua função personalizada para detectar quaisquer possíveis erros que ocorram.</span><span class="sxs-lookup"><span data-stu-id="89475-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="89475-134">Se você não tratar exceções no seu código, elas serão retornadas ao Excel.</span><span class="sxs-lookup"><span data-stu-id="89475-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="89475-135">Por padrão, o Excel retorna `#VALUE!` para uma exceção não tratada.</span><span class="sxs-lookup"><span data-stu-id="89475-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="89475-136">No exemplo de código a seguir, a função personalizada faz uma chamada de busca para um serviço REST.</span><span class="sxs-lookup"><span data-stu-id="89475-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="89475-137">É possível que a chamada falhe, por exemplo, se o serviço REST retornar um erro ou a rede cair.</span><span class="sxs-lookup"><span data-stu-id="89475-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="89475-138">Se isso acontecer, a função personalizada retornará `#N/A` para indicar que a chamada Web falhou.</span><span class="sxs-lookup"><span data-stu-id="89475-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="89475-139">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="89475-139">Next steps</span></span>

<span data-ttu-id="89475-140">Saiba como [solucionar problemas com as suas funções personalizadas](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="89475-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="89475-141">Confira também</span><span class="sxs-lookup"><span data-stu-id="89475-141">See also</span></span>

* [<span data-ttu-id="89475-142">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="89475-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="89475-143">Requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="89475-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="89475-144">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="89475-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
