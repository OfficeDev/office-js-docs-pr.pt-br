---
ms.date: 12/21/2020
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: 312046551236e96e67de6f63f3e3511aba6f50ce
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735526"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="f7c1f-103">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f7c1f-103">Custom functions parameter options</span></span>

<span data-ttu-id="f7c1f-104">As funções personalizadas são configuráveis com muitas opções de parâmetros diferentes.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="f7c1f-105">Parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="f7c1f-105">Optional parameters</span></span>

<span data-ttu-id="f7c1f-106">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="f7c1f-107">No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="f7c1f-108">Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="f7c1f-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f7c1f-109">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescript"></a>[<span data-ttu-id="f7c1f-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f7c1f-110">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> <span data-ttu-id="f7c1f-111">Quando nenhum valor é especificado para um parâmetro opcional, o Excel atribui a ele o valor `null` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="f7c1f-112">Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="f7c1f-113">Não use a sintaxe `function add(first:number, second:number, third=0):number` porque ela não será inicializada `third` como 0.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="f7c1f-114">Em vez disso, use a sintaxe do TypeScript, conforme mostrado no exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="f7c1f-115">Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontece quando os parâmetros opcionais são nulos.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="f7c1f-116">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="f7c1f-117">Se o `zipCode` parâmetro for NULL, o valor padrão será definido como `98052` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="f7c1f-118">Se o `dayOfWeek` parâmetro for NULL, ele será definido como quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="f7c1f-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f7c1f-119">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescript"></a>[<span data-ttu-id="f7c1f-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f7c1f-120">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a><span data-ttu-id="f7c1f-121">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="f7c1f-121">Range parameters</span></span>

<span data-ttu-id="f7c1f-122">Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="f7c1f-123">Uma função também pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-123">A function can also return a range of data.</span></span> <span data-ttu-id="f7c1f-124">O Excel passará um intervalo de dados de célula como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="f7c1f-125">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f7c1f-126">A função a seguir aceita o parâmetro `values` e a sintaxe JSDOC `number[][]` define a propriedade do parâmetro `dimensionality` como `matrix` nos metadados JSON para essa função.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a><span data-ttu-id="f7c1f-127">Parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="f7c1f-127">Repeating parameters</span></span>

<span data-ttu-id="f7c1f-128">Um parâmetro Repeating permite que o usuário insira uma série de argumentos opcionais para uma função.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="f7c1f-129">Quando a função é chamada, os valores são fornecidos em uma matriz para o parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="f7c1f-130">Se o nome do parâmetro terminar com um número, o número de cada argumento aumentará de forma incremental, como `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="f7c1f-131">Isso corresponde à Convenção usada para funções internas do Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="f7c1f-132">A função a seguir soma o total de números, endereços de célula, bem como intervalos, se inserido.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

<span data-ttu-id="f7c1f-133">Essa função é mostrada `=CONTOSO.ADD([operands], [operands]...)` na pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="f7c1f-134">Parâmetro de valor único repetido</span><span class="sxs-lookup"><span data-stu-id="f7c1f-134">Repeating single value parameter</span></span>

<span data-ttu-id="f7c1f-135">Um parâmetro de valor único repetido permite que vários valores únicos sejam passados.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="f7c1f-136">Por exemplo, o usuário pode inserir ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="f7c1f-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="f7c1f-137">O exemplo a seguir mostra como declarar um parâmetro de valor único.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-137">The following sample shows how to declare a single value parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a><span data-ttu-id="f7c1f-138">Parâmetro de intervalo único</span><span class="sxs-lookup"><span data-stu-id="f7c1f-138">Single range parameter</span></span>

<span data-ttu-id="f7c1f-139">Um único parâmetro de intervalo não é tecnicamente um parâmetro de repetição, mas é incluído aqui porque a declaração é muito parecida com os parâmetros de repetição.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="f7c1f-140">Ele apareceria para o usuário como ADD (a2: B3), em que um único intervalo é passado do Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="f7c1f-141">O exemplo a seguir mostra como declarar um único parâmetro de intervalo.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-141">The following sample shows how to declare a single range parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a><span data-ttu-id="f7c1f-142">Parâmetro de intervalo de repetição</span><span class="sxs-lookup"><span data-stu-id="f7c1f-142">Repeating range parameter</span></span>

<span data-ttu-id="f7c1f-143">Um parâmetro de intervalo de repetição permite que vários intervalos ou números sejam passados.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="f7c1f-144">Por exemplo, o usuário pode inserir ADD (5, B2, C3, 8, E5: E8).</span><span class="sxs-lookup"><span data-stu-id="f7c1f-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="f7c1f-145">Os intervalos de repetição normalmente são especificados com o tipo `number[][][]` , já que são matrizes tridimensionais.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="f7c1f-146">Para obter um exemplo, consulte o exemplo principal listado para parâmetros repetidos (#repeating-Parameters).</span><span class="sxs-lookup"><span data-stu-id="f7c1f-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="f7c1f-147">Declarando parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="f7c1f-147">Declaring repeating parameters</span></span>
<span data-ttu-id="f7c1f-148">No typescript, indique que o parâmetro é multidimensional.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="f7c1f-149">Por exemplo,  `ADD(values: number[])` indicaria uma matriz unidimensional, `ADD(values:number[][])` indicaria uma matriz bidimensional e assim por diante.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="f7c1f-150">Em JavaScript, use `@param values {number[]}` para matrizes unidimensionais, `@param <name> {number[][]}` para matrizes bidimensionais e assim por diante para mais dimensões.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="f7c1f-151">Para o JSON com autoria, certifique-se de que seu parâmetro é especificado como `"repeating": true` em seu arquivo JSON, bem como Verifique se os parâmetros estão marcados como `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="f7c1f-152">Parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="f7c1f-152">Invocation parameter</span></span>

<span data-ttu-id="f7c1f-153">Cada função personalizada é automaticamente passada um `invocation` argumento como o último parâmetro de entrada, mesmo que ele não seja explicitamente declarado.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="f7c1f-154">Esse `invocation` parâmetro corresponde ao objeto de [invocação](/javascript/api/custom-functions-runtime/customfunctions.invocation) .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="f7c1f-155">O `Invocation` objeto pode ser usado para recuperar contexto adicional, como o endereço da célula que chamou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f7c1f-156">Para acessar o `Invocation` objeto, você deve declarar `invocation` como o último parâmetro em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="f7c1f-157">O `invocation` parâmetro não aparece como um argumento de função personalizada para usuários no Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="f7c1f-158">O exemplo a seguir mostra como usar o `invocation` parâmetro para retornar o endereço da célula que chamou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f7c1f-159">Este exemplo usa a propriedade [Address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) do `Invocation` objeto.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="f7c1f-160">Para acessar o `Invocation` objeto, primeiro Declare `CustomFunctions.Invocation` como um parâmetro no seu JSDoc.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="f7c1f-161">Em seguida, declare `@requiresAddress` no JSDoc para acessar a `address` Propriedade do `Invocation` objeto.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="f7c1f-162">Por fim, na função, recupere e retorne a `address` propriedade.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

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

<span data-ttu-id="f7c1f-163">No Excel, uma função personalizada chamando a `address` Propriedade do `Invocation` objeto retornará o endereço absoluto após o formato `SheetName!RelativeCellAddress` na célula que chamou a função.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="f7c1f-164">Por exemplo, se o parâmetro input estiver localizado em uma planilha chamada **Prices** na célula F6, o valor de endereço de parâmetro retornado será `Prices!F6` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="f7c1f-165">O `invocation` parâmetro também pode ser usado para enviar informações para o Excel.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="f7c1f-166">Consulte [fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="f7c1f-167">Detectar o endereço de um parâmetro</span><span class="sxs-lookup"><span data-stu-id="f7c1f-167">Detect the address of a parameter</span></span>

<span data-ttu-id="f7c1f-168">Em combinação com o [parâmetro de chamada](#invocation-parameter), você pode usar o objeto de [invocação](/javascript/api/custom-functions-runtime/customfunctions.invocation) para recuperar o endereço de um parâmetro de entrada de função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="f7c1f-169">Quando invocado, a propriedade [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) do `Invocation` objeto permite que uma função retorne os endereços de todos os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="f7c1f-170">Isso é útil em cenários em que os tipos de dados de entrada podem variar.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="f7c1f-171">O endereço de um parâmetro de entrada pode ser usado para verificar o formato de número do valor de entrada.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="f7c1f-172">O formato de número pode ser ajustado antes da entrada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="f7c1f-173">O endereço de um parâmetro de entrada também pode ser usado para detectar se o valor de entrada tem qualquer propriedade relacionada que possa ser relevante para os cálculos subsequentes.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!IMPORTANT]
> <span data-ttu-id="f7c1f-174">A `parameterAddresses` Propriedade atualmente só funciona com [metadados JSON criados manualmente](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="f7c1f-174">The `parameterAddresses` property currently only works with [manually-created JSON metadata](custom-functions-json.md).</span></span> <span data-ttu-id="f7c1f-175">Para retornar endereços de parâmetro, o `options` objeto deve ter a `requiresParameterAddresses` propriedade definida como `true` e o `result` objeto deve ter a `dimensionality` propriedade definida como `matrix` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-175">To return parameter addresses, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="f7c1f-176">A função personalizada a seguir tem três parâmetros de entrada, recupera a `parameterAddresses` Propriedade do `Invocation` objeto para cada parâmetro e, em seguida, retorna os endereços.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-176">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
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

<span data-ttu-id="f7c1f-177">Quando uma função personalizada chamando a `parameterAddresses` propriedade é executada, o endereço do parâmetro é retornado seguindo o formato `SheetName!RelativeCellAddress` na célula que chamou a função.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-177">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="f7c1f-178">Por exemplo, se o parâmetro input estiver localizado em uma planilha chamada **custos** na célula D8, o valor de endereço de parâmetro retornado será `Costs!D8` .</span><span class="sxs-lookup"><span data-stu-id="f7c1f-178">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="f7c1f-179">Se a função personalizada tiver vários parâmetros e mais de um endereço de parâmetro for retornado, os endereços retornados serão despejados em várias células, decrescente verticalmente da célula que chamou a função.</span><span class="sxs-lookup"><span data-stu-id="f7c1f-179">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="f7c1f-180">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f7c1f-180">Next steps</span></span>

<span data-ttu-id="f7c1f-181">Saiba como usar [valores voláteis em suas funções personalizadas](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="f7c1f-181">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f7c1f-182">Confira também</span><span class="sxs-lookup"><span data-stu-id="f7c1f-182">See also</span></span>

* [<span data-ttu-id="f7c1f-183">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f7c1f-183">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="f7c1f-184">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f7c1f-184">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="f7c1f-185">Criar manualmente metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f7c1f-185">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f7c1f-186">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="f7c1f-186">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f7c1f-187">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="f7c1f-187">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
