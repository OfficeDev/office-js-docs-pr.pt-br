---
ms.date: 02/04/2021
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: afe6947b1a1b9022a0284535b9ab1d68c9777c14
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173903"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="ba5c0-103">Opções de parâmetro de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ba5c0-103">Custom functions parameter options</span></span>

<span data-ttu-id="ba5c0-104">Funções personalizadas são configuráveis com muitas opções de parâmetro diferentes.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="ba5c0-105">Parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="ba5c0-105">Optional parameters</span></span>

<span data-ttu-id="ba5c0-106">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="ba5c0-107">No exemplo a seguir, a função add pode, opcionalmente, adicionar um terceiro número.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="ba5c0-108">Esta função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="ba5c0-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="ba5c0-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="ba5c0-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="ba5c0-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="ba5c0-111">Quando nenhum valor é especificado para um parâmetro opcional, o Excel atribui a ele o valor `null` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="ba5c0-112">Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="ba5c0-113">Não use a sintaxe porque `function add(first:number, second:number, third=0):number` ela não será inicializada como `third` 0.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="ba5c0-114">Em vez disso, use a sintaxe TypeScript conforme mostrado no exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="ba5c0-115">Quando você define uma função que contém um ou mais parâmetros opcionais, especifique o que acontece quando os parâmetros opcionais são nulos.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="ba5c0-116">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="ba5c0-117">Se o `zipCode` parâmetro for nulo, o valor padrão será definido como `98052` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="ba5c0-118">Se o `dayOfWeek` parâmetro for nulo, ele será definido como quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="ba5c0-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="ba5c0-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="ba5c0-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="ba5c0-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="ba5c0-121">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="ba5c0-121">Range parameters</span></span>

<span data-ttu-id="ba5c0-122">Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="ba5c0-123">Uma função também pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-123">A function can also return a range of data.</span></span> <span data-ttu-id="ba5c0-124">O Excel passará um intervalo de dados de célula como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="ba5c0-125">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="ba5c0-126">A função a seguir aceita o parâmetro e a sintaxe JSDOC define a propriedade do parâmetro como nos metadados `values` `number[][]` `dimensionality` `matrix` JSON dessa função.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="ba5c0-127">Parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="ba5c0-127">Repeating parameters</span></span>

<span data-ttu-id="ba5c0-128">Um parâmetro de repetição permite que um usuário insira uma série de argumentos opcionais em uma função.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="ba5c0-129">Quando a função é chamada, os valores são fornecidos em uma matriz para o parâmetro.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="ba5c0-130">Se o nome do parâmetro terminar com um número, o número de cada argumento aumentará incrementalmente, como `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="ba5c0-131">Isso corresponde à convenção usada para funções do Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="ba5c0-132">A função a seguir soma o total de números, endereços de células, bem como intervalos, se inseridos.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="ba5c0-133">Esta função mostra `=CONTOSO.ADD([operands], [operands]...)` na planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="ba5c0-134">Parâmetro de valor único de repetição</span><span class="sxs-lookup"><span data-stu-id="ba5c0-134">Repeating single value parameter</span></span>

<span data-ttu-id="ba5c0-135">Um parâmetro de valor único de repetição permite que vários valores individuais sejam passados.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="ba5c0-136">Por exemplo, o usuário pode inserir ADD(1,B2,3).</span><span class="sxs-lookup"><span data-stu-id="ba5c0-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="ba5c0-137">O exemplo a seguir mostra como declarar um único parâmetro de valor.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="ba5c0-138">Parâmetro de intervalo único</span><span class="sxs-lookup"><span data-stu-id="ba5c0-138">Single range parameter</span></span>

<span data-ttu-id="ba5c0-139">Um parâmetro de intervalo único não é tecnicamente um parâmetro de repetição, mas está incluído aqui porque a declaração é muito semelhante aos parâmetros de repetição.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="ba5c0-140">Ele seria exibido para o usuário como ADD(A2:B3) onde um único intervalo é passado do Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="ba5c0-141">O exemplo a seguir mostra como declarar um único parâmetro de intervalo.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="ba5c0-142">Parâmetro de intervalo de repetição</span><span class="sxs-lookup"><span data-stu-id="ba5c0-142">Repeating range parameter</span></span>

<span data-ttu-id="ba5c0-143">Um parâmetro de intervalo de repetição permite que vários intervalos ou números sejam passados.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="ba5c0-144">Por exemplo, o usuário pode inserir ADD(5,B2,C3,8,E5:E8).</span><span class="sxs-lookup"><span data-stu-id="ba5c0-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="ba5c0-145">Os intervalos repetidos geralmente são especificados com o tipo, pois `number[][][]` são matrizes tridimensionais.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="ba5c0-146">Para ver um exemplo, consulte o exemplo principal listado para parâmetros de repetição(#repeating-parâmetros).</span><span class="sxs-lookup"><span data-stu-id="ba5c0-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="ba5c0-147">Declarando parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="ba5c0-147">Declaring repeating parameters</span></span>
<span data-ttu-id="ba5c0-148">Em Typescript, indique que o parâmetro é multidimensional.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="ba5c0-149">Por exemplo,  `ADD(values: number[])` indicaria uma matriz unidimensional, indicaria uma matriz `ADD(values:number[][])` bidimensional e assim por diante.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="ba5c0-150">Em JavaScript, use para matrizes unidimensionais, para matrizes bidimensionais e assim por diante `@param values {number[]}` `@param <name> {number[][]}` para mais dimensões.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="ba5c0-151">Para JSON de autoria manual, certifique-se de que seu parâmetro está especificado como em seu arquivo JSON, bem como verifique se seus parâmetros estão `"repeating": true` marcados como `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="ba5c0-152">Parâmetro invocation</span><span class="sxs-lookup"><span data-stu-id="ba5c0-152">Invocation parameter</span></span>

<span data-ttu-id="ba5c0-153">Cada função personalizada passa automaticamente um argumento como o último parâmetro de entrada, mesmo que não `invocation` seja explicitamente declarado.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="ba5c0-154">Esse `invocation` parâmetro corresponde ao objeto [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="ba5c0-155">O objeto pode ser usado para recuperar contexto adicional, como o endereço da `Invocation` célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="ba5c0-156">Para acessar o `Invocation` objeto, você deve `invocation` declarar como o último parâmetro em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="ba5c0-157">O `invocation` parâmetro não aparece como um argumento de função personalizada para usuários no Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="ba5c0-158">O exemplo a seguir mostra como usar o parâmetro para retornar o endereço da `invocation` célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="ba5c0-159">Este exemplo usa a [propriedade](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) de endereço do `Invocation` objeto.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="ba5c0-160">Para acessar o `Invocation` objeto, primeiro `CustomFunctions.Invocation` declare como um parâmetro em seu JSDoc.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="ba5c0-161">Em seguida, `@requiresAddress` declare em seu JSDoc para acessar `address` a propriedade do `Invocation` objeto.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="ba5c0-162">Por fim, dentro da função, recupere e retorne a `address` propriedade.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

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

<span data-ttu-id="ba5c0-163">No Excel, uma função personalizada que chama a propriedade do objeto retornará o endereço absoluto após o formato na célula `address` `Invocation` que `SheetName!RelativeCellAddress` invocou a função.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="ba5c0-164">Por exemplo, se o parâmetro de entrada estiver localizado em uma planilha chamada **Prices** na célula F6, o valor do endereço do parâmetro retornado será `Prices!F6` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="ba5c0-165">O `invocation` parâmetro também pode ser usado para enviar informações ao Excel.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="ba5c0-166">Consulte [Fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="ba5c0-167">Detectar o endereço de um parâmetro</span><span class="sxs-lookup"><span data-stu-id="ba5c0-167">Detect the address of a parameter</span></span>

<span data-ttu-id="ba5c0-168">Em combinação com o [parâmetro de invocação,](#invocation-parameter)você pode usar o [objeto Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) para recuperar o endereço de um parâmetro de entrada de função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="ba5c0-169">Quando invocada, a [propriedade parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) do objeto permite que uma função retorne os `Invocation` endereços de todos os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="ba5c0-170">Isso é útil em cenários em que os tipos de dados de entrada podem variar.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="ba5c0-171">O endereço de um parâmetro de entrada pode ser usado para verificar o formato de número do valor de entrada.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="ba5c0-172">O formato de número pode ser ajustado antes da entrada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="ba5c0-173">O endereço de um parâmetro de entrada também pode ser usado para detectar se o valor de entrada tem propriedades relacionadas que podem ser relevantes para cálculos subsequentes.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!NOTE]
> <span data-ttu-id="ba5c0-174">Se você estiver trabalhando com metadados [JSON](custom-functions-json.md) criados manualmente para retornar endereços de parâmetro em vez do gerador Yo Office, o objeto deve ter a propriedade definida como , e o objeto deve ter a propriedade definida como `options` `requiresParameterAddresses` `true` `result` `dimensionality` `matrix` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-174">If you're working with [manually-created JSON metadata](custom-functions-json.md) to return parameter addresses instead of the Yo Office generator, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="ba5c0-175">A função personalizada a seguir recebe três parâmetros de entrada, recupera a propriedade do objeto para cada parâmetro e `parameterAddresses` `Invocation` retorna os endereços.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-175">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

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

<span data-ttu-id="ba5c0-176">Quando uma função personalizada que chama a propriedade é executado, o endereço do parâmetro é retornado seguindo o formato na `parameterAddresses` `SheetName!RelativeCellAddress` célula que invocou a função.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-176">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="ba5c0-177">Por exemplo, se o parâmetro de entrada estiver localizado em uma planilha chamada **Custos** na célula D8, o valor do endereço do parâmetro retornado será `Costs!D8` .</span><span class="sxs-lookup"><span data-stu-id="ba5c0-177">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="ba5c0-178">Se a função personalizada tiver vários parâmetros e mais de um endereço de parâmetro for retornado, os endereços retornados serão derramamentos entre várias células, descendentes verticalmente da célula que invocou a função.</span><span class="sxs-lookup"><span data-stu-id="ba5c0-178">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="ba5c0-179">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ba5c0-179">Next steps</span></span>

<span data-ttu-id="ba5c0-180">Saiba como usar valores [voláteis em suas funções personalizadas.](custom-functions-volatile.md)</span><span class="sxs-lookup"><span data-stu-id="ba5c0-180">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ba5c0-181">Confira também</span><span class="sxs-lookup"><span data-stu-id="ba5c0-181">See also</span></span>

* [<span data-ttu-id="ba5c0-182">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ba5c0-182">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="ba5c0-183">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ba5c0-183">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="ba5c0-184">Criar manualmente metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ba5c0-184">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ba5c0-185">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="ba5c0-185">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="ba5c0-186">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="ba5c0-186">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
