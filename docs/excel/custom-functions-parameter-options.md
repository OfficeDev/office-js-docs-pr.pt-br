---
ms.date: 07/15/2019
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: e5b75b098d64d5998b0393d5995896f0289337fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771416"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="f19d9-103">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f19d9-103">Custom functions parameter options</span></span>

<span data-ttu-id="f19d9-104">Funções personalizadas são configuráveis com muitas opções diferentes para parâmetros.</span><span class="sxs-lookup"><span data-stu-id="f19d9-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="f19d9-105">Parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="f19d9-105">Optional parameters</span></span>

<span data-ttu-id="f19d9-106">Enquanto parâmetros regulares são necessários, os parâmetros opcionais não.</span><span class="sxs-lookup"><span data-stu-id="f19d9-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="f19d9-107">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="f19d9-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="f19d9-108">No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número.</span><span class="sxs-lookup"><span data-stu-id="f19d9-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="f19d9-109">Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="f19d9-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="f19d9-110">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f19d9-110">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="f19d9-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f19d9-111">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="f19d9-112">Quando nenhum valor é especificado para um parâmetro opcional, o Excel atribui a ele o valor `null`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="f19d9-113">Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="f19d9-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="f19d9-114">Portanto, não use a sintaxe `function add(first:number, second:number, third=0):number` porque ela não será inicializada `third` como 0.</span><span class="sxs-lookup"><span data-stu-id="f19d9-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="f19d9-115">Em vez disso, use a sintaxe do TypeScript, conforme mostrado no exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="f19d9-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="f19d9-116">Ao definir uma função que contenha um ou mais parâmetros opcionais, você deve especificar o que acontece quando os parâmetros opcionais são nulos.</span><span class="sxs-lookup"><span data-stu-id="f19d9-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="f19d9-117">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="f19d9-118">Se o `zipCode` parâmetro for NULL, o valor padrão será definido como `98052`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="f19d9-119">Se o `dayOfWeek` parâmetro for NULL, ele será definido como quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="f19d9-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="f19d9-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f19d9-120">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="f19d9-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f19d9-121">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="f19d9-122">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="f19d9-122">Range parameters</span></span>

<span data-ttu-id="f19d9-123">Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada.</span><span class="sxs-lookup"><span data-stu-id="f19d9-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="f19d9-124">Uma função também pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="f19d9-124">A function can also return a range of data.</span></span> <span data-ttu-id="f19d9-125">O Excel passará um intervalo de dados de célula como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="f19d9-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="f19d9-126">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="f19d9-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f19d9-127">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="f19d9-128">Observe que, nos metadados JSON dessa função, a propriedade do `type` parâmetro é definida como. `matrix`</span><span class="sxs-lookup"><span data-stu-id="f19d9-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

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

## <a name="repeating-parameters"></a><span data-ttu-id="f19d9-129">Parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="f19d9-129">Repeating parameters</span></span>

<span data-ttu-id="f19d9-130">Um parâmetro Repeating permite que o usuário insira uma série de argumentos opcionais para uma função.</span><span class="sxs-lookup"><span data-stu-id="f19d9-130">A repeating parameter allows a user to enter a series of optional of arguments to a function.</span></span> <span data-ttu-id="f19d9-131">Quando a função é chamada, os valores são fornecidos em uma matriz para o parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f19d9-131">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="f19d9-132">Se o nome do parâmetro terminar com um número, cada argumento aumentará o número, como `ADD(number1, [number2], [number3],…)`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-132">If the parameter name ends with a number, each argument will increment the number, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="f19d9-133">Isso corresponde à Convenção usada para funções internas do Excel.</span><span class="sxs-lookup"><span data-stu-id="f19d9-133">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="f19d9-134">A função a seguir soma o total de números, endereços de célula, bem como intervalos, se inserido.</span><span class="sxs-lookup"><span data-stu-id="f19d9-134">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="f19d9-135">Essa função é `=CONTOSO.ADD([operands], [operands]...)` mostrada na pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="f19d9-135">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="f19d9-136">Parâmetro de valor único repetido</span><span class="sxs-lookup"><span data-stu-id="f19d9-136">Repeating single value parameter</span></span>

<span data-ttu-id="f19d9-137">Um parâmetro de valor único repetido permite que vários valores únicos sejam passados.</span><span class="sxs-lookup"><span data-stu-id="f19d9-137">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="f19d9-138">Por exemplo, o usuário pode inserir ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="f19d9-138">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="f19d9-139">O exemplo a seguir mostra como declarar um parâmetro de valor único.</span><span class="sxs-lookup"><span data-stu-id="f19d9-139">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="f19d9-140">Parâmetro de intervalo único</span><span class="sxs-lookup"><span data-stu-id="f19d9-140">Single range parameter</span></span>

<span data-ttu-id="f19d9-141">Um único parâmetro de intervalo não é tecnicamente um parâmetro de repetição, mas é incluído aqui porque a declaração é muito parecida com os parâmetros de repetição.</span><span class="sxs-lookup"><span data-stu-id="f19d9-141">A single range parameter is not technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="f19d9-142">Ele apareceria para o usuário como ADD (a2: B3), em que um único intervalo é passado do Excel.</span><span class="sxs-lookup"><span data-stu-id="f19d9-142">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="f19d9-143">O exemplo a seguir mostra como declarar um único parâmetro de intervalo.</span><span class="sxs-lookup"><span data-stu-id="f19d9-143">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="f19d9-144">Parâmetro de intervalo de repetição</span><span class="sxs-lookup"><span data-stu-id="f19d9-144">Repeating range parameter</span></span>

<span data-ttu-id="f19d9-145">Um parâmetro de intervalo de repetição permite que vários intervalos ou números sejam passados.</span><span class="sxs-lookup"><span data-stu-id="f19d9-145">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="f19d9-146">Por exemplo, o usuário pode inserir ADD (5, B2, C3, 8, E5: E8).</span><span class="sxs-lookup"><span data-stu-id="f19d9-146">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="f19d9-147">Os intervalos de repetição normalmente são especificados com `number[][][]` o tipo, já que são matrizes tridimensionais.</span><span class="sxs-lookup"><span data-stu-id="f19d9-147">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="f19d9-148">Para obter um exemplo, consulte o exemplo principal listado para parâmetros repetidos (#repeating-Parameters).</span><span class="sxs-lookup"><span data-stu-id="f19d9-148">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="f19d9-149">Declarando parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="f19d9-149">Declaring repeating parameters</span></span>
<span data-ttu-id="f19d9-150">No typescript, indique que o parâmetro é multidimensional.</span><span class="sxs-lookup"><span data-stu-id="f19d9-150">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="f19d9-151">Por exemplo, `ADD(values: number[])` indicaria uma matriz unidimensional, `ADD(values:number[][])` indicaria uma matriz bidimensional e assim por diante.</span><span class="sxs-lookup"><span data-stu-id="f19d9-151">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="f19d9-152">Em JavaScript, use `@param values {number[]}` para matrizes unidimensionais `@param <name> {number[][]}` , para matrizes bidimensionais e assim por diante para mais dimensões.</span><span class="sxs-lookup"><span data-stu-id="f19d9-152">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="f19d9-153">Para o JSON com autoria, certifique-se de que seu `"repeating": true` parâmetro é especificado como em seu arquivo JSON, bem como Verifique se os parâmetros `"dimensionality”: matrix`estão marcados como.</span><span class="sxs-lookup"><span data-stu-id="f19d9-153">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality”: matrix`.</span></span>

>[!NOTE]
><span data-ttu-id="f19d9-154">Funções contendo parâmetros repetidos contêm automaticamente um parâmetro de chamada como o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f19d9-154">Functions containing repeating parameters automatically contain an invocation parameter as the last parameter.</span></span> <span data-ttu-id="f19d9-155">Para obter mais informações sobre parâmetros de chamada, consulte a seção a seguir.</span><span class="sxs-lookup"><span data-stu-id="f19d9-155">For more information on invocation parameters, see the following section.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="f19d9-156">Parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="f19d9-156">Invocation parameter</span></span>

<span data-ttu-id="f19d9-157">Cada função personalizada é automaticamente passada um `invocation` argumento como o último argumento.</span><span class="sxs-lookup"><span data-stu-id="f19d9-157">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="f19d9-158">Esse argumento pode ser usado para recuperar contexto adicional, como o endereço da célula de chamada.</span><span class="sxs-lookup"><span data-stu-id="f19d9-158">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="f19d9-159">Ou pode ser usado para enviar informações para o Excel, como um manipulador de função para [cancelar uma função](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="f19d9-159">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="f19d9-160">Mesmo que você declare nenhum parâmetro, sua função personalizada tem esse parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f19d9-160">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="f19d9-161">Esse argumento não aparece para um usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="f19d9-161">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="f19d9-162">Se você deseja usar `invocation` em sua função personalizada, declare-a como o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f19d9-162">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="f19d9-163">No exemplo de código a seguir, `invocation` o contexto é explicitamente declarado para sua referência.</span><span class="sxs-lookup"><span data-stu-id="f19d9-163">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
```

<span data-ttu-id="f19d9-164">O parâmetro permite que você obtenha o contexto da célula de invocação, que pode ser útil em alguns cenários, incluindo [a descoberta do endereço de uma célula que invoque uma função personalizada](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="f19d9-164">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="f19d9-165">Parâmetro de contexto da célula de endereçamento</span><span class="sxs-lookup"><span data-stu-id="f19d9-165">Addressing cell's context parameter</span></span>

<span data-ttu-id="f19d9-166">Em alguns casos, você precisa obter o endereço da célula que chamou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f19d9-166">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f19d9-167">Isso é útil nos seguintes cenários:</span><span class="sxs-lookup"><span data-stu-id="f19d9-167">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="f19d9-168">Intervalos de formatação: Use o endereço da célula como a chave para armazenar informações no [OfficeRuntime. armazenamento](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="f19d9-168">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="f19d9-169">Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-169">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="f19d9-170">Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `OfficeRuntime.storage` usando `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-170">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="f19d9-171">Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.</span><span class="sxs-lookup"><span data-stu-id="f19d9-171">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="f19d9-172">Para solicitar um contexto de uma célula de endereçamento em uma função, você precisa usar uma função para localizar o endereço da célula, como a do exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f19d9-172">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="f19d9-173">As informações sobre o endereço de uma célula são expostas apenas se `@requiresAddress` o estiver marcado nos comentários da função.</span><span class="sxs-lookup"><span data-stu-id="f19d9-173">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
```

<span data-ttu-id="f19d9-174">Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-174">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="f19d9-175">Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="f19d9-175">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f19d9-176">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f19d9-176">Next steps</span></span>

<span data-ttu-id="f19d9-177">Saiba como [salvar o estado em suas funções personalizadas](custom-functions-save-state.md) ou usar [valores voláteis em suas funções personalizadas](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="f19d9-177">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f19d9-178">Confira também</span><span class="sxs-lookup"><span data-stu-id="f19d9-178">See also</span></span>

* [<span data-ttu-id="f19d9-179">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f19d9-179">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="f19d9-180">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f19d9-180">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f19d9-181">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f19d9-181">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="f19d9-182">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="f19d9-182">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f19d9-183">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="f19d9-183">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)