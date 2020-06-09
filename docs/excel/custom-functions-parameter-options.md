---
ms.date: 04/29/2020
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: ee193ed68ef59bfd9068bc43cd30721d6bb7b86a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609272"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="85aff-103">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85aff-103">Custom functions parameter options</span></span>

<span data-ttu-id="85aff-104">As funções personalizadas são configuráveis com muitas opções de parâmetros diferentes.</span><span class="sxs-lookup"><span data-stu-id="85aff-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="85aff-105">Parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="85aff-105">Optional parameters</span></span>

<span data-ttu-id="85aff-106">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="85aff-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="85aff-107">No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número.</span><span class="sxs-lookup"><span data-stu-id="85aff-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="85aff-108">Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="85aff-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="85aff-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="85aff-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="85aff-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="85aff-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="85aff-111">Quando nenhum valor é especificado para um parâmetro opcional, o Excel atribui a ele o valor `null` .</span><span class="sxs-lookup"><span data-stu-id="85aff-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="85aff-112">Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="85aff-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="85aff-113">Não use a sintaxe `function add(first:number, second:number, third=0):number` porque ela não será inicializada `third` como 0.</span><span class="sxs-lookup"><span data-stu-id="85aff-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="85aff-114">Em vez disso, use a sintaxe do TypeScript, conforme mostrado no exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="85aff-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="85aff-115">Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontece quando os parâmetros opcionais são nulos.</span><span class="sxs-lookup"><span data-stu-id="85aff-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="85aff-116">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="85aff-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="85aff-117">Se o `zipCode` parâmetro for NULL, o valor padrão será definido como `98052` .</span><span class="sxs-lookup"><span data-stu-id="85aff-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="85aff-118">Se o `dayOfWeek` parâmetro for NULL, ele será definido como quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="85aff-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="85aff-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="85aff-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="85aff-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="85aff-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="85aff-121">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="85aff-121">Range parameters</span></span>

<span data-ttu-id="85aff-122">Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada.</span><span class="sxs-lookup"><span data-stu-id="85aff-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="85aff-123">Uma função também pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="85aff-123">A function can also return a range of data.</span></span> <span data-ttu-id="85aff-124">O Excel passará um intervalo de dados de célula como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="85aff-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="85aff-125">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="85aff-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="85aff-126">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="85aff-126">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="85aff-127">Observe que, nos metadados JSON dessa função, a propriedade do parâmetro `type` é definida como `matrix` .</span><span class="sxs-lookup"><span data-stu-id="85aff-127">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

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

## <a name="repeating-parameters"></a><span data-ttu-id="85aff-128">Parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="85aff-128">Repeating parameters</span></span>

<span data-ttu-id="85aff-129">Um parâmetro Repeating permite que o usuário insira uma série de argumentos opcionais para uma função.</span><span class="sxs-lookup"><span data-stu-id="85aff-129">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="85aff-130">Quando a função é chamada, os valores são fornecidos em uma matriz para o parâmetro.</span><span class="sxs-lookup"><span data-stu-id="85aff-130">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="85aff-131">Se o nome do parâmetro terminar com um número, o número de cada argumento aumentará de forma incremental, como `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="85aff-131">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="85aff-132">Isso corresponde à Convenção usada para funções internas do Excel.</span><span class="sxs-lookup"><span data-stu-id="85aff-132">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="85aff-133">A função a seguir soma o total de números, endereços de célula, bem como intervalos, se inserido.</span><span class="sxs-lookup"><span data-stu-id="85aff-133">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="85aff-134">Essa função é mostrada `=CONTOSO.ADD([operands], [operands]...)` na pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="85aff-134">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="85aff-135">Parâmetro de valor único repetido</span><span class="sxs-lookup"><span data-stu-id="85aff-135">Repeating single value parameter</span></span>

<span data-ttu-id="85aff-136">Um parâmetro de valor único repetido permite que vários valores únicos sejam passados.</span><span class="sxs-lookup"><span data-stu-id="85aff-136">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="85aff-137">Por exemplo, o usuário pode inserir ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="85aff-137">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="85aff-138">O exemplo a seguir mostra como declarar um parâmetro de valor único.</span><span class="sxs-lookup"><span data-stu-id="85aff-138">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="85aff-139">Parâmetro de intervalo único</span><span class="sxs-lookup"><span data-stu-id="85aff-139">Single range parameter</span></span>

<span data-ttu-id="85aff-140">Um único parâmetro de intervalo não é tecnicamente um parâmetro de repetição, mas é incluído aqui porque a declaração é muito parecida com os parâmetros de repetição.</span><span class="sxs-lookup"><span data-stu-id="85aff-140">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="85aff-141">Ele apareceria para o usuário como ADD (a2: B3), em que um único intervalo é passado do Excel.</span><span class="sxs-lookup"><span data-stu-id="85aff-141">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="85aff-142">O exemplo a seguir mostra como declarar um único parâmetro de intervalo.</span><span class="sxs-lookup"><span data-stu-id="85aff-142">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="85aff-143">Parâmetro de intervalo de repetição</span><span class="sxs-lookup"><span data-stu-id="85aff-143">Repeating range parameter</span></span>

<span data-ttu-id="85aff-144">Um parâmetro de intervalo de repetição permite que vários intervalos ou números sejam passados.</span><span class="sxs-lookup"><span data-stu-id="85aff-144">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="85aff-145">Por exemplo, o usuário pode inserir ADD (5, B2, C3, 8, E5: E8).</span><span class="sxs-lookup"><span data-stu-id="85aff-145">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="85aff-146">Os intervalos de repetição normalmente são especificados com o tipo `number[][][]` , já que são matrizes tridimensionais.</span><span class="sxs-lookup"><span data-stu-id="85aff-146">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="85aff-147">Para obter um exemplo, consulte o exemplo principal listado para parâmetros repetidos (#repeating-Parameters).</span><span class="sxs-lookup"><span data-stu-id="85aff-147">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="85aff-148">Declarando parâmetros de repetição</span><span class="sxs-lookup"><span data-stu-id="85aff-148">Declaring repeating parameters</span></span>
<span data-ttu-id="85aff-149">No typescript, indique que o parâmetro é multidimensional.</span><span class="sxs-lookup"><span data-stu-id="85aff-149">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="85aff-150">Por exemplo, `ADD(values: number[])` indicaria uma matriz unidimensional, `ADD(values:number[][])` indicaria uma matriz bidimensional e assim por diante.</span><span class="sxs-lookup"><span data-stu-id="85aff-150">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="85aff-151">Em JavaScript, use `@param values {number[]}` para matrizes unidimensionais, `@param <name> {number[][]}` para matrizes bidimensionais e assim por diante para mais dimensões.</span><span class="sxs-lookup"><span data-stu-id="85aff-151">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="85aff-152">Para o JSON com autoria, certifique-se de que seu parâmetro é especificado como `"repeating": true` em seu arquivo JSON, bem como Verifique se os parâmetros estão marcados como `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="85aff-152">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="85aff-153">Parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="85aff-153">Invocation parameter</span></span>

<span data-ttu-id="85aff-154">Cada função personalizada é automaticamente passada um `invocation` argumento como o último argumento.</span><span class="sxs-lookup"><span data-stu-id="85aff-154">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="85aff-155">Esse argumento pode ser usado para recuperar contexto adicional, como o endereço da célula de chamada.</span><span class="sxs-lookup"><span data-stu-id="85aff-155">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="85aff-156">Ou pode ser usado para enviar informações para o Excel, como um manipulador de função para [cancelar uma função](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="85aff-156">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="85aff-157">Mesmo que você declare nenhum parâmetro, sua função personalizada tem esse parâmetro.</span><span class="sxs-lookup"><span data-stu-id="85aff-157">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="85aff-158">Esse argumento não aparece para um usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="85aff-158">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="85aff-159">Se você deseja usar `invocation` em sua função personalizada, declare-a como o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="85aff-159">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="85aff-160">No exemplo de código a seguir, o `invocation` contexto é explicitamente declarado para sua referência.</span><span class="sxs-lookup"><span data-stu-id="85aff-160">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="85aff-161">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="85aff-161">Next steps</span></span>

<span data-ttu-id="85aff-162">Saiba como usar [valores voláteis em suas funções personalizadas](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="85aff-162">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="85aff-163">Confira também</span><span class="sxs-lookup"><span data-stu-id="85aff-163">See also</span></span>

* [<span data-ttu-id="85aff-164">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85aff-164">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="85aff-165">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85aff-165">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="85aff-166">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85aff-166">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="85aff-167">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="85aff-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="85aff-168">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="85aff-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
