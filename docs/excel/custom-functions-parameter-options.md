---
ms.date: 07/01/2019
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: 9416653d697bdf36ca698271e00d9742ff0e75a9
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617041"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="81240-103">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="81240-103">Custom functions parameter options</span></span>

<span data-ttu-id="81240-104">Funções personalizadas são configuráveis com muitas opções diferentes para parâmetros.</span><span class="sxs-lookup"><span data-stu-id="81240-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="81240-105">Parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="81240-105">Optional parameters</span></span>

<span data-ttu-id="81240-106">Enquanto parâmetros regulares são necessários, os parâmetros opcionais não.</span><span class="sxs-lookup"><span data-stu-id="81240-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="81240-107">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="81240-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="81240-108">No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número.</span><span class="sxs-lookup"><span data-stu-id="81240-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="81240-109">Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="81240-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="81240-110">JavaScript</span><span class="sxs-lookup"><span data-stu-id="81240-110">JavaScript</span></span>](#tab/javascript)

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
CustomFunctions.associate("ADD", add);
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="81240-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="81240-111">TypeScript</span></span>](#tab/typescript)

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
CustomFunctions.associate("ADD", add);
```

---

> [!NOTE]
> <span data-ttu-id="81240-112">Quando nenhum valor é especificado para um parâmetro opcional, o Excel atribui a ele o valor `null`.</span><span class="sxs-lookup"><span data-stu-id="81240-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="81240-113">Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="81240-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="81240-114">Portanto, não use a sintaxe `function add(first:number, second:number, third=0):number` porque ela não será inicializada `third` como 0.</span><span class="sxs-lookup"><span data-stu-id="81240-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="81240-115">Em vez disso, use a sintaxe do TypeScript, conforme mostrado no exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="81240-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="81240-116">Ao definir uma função que contenha um ou mais parâmetros opcionais, você deve especificar o que acontece quando os parâmetros opcionais são nulos.</span><span class="sxs-lookup"><span data-stu-id="81240-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="81240-117">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="81240-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="81240-118">Se o `zipCode` parâmetro for NULL, o valor padrão será definido como `98052`.</span><span class="sxs-lookup"><span data-stu-id="81240-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="81240-119">Se o `dayOfWeek` parâmetro for NULL, ele será definido como quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="81240-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="81240-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="81240-120">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
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

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="81240-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="81240-121">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string
{
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

## <a name="range-parameters"></a><span data-ttu-id="81240-122">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="81240-122">Range parameters</span></span>

<span data-ttu-id="81240-123">Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada.</span><span class="sxs-lookup"><span data-stu-id="81240-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="81240-124">Uma função também pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="81240-124">A function can also return a range of data.</span></span> <span data-ttu-id="81240-125">O Excel passará um intervalo de dados de célula como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="81240-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="81240-126">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="81240-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="81240-127">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="81240-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="81240-128">Observe que, nos metadados JSON dessa função, a propriedade do `type` parâmetro é definida como. `matrix`</span><span class="sxs-lookup"><span data-stu-id="81240-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a><span data-ttu-id="81240-129">Parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="81240-129">Invocation parameter</span></span>

<span data-ttu-id="81240-130">Cada função personalizada é automaticamente passada um `invocation` argumento como o último argumento.</span><span class="sxs-lookup"><span data-stu-id="81240-130">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="81240-131">Esse argumento pode ser usado para recuperar contexto adicional, como o endereço da célula de chamada.</span><span class="sxs-lookup"><span data-stu-id="81240-131">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="81240-132">Ou pode ser usado para enviar informações para o Excel, como um manipulador de função para [cancelar uma função](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="81240-132">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="81240-133">Mesmo que você declare nenhum parâmetro, sua função personalizada tem esse parâmetro.</span><span class="sxs-lookup"><span data-stu-id="81240-133">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="81240-134">Esse argumento não aparece para um usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="81240-134">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="81240-135">Se você deseja usar `invocation` em sua função personalizada, declare-a como o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="81240-135">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="81240-136">No exemplo de código a seguir, `invocation` o contexto é explicitamente declarado para sua referência.</span><span class="sxs-lookup"><span data-stu-id="81240-136">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="81240-137">O parâmetro permite que você obtenha o contexto da célula de invocação, que pode ser útil em alguns cenários, incluindo [a descoberta do endereço de uma célula que invoque uma função personalizada](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="81240-137">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="81240-138">Parâmetro de contexto da célula de endereçamento</span><span class="sxs-lookup"><span data-stu-id="81240-138">Addressing cell's context parameter</span></span>

<span data-ttu-id="81240-139">Em alguns casos, você precisa obter o endereço da célula que chamou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="81240-139">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="81240-140">Isso é útil nos seguintes cenários:</span><span class="sxs-lookup"><span data-stu-id="81240-140">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="81240-141">Intervalos de formatação: Use o endereço da célula como a chave para armazenar informações no [OfficeRuntime. armazenamento](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="81240-141">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="81240-142">Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="81240-142">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="81240-143">Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `OfficeRuntime.storage` usando `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="81240-143">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="81240-144">Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.</span><span class="sxs-lookup"><span data-stu-id="81240-144">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="81240-145">Para solicitar um contexto de uma célula de endereçamento em uma função, você precisa usar uma função para localizar o endereço da célula, como a do exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="81240-145">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="81240-146">As informações sobre o endereço de uma célula são expostas apenas se `@requiresAddress` o estiver marcado nos comentários da função.</span><span class="sxs-lookup"><span data-stu-id="81240-146">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

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
CustomFunctions.associate("GETADDRESS", getAddress);
```

<span data-ttu-id="81240-147">Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="81240-147">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="81240-148">Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="81240-148">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="81240-149">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="81240-149">Next steps</span></span>
<span data-ttu-id="81240-150">Saiba como [salvar o estado em suas funções personalizadas](custom-functions-save-state.md) ou usar [valores voláteis em suas funções personalizadas](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="81240-150">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="81240-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="81240-151">See also</span></span>

* [<span data-ttu-id="81240-152">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="81240-152">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* <span data-ttu-id="81240-153">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="81240-153">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="81240-154">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="81240-154">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="81240-155">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="81240-155">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="81240-156">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="81240-156">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="81240-157">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="81240-157">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
