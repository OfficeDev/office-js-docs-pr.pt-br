---
ms.date: 04/30/2019
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel (visualização)
localization_priority: Normal
ms.openlocfilehash: b5dba59431f4c6ec4ee08c563e7cb3affeb06608
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527294"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="bad73-103">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bad73-103">Custom functions parameter options</span></span>

<span data-ttu-id="bad73-104">As funções personalizadas são configuráveis com muitas opções diferentes para parâmetros:</span><span class="sxs-lookup"><span data-stu-id="bad73-104">Custom functions are configurable with many different options for parameters:</span></span> 
- [<span data-ttu-id="bad73-105">Parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="bad73-105">Optional parameters</span></span>](#custom-functions-optional-parameters)
- [<span data-ttu-id="bad73-106">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="bad73-106">Range parameters</span></span>](#range-parameters)
- [<span data-ttu-id="bad73-107">Parâmetro de contexto de invocação</span><span class="sxs-lookup"><span data-stu-id="bad73-107">Invocation context parameter</span></span>](#invocation-parameter)

## <a name="custom-functions-optional-parameters"></a><span data-ttu-id="bad73-108">Parâmetros opcionais de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bad73-108">Custom functions optional parameters</span></span>

<span data-ttu-id="bad73-109">Enquanto parâmetros regulares são necessários, os parâmetros opcionais não.</span><span class="sxs-lookup"><span data-stu-id="bad73-109">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="bad73-110">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="bad73-110">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="bad73-111">No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número.</span><span class="sxs-lookup"><span data-stu-id="bad73-111">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="bad73-112">Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="bad73-112">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="bad73-113">Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontecerá quando os parâmetros opcionais forem indefinidos.</span><span class="sxs-lookup"><span data-stu-id="bad73-113">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="bad73-114">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="bad73-114">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="bad73-115">Se o `zipCode` parâmetro estiver indefinido, o valor padrão será definido como `98052`.</span><span class="sxs-lookup"><span data-stu-id="bad73-115">If the `zipCode` parameter is undefined, the default value is set to `98052`.</span></span> <span data-ttu-id="bad73-116">Se o parâmetro `dayOfWeek` estiver indefinido, ele será definido como Quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="bad73-116">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a><span data-ttu-id="bad73-117">Parâmetros de intervalo</span><span class="sxs-lookup"><span data-stu-id="bad73-117">Range parameters</span></span>

<span data-ttu-id="bad73-118">Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada.</span><span class="sxs-lookup"><span data-stu-id="bad73-118">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="bad73-119">Uma função também pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="bad73-119">A function can also return a range of data.</span></span> <span data-ttu-id="bad73-120">O Excel passará um intervalo de dados de célula como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="bad73-120">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="bad73-121">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="bad73-121">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="bad73-122">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="bad73-122">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="bad73-123">Observe que, nos metadados JSON dessa função, a propriedade do `type` parâmetro é definida como. `matrix`</span><span class="sxs-lookup"><span data-stu-id="bad73-123">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {[][]} values Multiple ranges of values.  
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

## <a name="invocation-parameter"></a><span data-ttu-id="bad73-124">Parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="bad73-124">Invocation parameter</span></span>

<span data-ttu-id="bad73-125">Cada função personalizada é automaticamente passada um `invocation` argumento como o último argumento.</span><span class="sxs-lookup"><span data-stu-id="bad73-125">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="bad73-126">Esse argumento pode ser usado para recuperar contexto adicional, como o endereço da célula de chamada.</span><span class="sxs-lookup"><span data-stu-id="bad73-126">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="bad73-127">Ou pode ser usado para enviar informações para o Excel, como um manipulador de função para [cancelar uma função](custom-functions-web-reqs.md#stream-and-cancel-functions).</span><span class="sxs-lookup"><span data-stu-id="bad73-127">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#stream-and-cancel-functions).</span></span> <span data-ttu-id="bad73-128">Mesmo que você declare nenhum parâmetro, sua função personalizada tem esse parâmetro.</span><span class="sxs-lookup"><span data-stu-id="bad73-128">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="bad73-129">Esse argumento não aparece para um usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="bad73-129">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="bad73-130">Se você deseja usar `invocation` em sua função personalizada, declare-a como o último parâmetro.</span><span class="sxs-lookup"><span data-stu-id="bad73-130">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="bad73-131">No exemplo de código a seguir, `invocation` o contexto é explicitamente declarado para sua referência.</span><span class="sxs-lookup"><span data-stu-id="bad73-131">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

<span data-ttu-id="bad73-132">O parâmetro permite que você obtenha o contexto da célula de invocação, que pode ser útil em alguns cenários, incluindo [a descoberta do endereço de uma célula que invoque uma função personalizada](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="bad73-132">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="bad73-133">Parâmetro de contexto da célula de endereçamento</span><span class="sxs-lookup"><span data-stu-id="bad73-133">Addressing cell's context parameter</span></span>

<span data-ttu-id="bad73-134">Em alguns casos, você precisa obter o endereço da célula que chamou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="bad73-134">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="bad73-135">Isso é útil nos seguintes tipos de cenários:</span><span class="sxs-lookup"><span data-stu-id="bad73-135">This is useful in the following types of scenarios:</span></span>

- <span data-ttu-id="bad73-136">Intervalos de formatação: Use o endereço da célula como a chave para armazenar informações no [Office. armazenamento](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="bad73-136">Formatting ranges: Use the cell's address as the key to store information in [Office.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="bad73-137">Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `Office.storage`.</span><span class="sxs-lookup"><span data-stu-id="bad73-137">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `Office.storage`.</span></span>
- <span data-ttu-id="bad73-138">Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `Office.storage` usando `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="bad73-138">Displaying cached values: If your function is used offline, display stored cached values from `Office.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="bad73-139">Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.</span><span class="sxs-lookup"><span data-stu-id="bad73-139">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="bad73-140">Para solicitar um contexto de uma célula de endereçamento em uma função, você precisa usar uma função para localizar o endereço da célula, como a do exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="bad73-140">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="bad73-141">As informações sobre o endereço de uma célula são expostas apenas se `@requiresAddress` o estiver marcado nos comentários da função.</span><span class="sxs-lookup"><span data-stu-id="bad73-141">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

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

<span data-ttu-id="bad73-142">Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="bad73-142">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="bad73-143">Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="bad73-143">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="see-also"></a><span data-ttu-id="bad73-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="bad73-144">See also</span></span>

* [<span data-ttu-id="bad73-145">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="bad73-145">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="bad73-146">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bad73-146">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="bad73-147">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="bad73-147">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="bad73-148">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bad73-148">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="bad73-149">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="bad73-149">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)