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
# <a name="custom-functions-parameter-options"></a>Opções de parâmetros de funções personalizadas

Funções personalizadas são configuráveis com muitas opções diferentes para parâmetros.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Parâmetros opcionais

Enquanto parâmetros regulares são necessários, os parâmetros opcionais não. Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes. No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número. Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.

#### <a name="javascripttabjavascript"></a>[JavaScript](#tab/javascript)

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

#### <a name="typescripttabtypescript"></a>[TypeScript](#tab/typescript)

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
> Quando nenhum valor é especificado para um parâmetro opcional, o Excel atribui a ele o valor `null`. Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado. Portanto, não use a sintaxe `function add(first:number, second:number, third=0):number` porque ela não será inicializada `third` como 0. Em vez disso, use a sintaxe do TypeScript, conforme mostrado no exemplo anterior.

Ao definir uma função que contenha um ou mais parâmetros opcionais, você deve especificar o que acontece quando os parâmetros opcionais são nulos. No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`. Se o `zipCode` parâmetro for NULL, o valor padrão será definido como `98052`. Se o `dayOfWeek` parâmetro for NULL, ele será definido como quarta-feira.

#### <a name="javascripttabjavascript"></a>[JavaScript](#tab/javascript)

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

#### <a name="typescripttabtypescript"></a>[TypeScript](#tab/typescript)

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

## <a name="range-parameters"></a>Parâmetros de intervalo

Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada. Uma função também pode retornar um intervalo de dados. O Excel passará um intervalo de dados de célula como uma matriz bidimensional.

Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel. A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`. Observe que, nos metadados JSON dessa função, a propriedade do `type` parâmetro é definida como. `matrix`

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

## <a name="invocation-parameter"></a>Parâmetro de invocação

Cada função personalizada é automaticamente passada um `invocation` argumento como o último argumento. Esse argumento pode ser usado para recuperar contexto adicional, como o endereço da célula de chamada. Ou pode ser usado para enviar informações para o Excel, como um manipulador de função para [cancelar uma função](custom-functions-web-reqs.md#make-a-streaming-function). Mesmo que você declare nenhum parâmetro, sua função personalizada tem esse parâmetro. Esse argumento não aparece para um usuário no Excel. Se você deseja usar `invocation` em sua função personalizada, declare-a como o último parâmetro.

No exemplo de código a seguir, `invocation` o contexto é explicitamente declarado para sua referência.

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

O parâmetro permite que você obtenha o contexto da célula de invocação, que pode ser útil em alguns cenários, incluindo [a descoberta do endereço de uma célula que invoque uma função personalizada](#addressing-cells-context-parameter).

### <a name="addressing-cells-context-parameter"></a>Parâmetro de contexto da célula de endereçamento

Em alguns casos, você precisa obter o endereço da célula que chamou sua função personalizada. Isso é útil nos seguintes cenários:

- Intervalos de formatação: Use o endereço da célula como a chave para armazenar informações no [OfficeRuntime. armazenamento](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `OfficeRuntime.storage`.
- Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `OfficeRuntime.storage` usando `onCalculated`.
- Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.

Para solicitar um contexto de uma célula de endereçamento em uma função, você precisa usar uma função para localizar o endereço da célula, como a do exemplo a seguir. As informações sobre o endereço de uma célula são expostas apenas se `@requiresAddress` o estiver marcado nos comentários da função.

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

Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`. Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.

## <a name="next-steps"></a>Próximas etapas
Saiba como [salvar o estado em suas funções personalizadas](custom-functions-save-state.md) ou usar [valores voláteis em suas funções personalizadas](custom-functions-volatile.md).

## <a name="see-also"></a>Confira também

* [Receber e tratar dados com funções personalizadas](custom-functions-web-reqs.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
