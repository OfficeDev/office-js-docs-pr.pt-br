---
ms.date: 05/30/2019
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos do Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: 7bc907157810ce88330fe41b21ca6ff115525491
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706054"
---
# <a name="custom-functions-parameter-options"></a>Opções de parâmetros de funções personalizadas

As funções personalizadas são configuráveis com muitas opções diferentes para parâmetros:
- [Parâmetros opcionais](#custom-functions-optional-parameters)
- [Parâmetros de intervalo](#range-parameters)
- [Parâmetro de contexto de invocação](#invocation-parameter)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-optional-parameters"></a>Parâmetros opcionais de funções personalizadas

Enquanto parâmetros regulares são necessários, os parâmetros opcionais não. Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes. No exemplo a seguir, a função Add pode opcionalmente adicionar um terceiro número. Essa função aparece como `=CONTOSO.ADD(first, second, [third])` no Excel.

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
  if (third !== undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontecerá quando os parâmetros opcionais forem indefinidos. No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`. Se o `zipCode` parâmetro estiver indefinido, o valor padrão será definido como `98052`. Se o parâmetro `dayOfWeek` estiver indefinido, ele será definido como Quarta-feira.

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
