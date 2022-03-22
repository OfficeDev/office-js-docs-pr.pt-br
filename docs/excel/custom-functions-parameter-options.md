---
ms.date: 03/08/2021
description: Saiba como usar parâmetros diferentes em suas funções personalizadas, como intervalos Excel, parâmetros opcionais, contexto de invocação e muito mais.
title: Opções para Excel funções personalizadas
ms.localizationpriority: medium
ms.openlocfilehash: 2cc0c825932afe3a70d0f9ab6483327051c199fd
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711018"
---
# <a name="custom-functions-parameter-options"></a>Opções de parâmetro de funções personalizadas

Funções personalizadas são configuráveis com muitas opções de parâmetros diferentes.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Parâmetros opcionais

Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes. No exemplo a seguir, a função add pode, opcionalmente, adicionar um terceiro número. Essa função aparece como `=CONTOSO.ADD(first, second, [third])` Excel.

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

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

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

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
> Quando nenhum valor é especificado para um parâmetro opcional, Excel atribui a ele o valor `null`. Isso significa que os parâmetros inicializados por padrão no TypeScript não funcionarão conforme o esperado. Não use a sintaxe porque `function add(first:number, second:number, third=0):number` ela não será inicializada `third` como 0. Em vez disso, use a sintaxe TypeScript conforme mostrado no exemplo anterior.

Quando você definir uma função que contém um ou mais parâmetros opcionais, especifique o que acontece quando os parâmetros opcionais são nulos. No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`. Se o `zipCode` parâmetro for nulo, o valor padrão será definido como `98052`. Se o `dayOfWeek` parâmetro for nulo, ele será definido como quarta-feira.

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

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

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

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

## <a name="range-parameters"></a>Parâmetros de intervalo

Sua função personalizada pode aceitar um intervalo de dados de célula como um parâmetro de entrada. Uma função também pode retornar um intervalo de dados. Excel passará um intervalo de dados de célula como uma matriz bidimensional.

Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel. A função a seguir aceita o `values`parâmetro , e a sintaxe JSDOC `number[][]` `dimensionality` `matrix` define a propriedade do parâmetro como nos metadados JSON para essa função. 

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

## <a name="repeating-parameters"></a>Parâmetros repetidos

Um parâmetro repetido permite que um usuário insira uma série de argumentos opcionais em uma função. Quando a função é chamada, os valores são fornecidos em uma matriz para o parâmetro. Se o nome do parâmetro terminar com um número, o número de cada argumento aumentará incrementalmente, como `ADD(number1, [number2], [number3],…)`. Isso corresponde à convenção usada para funções de Excel de usuário.

A função a seguir soma o total de números, endereços de célula, bem como intervalos, se inseridos.

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

Esta função mostra `=CONTOSO.ADD([operands], [operands]...)` na Excel de trabalho.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>Parâmetro de valor único repetido

Um parâmetro de valor único que se repete permite que vários valores individuais sejam passados. Por exemplo, o usuário pode inserir ADD(1,B2,3). O exemplo a seguir mostra como declarar um parâmetro de valor único.

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

### <a name="single-range-parameter"></a>Parâmetro de intervalo único

Um parâmetro de intervalo único não é tecnicamente um parâmetro repetidor, mas é incluído aqui porque a declaração é muito semelhante a parâmetros repetidos. Ele aparecerá para o usuário como ADD(A2:B3) onde um único intervalo é passado de Excel. O exemplo a seguir mostra como declarar um único parâmetro de intervalo.

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

### <a name="repeating-range-parameter"></a>Parâmetro de intervalo de repetição

Um parâmetro de intervalo repetido permite que vários intervalos ou números sejam passados. Por exemplo, o usuário pode inserir ADD(5,B2,C3,8,E5:E8). Intervalos repetidos geralmente são especificados com o tipo `number[][][]` como são matrizes tridimensionais. Para um exemplo, consulte o exemplo principal listado para [parâmetros repetidos](#repeating-parameters).


### <a name="declaring-repeating-parameters"></a>Declarando parâmetros repetidos
Em Typescript, indique que o parâmetro é multidimensional. Por exemplo,  `ADD(values: number[])` indicaria uma matriz unidimensional, `ADD(values:number[][])` indicaria uma matriz bidimensional e assim por diante.

Em JavaScript, use `@param values {number[]}` para matrizes unidimensionais, `@param <name> {number[][]}` para matrizes bidimensionais e assim por diante para mais dimensões.

Para JSON de autoria manual, `"repeating": true` verifique se o parâmetro está especificado como no arquivo JSON, bem como verifique se seus parâmetros estão marcados como `"dimensionality": matrix`.

## <a name="invocation-parameter"></a>Parâmetro Invocation

Todas as funções personalizadas são automaticamente passadas como `invocation` o último parâmetro de entrada, mesmo que não seja declarada explicitamente. Esse `invocation` parâmetro corresponde ao [objeto Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) . O `Invocation` objeto pode ser usado para recuperar contexto adicional, como o endereço da célula que invocou sua função personalizada. Para acessar o `Invocation` objeto, você deve declarar `invocation` como o último parâmetro em sua função personalizada. 

> [!NOTE]
> O `invocation` parâmetro não aparece como um argumento de função personalizado para usuários no Excel.

O exemplo a seguir mostra como usar o `invocation` parâmetro para retornar o endereço da célula que invocou sua função personalizada. Este exemplo usa a [propriedade address](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member) do `Invocation` objeto. Para acessar o `Invocation` objeto, primeiro declare `CustomFunctions.Invocation` como um parâmetro em seu JSDoc. Em seguida, declare `@requiresAddress` em seu JSDoc para acessar a `address` propriedade do `Invocation` objeto. Por fim, dentro da função, recupere e retorne a `address` propriedade. 

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

Em Excel, uma função personalizada chamando a `address` `Invocation` `SheetName!RelativeCellAddress` propriedade do objeto retornará o endereço absoluto após o formato na célula que invocou a função. Por exemplo, se o parâmetro de entrada estiver localizado em uma planilha chamada **Prices** na célula F6, o valor de endereço do parâmetro retornado será `Prices!F6`. 

O `invocation` parâmetro também pode ser usado para enviar informações para Excel. Confira [Fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function) para saber mais.

## <a name="detect-the-address-of-a-parameter"></a>Detectar o endereço de um parâmetro

Em combinação com o [parâmetro invocação](#invocation-parameter), você pode usar o [objeto Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) para recuperar o endereço de um parâmetro de entrada de função personalizada. Quando invocada, a [propriedade parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) `Invocation` do objeto permite que uma função retorne os endereços de todos os parâmetros de entrada. 

Isso é útil em cenários em que os tipos de dados de entrada podem variar. O endereço de um parâmetro de entrada pode ser usado para verificar o formato de número do valor de entrada. O formato de número pode ser ajustado antes da entrada, se necessário. O endereço de um parâmetro de entrada também pode ser usado para detectar se o valor de entrada tem propriedades relacionadas que podem ser relevantes para cálculos subsequentes. 

>[!NOTE]
> Se você estiver trabalhando com metadados [JSON criados manualmente](custom-functions-json.md) para retornar endereços de parâmetro em vez do gerador [Yeoman para os complementos do Office](../develop/yeoman-generator-overview.md), `options` `requiresParameterAddresses` `true`o objeto deve ter a propriedade definida como , `result` `dimensionality` e o objeto deve ter a propriedade definida como .`matrix`

A função personalizada a seguir recebe três parâmetros de entrada, `parameterAddresses` `Invocation` recupera a propriedade do objeto para cada parâmetro e retorna os endereços. 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array. 
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

Quando uma função personalizada chamando a `parameterAddresses` propriedade é executado, o endereço do parâmetro `SheetName!RelativeCellAddress` é retornado seguindo o formato na célula que invocou a função. Por exemplo, se o parâmetro de entrada estiver localizado em uma planilha chamada **Custos** na célula D8, o valor de endereço do parâmetro retornado será `Costs!D8`. Se a função personalizada tiver vários parâmetros e mais de um endereço de parâmetro for retornado, os endereços retornados serão respingados em várias células, descendentes verticalmente da célula que invocou a função. 

## <a name="next-steps"></a>Próximas etapas

Saiba como usar valores [voláteis em suas funções personalizadas](custom-functions-volatile.md).

## <a name="see-also"></a>Confira também

* [Receber e tratar dados com funções personalizadas](custom-functions-web-reqs.md)
* [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
* [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
