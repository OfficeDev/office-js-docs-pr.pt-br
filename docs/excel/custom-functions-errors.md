---
title: Manipular e retornar erros de sua função personalizada
description: 'Manipular e retornar erros como #NULL! de sua função personalizada.'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: a2f93059f9082bc5a53c07159c9356a41cf16729
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443542"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Manipular e retornar erros de sua função personalizada

Se algo der errado enquanto sua função personalizada é executado, retorne um erro para informar o usuário. Se você tiver requisitos de parâmetro específicos, como apenas números positivos, teste os parâmetros e lance um erro se eles não estão corretos. Você também pode usar um [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) bloco para capturar quaisquer erros que ocorram enquanto sua função personalizada é executado.

## <a name="detect-and-throw-an-error"></a>Detectar e lançar um erro

Vamos ver um caso em que você precisa garantir que um parâmetro de CEP está no formato correto para que a função personalizada funcione. A função personalizada a seguir usa uma expressão regular para verificar o CEP. Se o formato do CEP estiver correto, ele procurará a cidade usando outra função e retornará o valor. Se o formato não for válido, a função retornará um `#VALUE!` erro para a célula.

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

## <a name="the-customfunctionserror-object"></a>O objeto CustomFunctions.Error

O [objeto CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) é usado para retornar um erro à célula. Ao criar o objeto, especifique qual erro você deseja usar escolhendo um dos seguintes valores `ErrorCode` de número.

|Valor de enumeração ErrorCode  |Valor da célula do Excel  |Descrição  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | A função está tentando dividir por zero. |
|`invalidName`    | `#NAME?`  | Há um erro de digitação no nome da função. Observe que esse erro é suportado como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada. |
|`invalidNumber`  | `#NUM!`   | Há um problema com um número na fórmula. |
|`invalidReference` | `#REF!` | A função se refere a uma célula inválida. Observe que esse erro é suportado como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada.|
|`invalidValue`   | `#VALUE!` | Um valor na fórmula é do tipo errado. |
|`notAvailable`   | `#N/A`    | A função ou serviço não está disponível. |
|`nullReference`  | `#NULL!`  | Os intervalos na fórmula não se cruzam. |

O exemplo de código a seguir mostra como criar e retornar um erro para um número inválido (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Os `#VALUE!` erros e também `#N/A` suportam mensagens de erro personalizadas. As mensagens de erro personalizadas são exibidas no menu indicador de erro, que é acessado ao passar o mouse sobre o sinalizador de erro em cada célula com um erro. O exemplo a seguir mostra como retornar uma mensagem de erro personalizada com o `#VALUE!` erro.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>Manipular erros ao trabalhar com matrizes dinâmicas

Além de retornar um único erro, uma função personalizada pode saída de uma matriz dinâmica que inclui um erro. Por exemplo, uma função personalizada poderia fazer a saída da matriz `[1],[#NUM!],[3]` . O exemplo de código a seguir mostra como inserir três parâmetros em uma função personalizada, substituir um dos parâmetros de entrada por um erro e retornar uma matriz bidimensional com os resultados do processamento de cada parâmetro `#NUM!` de entrada.

```js
/**
* Returns the #NUM! error as part of a 2-dimensional array.
* @customfunction
* @param {number} first First parameter.
* @param {number} second Second parameter.
* @param {number} third Third parameter.
* @returns {number[][]} Three results, as a 2-dimensional array.
*/
function returnInvalidNumberError(first, second, third) {
  // Use the `CustomFunctions.Error` object to retrieve an invalid number error.
  var error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  var firstResult = first;
  var secondResult =  error;
  var thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>Erros como entradas de função personalizadas

Uma função personalizada pode ser avaliada mesmo se o intervalo de entrada contiver um erro. Por exemplo, uma função personalizada pode usar o intervalo **A2:A7** como uma entrada, mesmo que **A6:A7** contenha um erro.

Para processar entradas que contêm erros, uma função personalizada deve ter a propriedade de metadados JSON `allowErrorForDataTypeAny` definida como `true` . Confira [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md#metadata-reference) para obter mais informações.

> [!IMPORTANT]
> A `allowErrorForDataTypeAny` propriedade só pode ser usada com [metadados JSON criados manualmente.](custom-functions-json.md) Essa propriedade não funciona com o processo de metadados JSON gerado automaticamente.

## <a name="use-trycatch-blocks"></a>Usar `try...catch` blocos

Em geral, use [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) blocos em sua função personalizada para capturar possíveis erros que ocorram. Se você não lidar com exceções em seu código, elas serão retornadas para Excel. Por padrão, Excel retorna para erros ou exceções não a `#VALUE!` mão.

No exemplo de código a seguir, a função personalizada faz uma chamada de busca para um serviço REST. É possível que a chamada falhe, por exemplo, se o serviço REST retornar um erro ou a rede cair. Se isso acontecer, a função personalizada retornará `#N/A` para indicar que a chamada da Web falhou.

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

## <a name="next-steps"></a>Próximas etapas

Saiba como [solucionar problemas com as suas funções personalizadas](custom-functions-troubleshooting.md).

## <a name="see-also"></a>Confira também

* [Depuração de funções personalizadas](custom-functions-debugging.md)
* [Conjuntos de requisitos de funções personalizadas](../reference/requirement-sets/custom-functions-requirement-sets.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
