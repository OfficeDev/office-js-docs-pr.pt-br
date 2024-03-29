---
title: Manipular e retornar erros de sua função personalizada
description: 'Manipular e retornar erros como #NULL! de sua função personalizada.'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c93c13aac1457e776ba8441565c11a23074a8d97
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958563"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Manipular e retornar erros de sua função personalizada

Se algo der errado enquanto sua função personalizada é executada, retorne um erro para informar o usuário. Se você tiver requisitos de parâmetro específicos, como apenas números positivos, teste os parâmetros e gere um erro se eles não estiverem corretos. Você também pode usar um bloco [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) para capturar todos os erros que ocorrem enquanto sua função personalizada é executada.

## <a name="detect-and-throw-an-error"></a>Detectar e lançar um erro

Vamos examinar um caso em que você precisa garantir que um parâmetro de cep esteja no formato correto para que a função personalizada funcione. A função personalizada a seguir usa uma expressão regular para verificar o CEP. Se o formato de cep estiver correto, ele pesquisará a cidade usando outra função e retornará o valor. Se o formato não for válido, a função retornará um `#VALUE!` erro para a célula.

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

O [objeto CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) é usado para retornar um erro de volta à célula. Ao criar o objeto, especifique qual erro você deseja usar escolhendo um dos valores `ErrorCode` de enumeração a seguir.

|Valor de enumeração ErrorCode  |Valor da célula do Excel  |Descrição  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | A função está tentando dividir por zero. |
|`invalidName`    | `#NAME?`  | Há um erro de digitação no nome da função. Observe que esse erro tem suporte como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada. |
|`invalidNumber`  | `#NUM!`   | Há um problema com um número na fórmula. |
|`invalidReference` | `#REF!` | A função refere-se a uma célula inválida. Observe que esse erro tem suporte como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada.|
|`invalidValue`   | `#VALUE!` | Um valor na fórmula é do tipo errado. |
|`notAvailable`   | `#N/A`    | A função ou serviço não está disponível. |
|`nullReference`  | `#NULL!`  | Os intervalos na fórmula não são interseccionados. |

O exemplo de código a seguir mostra como criar e retornar um erro para um número inválido (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Os `#VALUE!` erros e `#N/A` os erros também dão suporte a mensagens de erro personalizadas. Mensagens de erro personalizadas são exibidas no menu do indicador de erro, que é acessado passando o mouse sobre o sinalizador de erro em cada célula com um erro. O exemplo a seguir mostra como retornar uma mensagem de erro personalizada com o `#VALUE!` erro.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>Tratar erros ao trabalhar com matrizes dinâmicas

Além de retornar um único erro, uma função personalizada pode gerar uma matriz dinâmica que inclui um erro. Por exemplo, uma função personalizada pode gerar a saída da matriz `[1],[#NUM!],[3]`. O exemplo de código a seguir mostra como inserir três parâmetros em uma função personalizada, substituir um dos parâmetros `#NUM!` de entrada por um erro e retornar uma matriz bidimensional com os resultados do processamento de cada parâmetro de entrada.

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
  const error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  const firstResult = first;
  const secondResult =  error;
  const thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>Erros como entradas de função personalizadas

Uma função personalizada pode ser avaliada mesmo se o intervalo de entrada contiver um erro. Por exemplo, uma função personalizada pode usar o intervalo **A2:A7** como uma entrada, mesmo que **A6:A7** contenha um erro.

Para processar entradas que contêm erros, uma função personalizada deve ter a propriedade de metadados `allowErrorForDataTypeAny` JSON definida como `true`. Consulte [Criar manualmente metadados JSON para funções personalizadas](custom-functions-json.md#metadata-reference) para obter mais informações.

> [!IMPORTANT]
> A `allowErrorForDataTypeAny` propriedade só pode ser usada com [metadados JSON criados manualmente](custom-functions-json.md). Essa propriedade não funciona com o processo de metadados JSON gerado automaticamente.

## <a name="use-trycatch-blocks"></a>Blocos de `try...catch` uso

Em geral, use [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) blocos em sua função personalizada para capturar possíveis erros que ocorram. Se você não tratar exceções em seu código, elas serão retornadas ao Excel. Por padrão, o Excel retorna `#VALUE!` para erros ou exceções sem tratamento.

No exemplo de código a seguir, a função personalizada faz uma chamada de busca para um serviço REST. É possível que a chamada falhe, por exemplo, se o serviço REST retornar um erro ou a rede cair. Se isso acontecer, a função personalizada retornará para `#N/A` indicar que a chamada à Web falhou.

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
* [Conjuntos de requisitos de funções personalizadas](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
