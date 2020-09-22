---
ms.date: 09/21/2020
description: 'Manipular e retornar erros como #NULL! a partir de sua função personalizada.'
title: Manipular e retornar erros de sua função personalizada
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175525"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Manipular e retornar erros de sua função personalizada

Se algo der errado enquanto sua função personalizada é executada, retorne um erro para informar ao usuário. Se você tiver requisitos de parâmetros específicos, como apenas números positivos, teste os parâmetros e acione um erro se eles não estiverem corretos. Você também pode usar um bloco `try`-`catch` para detectar quaisquer erros que ocorram enquanto sua função personalizada é executada.

## <a name="detect-and-throw-an-error"></a>Detectar e lançar um erro

Vamos examinar um caso em que você precisa garantir que um parâmetro de CEP esteja no formato correto para que a função personalizada funcione. A função personalizada a seguir usa uma expressão regular para verificar o CEP. Se o formato do CEP estiver correto, ele pesquisará a cidade usando outra função e retornará o valor. Se o formato não for válido, a função retornará um `#VALUE!` erro à célula.

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

O objeto [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) é usado para retornar um erro de volta para a célula. Ao criar o objeto, especifique o erro que você deseja usar, escolhendo um dos seguintes valores de `ErrorCode` enumeração.


|Valor de enumeração ErrorCode  |Valor da célula do Excel  |Significado  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | Esteja ciente de que o JavaScript permite a divisão por zero, portanto, você precisa escrever um manipulador de erros com cuidado para detectar essa condição. |
|`invalidName`    | `#NAME?`  | Há um erro de digitação no nome da função. Observe que esse erro é suportado como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada. | 
|`invalidNumber`  | `#NUM!`   | Há um problema com um número na fórmula. |
|`invalidReference` | `#REF!` | A função se refere a uma célula inválida. Observe que esse erro é suportado como um erro de entrada de função personalizada, mas não como um erro de saída de função personalizada.|
|`invalidValue`   | `#VALUE!` | Um valor na fórmula é do tipo incorreto. |
|`notAvailable`   | `#N/A`    | A função ou o serviço não está disponível. |
|`nullReference`  | `#NULL!`  | Os intervalos na fórmula não fazem interseção. |

O exemplo de código a seguir mostra como criar e retornar um erro para um número inválido (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

O `#VALUE!` e os `#N/A` erros também dão suporte a mensagens de erro personalizadas. As mensagens de erro personalizadas são exibidas no menu indicador de erro, que é acessado ao passar o sinalizador de erro em cada célula com um erro. O exemplo a seguir mostra como retornar uma mensagem de erro personalizada com o `#VALUE!` erro.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Use blocos try-catch

Em geral, use `try` - `catch` blocos em sua função personalizada para detectar quaisquer possíveis erros que ocorram. Se você não tratar exceções no seu código, elas serão retornadas ao Excel. Por padrão, o Excel retorna `#VALUE!` para erros ou exceções não manipuladas.

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
* [Requisitos de funções personalizadas](custom-functions-requirement-sets.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
