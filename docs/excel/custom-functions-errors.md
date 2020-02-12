---
ms.date: 11/04/2019
description: 'Manipular e retornar erros como #NULL! da sua função personalizada'
title: Manipular e retornar erros da sua função personalizada (visualização)
localization_priority: Normal
ms.openlocfilehash: 19199a56d6699afd013c98c7b117b93528deb304
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950821"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a>Manipular e retornar erros da sua função personalizada (visualização)

> [!NOTE]
> Os recursos descritos neste artigo estão atualmente em visualização, estando sujeitos a alterações. No momento, eles não têm suporte para utilização em ambientes de produção. Você precisará do [Office Insider](https://insider.office.com/join) para experimentar os recursos de visualização.  Uma boa maneira de experimentar recursos de versão prévia é usar uma assinatura do Office 365. Caso você ainda não tenha uma assinatura do Office 365, obtenha uma assinatura do Office 365 gratuita e renovável por 90 dias ingressando no [Programa para Desenvolvedores do Office 365](https://developer.microsoft.com/office/dev-program).

Se algo der errado enquanto sua função personalizada é executada, você precisará retornar um erro para informar o usuário. Se você tiver requisitos de parâmetros específicos, como apenas números positivos, será necessário testar os parâmetros e gerar um erro se eles não estiverem corretos. Você também pode usar um bloco `try`-`catch` para detectar quaisquer erros que ocorram enquanto sua função personalizada é executada.

## <a name="detect-and-throw-an-error"></a>Detectar e lançar um erro

Vamos analisar um caso em que você precisa garantir que um parâmetro de código postal esteja no formato correto para que a função personalizada funcione. A função personalizada a seguir usa uma expressão regular para verificar o CEP. Se este estiver correto, procurará a cidade (em outra função) e retornará o valor. Se não estiver correto, ele retornará um erro `#VALUE!` para a célula.

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

O objeto `CustomFunctions.Error` é usado para retornar um erro de volta à célula. Ao criar o objeto, especifique qual erro você deseja usar usando um dos seguintes valores de enumeração `ErrorCode`.


|Valor de enumeração ErrorCode  |Valor da célula do Excel  |Significado  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | Um valor usado na fórmula é de tipo incorreto. |
|`notAvailable`   | `#N/A`    | A função ou serviço não está disponível. |
|`divisionByZero` | `#DIV/0`  | Esteja ciente de que o JavaScript permite a divisão por zero, portanto, você precisa escrever um manipulador de erros com cuidado para detectar essa condição. |
|`invalidNumber`  | `#NUM!`   | Há um problema com o número usado na fórmula |
|`nullReference`  | `#NULL!`  | Os intervalos na fórmula não se interceptam. |

O exemplo de código a seguir mostra como criar e retornar um erro para um número inválido (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Quando você retorna um erro `#VALUE!`, também pode incluir uma mensagem personalizada que será mostrada em um pop-up quando o usuário passar o mouse sobre a célula. O exemplo a seguir mostra como retornar uma mensagem de erro personalizada.

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, “The parameter can only contain lowercase characters.”);
throw error;
```

## <a name="use-try-catch-blocks"></a>Use blocos try-catch

Em geral, você deve usar blocos `try`-`catch` em sua função personalizada para detectar possíveis erros que ocorram. Se você não tratar exceções no seu código, elas serão retornadas ao Excel. Por padrão, o Excel retorna `#VALUE!` para uma exceção não tratada.

No exemplo de código a seguir, a função personalizada faz uma chamada de busca para um serviço REST. É possível que a chamada falhe, por exemplo, se o serviço REST retornar um erro ou a rede cair. Se isso acontecer, a função personalizada retornará `#N/A` para indicar que a chamada Web falhou.


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
