---
ms.date: 09/20/2018
description: Saiba mais sobre melhores práticas e padrões recomendados para funções personalizadas do Excel.
title: Práticas recomendadas de funções personalizadas
ms.openlocfilehash: 1f2c0a80e62b65523fcc1673ba2ca4be444e6ce0
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068807"
---
# <a name="custom-functions-best-practices"></a>Práticas recomendadas de funções personalizadas

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.

## <a name="error-handling"></a>Manipulação de erros

Quando você cria um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erros para considerar os erros de tempo de execução. Em geral, a manipulação de erros para funções personalizadas é a mesma que [a manipulação de erros para a API JavaScript do Excel](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` manipulará os erros que ocorram anteriormente no código.

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="error-logging"></a>Log de erros

Você pode ativar o log de erros para o suplemento de funções personalizadas de várias maneiras, como: 

- [Use o log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) para depurar o arquivo de manifesto XML do suplemento. 

- Use `console.log` instruções dentro do seu código de funções personalizadas para enviar a saída para o console em tempo real.

> [!NOTE]
> O log de tempo de execução está atualmente disponível apenas para a área de trabalho do Office 2016.

## <a name="debugging"></a>Depuração

Atualmente, o melhor método para depurar as funções personalizadas do Excel é usar o [Excel Online](https://www.office.com/launch/excel) e usar a ferramenta de depuração F12 nativa em seu navegador. Outras ferramentas de depuração para funções personalizadas podem estar disponíveis no futuro.

## <a name="mapping-names"></a>Nomes de mapeamento

Por padrão, o nome de uma função personalizada no seu arquivo JavaScript geralmente é declarado usando letras maiúsculas e corresponde exatamente ao nome da função que os usuários finais veem no Excel. No entanto, você pode alterar isso usando o `CustomFunctionsMappings` objeto para mapear um ou mais nomes das funções do arquivo JavaScript para diferentes valores que os usuários finais verão como nomes de função no Excel. Embora você não precisa usar `CustomFunctionsMapping`, pode ser útil se estiver usando uma sintaxe uglifier, webpack ou de importação - todas as quais têm dificuldade com nomes de função em letras maiúsculas.
  
O exemplo de código a seguir define um único par chave-valor que mapeia o nome da função JavaScript `plusFortyTwo` para o `ADD42` nome da função na interface do usuário do Excel. Quando o usuário final escolhe a função `ADD42` no Excel, a função `plusFortyTwo` JavaScript será executada.

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

O exemplo de código a seguir define dois pares chave-valor. O primeiro par mapeia o nome da função JavaScript `plusFifty` para o `ADD50` nome da função na interface do usuário do Excel, e o segundo par mapeia o nome da função JavaScript `plusOneHundred` para o `ADD100` nome da função na interface do usuário do Excel. Quando o usuário final escolhe a função `ADD50` no Excel, a função `plusFifty` JavaScript será executada. Quando o usuário final escolhe a função `ADD100` no Excel, a função `plusOneHundred` JavaScript será executada.

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução para funções personalizadas do Excel](custom-functions-runtime.md)