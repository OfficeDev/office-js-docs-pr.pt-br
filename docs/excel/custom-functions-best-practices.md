---
ms.date: 09/20/2018
description: Saiba mais sobre melhores práticas e padrões recomendados para funções personalizadas do Excel.
title: Práticas recomendadas de funções personalizadas
ms.openlocfilehash: 4fe0ddc36ce1b08ea360bb556121e76cd57c3823
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004907"
---
# <a name="custom-functions-best-practices"></a>Práticas recomendadas de funções personalizadas

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.

## <a name="error-handling"></a>Manipulação de erro

Ao criar um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erro para considerar os erros no tempo de execução. O tratamento de erros de funções personalizadas é o mesmo que [o tratamento de erros para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` manipulará os erros que ocorreram anteriormente no código.

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi.com/comments/" + x; 
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

## <a name="debugging"></a>Depuração
Atualmente, o melhor método para depurar funções personalizadas do Excel é primeiro [fazer sideload](../testing/sideload-office-add-ins-for-testing.md) do seu suplemento no **Excel Online**. Você pode depurar suas funções personalizadas usando [F12, a ferramenta de depuração nativa do seu navegador](../testing/debug-add-ins-in-office-online.md). Use `console.log` instruções dentro do seu código de funções personalizadas para enviar a saída para o console em tempo real.

Se seu suplemento falhar ao registrar, [verifique se os certificados SSL estão configurados corretamente](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que está hospedando o seu aplicativo de suplemento.

Se você estiver testando seu suplemento de área de trabalho do Office 2016, é possível habilitar o [log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) para depurar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução. 


## <a name="mapping-names"></a>Nomes de mapeamento

Por padrão, o nome de uma função personalizada no seu arquivo JavaScript geralmente é declarado usando letras maiúsculas e corresponde exatamente ao nome da função que os usuários finais veem no Excel. No entanto, você pode alterar isso usando o `CustomFunctionsMappings` objeto para mapear um ou mais nomes das funções do arquivo JavaScript para diferentes valores que os usuários finais verão como nomes de função no Excel. Isso é útil se você estiver usando um uglifier, webpack ou sintaxe de importação - todas eles têm dificuldade com nomes de função em letras maiúsculas. `CustomFunctionsMappings` é opcional, possivelmente, para projetos que usam JavaScript, mas deve ser usado se o seu projeto usa TypeScript.  
  
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

- [Criar funções personalizadas no Excel](custom-functions-overview.md)
- [Metadados de funções personalizadas](custom-functions-json.md)
- [Tempo de execução para funções personalizadas do Excel](custom-functions-runtime.md)
