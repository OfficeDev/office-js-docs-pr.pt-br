---
ms.date: 09/27/2018
description: Saiba mais sobre melhores práticas e padrões recomendados para funções personalizadas do Excel.
title: Práticas recomendadas para funções personalizadas
ms.openlocfilehash: d157464a3a8bf453cd0970281f1a4fdd27df5d25
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348784"
---
# <a name="custom-functions-best-practices-preview"></a>Práticas recomendadas para funções personalizadas (versão prévia)

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>Manipulação de erro

Ao criar um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erro para considerar os erros no tempo de execução. O tratamento de erros de funções personalizadas é o mesmo que [tratamento de erros para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` tratará os erros que ocorreram anteriormente no código.

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
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

Atualmente, o melhor método para depuração de funções personalizadas do Excel é primeiro fazer o [sideload](../testing/sideload-office-add-ins-for-testing.md) do seu suplemento no **Excel Online**. Dessa forma você pode depurar as funções personalizadas usando a [ferramenta de depuração F12 nativa do navegador](../testing/debug-add-ins-in-office-online.md) em combinação com as técnicas a seguir:

- Use instruções `console.log` no seu código de funções personalizadas para enviar a saída para o console em tempo real.

- Use instruções `debugger;` no seu código de funções personalizadas para especificar pontos de interrupção onde a execução será interrompida quando a janela F12 estiver aberta. Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução será interrompida na instrução `debugger;`, permitindo que você inspecione manualmente os valores dos parâmetros antes que a função retorne. A instrução `debugger;` não tem efeito no Excel Online quando a janela F12 não está aberta. Atualmente, a instrução `debugger;` não tem efeito no Excel para Windows.

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

Se seu suplemento falhar ao registrar, [verifique se os certificados SSL estão configurados corretamente](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que está hospedando o seu aplicativo de suplemento.

Se você estiver testando seu suplemento de área de trabalho do Office 2016, é possível habilitar o [log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) para depurar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.

## <a name="mapping-function-names-to-json-metadata"></a>Mapeamento de nomes de função para metadados JSON

Conforme descrito no artigo de [Visão geral de funções personalizadas](custom-functions-overview.md), um projeto de funções personalizadas deve incluir um arquivo de metadados JSON que fornece as informações exigidas pelo Excel para registrar as funções personalizadas e torná-las disponíveis aos usuários finais. Além disso, dentro do arquivo JavaScript que define as funções personalizadas, você deve fornecer informações para especificar qual objeto de função no arquivo de metadados JSON corresponde a cada função personalizada no arquivo JavaScript.

Por exemplo, o exemplo de código a seguir define a função personalizada `add` e, em seguida, especifica que a função `add` corresponde ao objeto no arquivo de metadados JSON onde o valor da propriedade `id` é **ADD**.

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

Tenha em mente as seguintes práticas recomendadas ao criar funções personalizadas no seu arquivo JavaScript e especificar informações correspondentes no arquivo de metadados JSON.

* No arquivo JavaScript, especifique os nomes de função em camelCase. Por exemplo, o nome da função `addTenToInput` está escrito em camelCase: a primeira palavra no nome começa com uma letra minúscula e cada palavra subsequente no nome começa com uma letra maiúscula.

* No arquivo de metadados JSON, especifique o valor de cada propriedade`name` em letras maiúsculas. A propriedade `name` define o nome da função que os usuários finais verão no Excel. Usar letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente para o usuário do Excel, pois os nomes de todas as funções internas estão em letras maiúsculas.

* No arquivo de metadados JSON, especifique o valor de cada propriedade`id` em letras maiúsculas. Isso torna óbvio qual parte da instrução `CustomFunctionMappings` no seu código JavaScript corresponde à propriedade `id` no arquivo de metadados JSON (desde que o seu nome de função use camelCase, conforme recomendado anteriormente).

* No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` é exclusivo dentro do escopo do arquivo. Ou seja, não deve haver dois objetos de função no arquivo de metadados com o mesmo valor `id`. Além disso, não especifique dois valores `id` no arquivo de metadados que se distinguam apenas por letras maiúsculas ou minúsculas. Por exemplo, não defina um objeto de função com um `id` valor de **add** e outro objeto de função com um `id` valor de **ADD**.

* Não altere o valor de uma propriedade `id` no arquivo de metadados JSON depois que ela tiver sido mapeada para um nome de função JavaScript correspondente. Você pode alterar o nome da função que os usuários finais veem no Excel, atualizando a propriedade `name` dentro do arquivo de metadados JSON, mas você nunca deve alterar o valor de uma propriedade `id` depois que ele for estabelecido.

* No arquivo JavaScript, especifique todos os mapeamentos da função personalizada no mesmo local. Por exemplo, o exemplo de código a seguir define duas funções personalizadas e, em seguida, especifica a informação de mapeamento de ambas as funções.

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas neste exemplo de código JavaScript.

    ```json
    {
      "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
      "functions": [
        {
          "id": "ADD",
          "name": "ADD",
          ...
        },
        {
          "id": "INCREMENT",
          "name": "INCREMENT",
          ...
        }
      ]
    }
    ```

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Runtime de funções personalizadas do Excel](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
