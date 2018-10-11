---
ms.date: 10/03/2018
description: Saiba mais sobre melhores práticas e padrões recomendados para funções personalizadas do Excel.
title: Práticas recomendadas para funções personalizadas
ms.openlocfilehash: f6781de97f912df70800532032162187ae9f9344
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459109"
---
# <a name="custom-functions-best-practices-preview"></a>Práticas recomendadas para funções personalizadas (versão prévia)

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>Manipulação de erro

Ao construir um suplemento que define as funções personalizadas, certifique-se de incluir lógica para manipulação de erro para lidar com erros em tempo de execução. Manipulação de erro para funções personalizadas é o mesmo que [manipulação de erros para a API JavaScript do Excel, de maneira geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` manipulará quaisquer erros que ocorram anteriormente no código.

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

Atualmente, o melhor método para depurar funções personalizadas do Excel é primeiro [fazer sideload](../testing/sideload-office-add-ins-for-testing.md) do seu suplemento no **Excel Online**. Em seguida, você pode depurar suas funções personalizadas, usando a [ferramenta de depuração F12 nativa do seu navegador](../testing/debug-add-ins-in-office-online.md) em combinação com as técnicas a seguir:

- Use `console.log` instruções dentro do seu código de funções personalizadas para enviar a saída para o console em tempo real.

- Use `debugger;` instruções dentro de seu código de funções personalizadas para especificar os pontos de interrupção onde a execução fará uma pausa quando a janela F12 estiver aberta. Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução fará uma pausa sobre a instrução `debugger;` , permitindo que você inspecione manualmente os valores de parâmetro antes que a função retorne. A instrução `debugger;` não tem efeito no Excel Online quando a janela F12 não estiver aberta. Atualmente, a instrução `debugger;` não tem efeito no Excel para Windows.

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

Se seu suplemento falhar ao registrar, [verifique se os certificados SSL estão configurados corretamente](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que está hospedando o seu aplicativo de suplemento.

Se você estiver testando seu suplemento no Office na área de trabalho do Windows, é possível habilitar o [log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) para depurar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.

## <a name="mapping-function-names-to-json-metadata"></a>Mapeamento de nomes de função para metadados JSON

Conforme descrito no artigo [Visão geral de funções personalizadas](custom-functions-overview.md) , um projeto de funções personalizadas deve incluir um arquivo de metadados JSON que forneça as informações exigidas pelo Excel para registrar as funções personalizadas e torná-las disponíveis aos usuários finais. Além disso, dentro do arquivo JavaScript que define suas funções personalizadas, você deve fornecer informações para especificar qual objeto de função no arquivo de metadados JSON corresponde a cada função personalizada no arquivo JavaScript.

Por exemplo, o exemplo de código a seguir define a função personalizada `add` e, em seguida, especifica que a função `add` corresponde ao objeto no arquivo de metadados JSON onde o valor da propriedade `id` é **ADD**.

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

Tenha em mente as seguintes práticas recomendadas ao criar funções personalizadas no seu arquivo JavaScript e especificar informações correspondentes no arquivo de metadados JSON.

* No arquivo JavaScript, especifique nomes de função em camelCase. Por exemplo, o nome da função `addTenToInput` está escrito em camelCase: a primeira palavra no nome começa com uma letra minúscula e cada palavra subsequente no nome começa com uma letra maiúscula.

* No arquivo de metadados JSON, especifique o valor de cada propriedade `name` em letras maiusculas. A propriedade `name` define o nome da função que os usuários finais verão no Excel. Usar letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente para usuários finais no Excel, onde todos os nomes de função interna estão em letras maiúsculas.

* No arquivo de metadados JSON, especifique o valor de cada propriedade `id` em letras maiúsculas. Isso torna óbvio qual parte da instrução `CustomFunctionMappings` em seu código JavaScript corresponde à propriedade `id` no arquivo de metadados JSON (desde que o seu nome de função use camelCase, conforme recomendado anteriormente).

* No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` é exclusivo dentro do escopo do arquivo. Ou seja, não deve haver dois objetos de função no arquivo de metadados com o mesmo valor `id` . Além disso, não especifique dois valores `id` no arquivo de metadados que diferem somente por caso. Por exemplo, não defina um objeto de função com um valor `id` de **add** e outro objeto de função com um valor `id` de **ADD**.

* Não altere o valor de uma propriedade `id` no arquivo de metadados JSON depois que ela foi mapeada para um nome de função JavaScript correspondente. Você pode alterar o nome da função que os usuários finais veem no Excel, atualizando a propriedade `name` dentro do arquivo de metadados JSON, mas você nunca deve alterar o valor de uma propriedade `id` depois que ela foi estabelecida.

* No arquivo JavaScript, especifique todos os mapeamentos de função personalizada no mesmo local. Por exemplo, o exemplo de código a seguir define duas funções personalizadas e especifica as informações de mapeamento para ambas as funções.

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

## <a name="additional-considerations"></a>Considerações adicionais

Para criar um suplemento que possa ser executado em múltiplas plataformas (um dos locatários chaves de Suplementos do Office), você não deve acessar o Document Object Model (DOM) em funções personalizadas ou usar bibliotecas como a jQuery que dependem do DOM. No Excel para Windows, onde as funções personalizadas usam o  [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o DOM.

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Runtime de funções personalizadas do Excel](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
