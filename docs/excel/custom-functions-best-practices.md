---
ms.date: 11/29/2018
description: Saiba mais sobre as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.
title: Práticas recomendadas para funções personalizadas
ms.openlocfilehash: c1be1d01a88d50bb0f3aee8af1aea7c47658bc10
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724883"
---
# <a name="custom-functions-best-practices-preview"></a>Práticas recomendadas para funções personalizadas (versão prévia)

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>Tratamento de erros

Quando criar um suplemento que define funções personalizadas, não deixe de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md). No seguinte exemplo de código, `.catch` tratará os erros que ocorreram anteriormente no código.

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

## <a name="troubleshooting"></a>Solução de problemas

Quando testar o suplemento no Office para Windows, habilite o **[log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** para solucionar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução. O log de tempo de execução grava instruções `console.log` em um arquivo de log para ajudá-lo a descobrir problemas.

Para relatar problemas sobre este método de solução de problemas, envie comentários à equipe de funções personalizadas do Excel. Para fazer isso, selecione **Arquivo | Comentários | Enviar um Rosto Triste**. Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.

## <a name="debugging"></a>Depuração

Atualmente, o método ideal para depuração de funções personalizadas do Excel consiste primeiro em [sideload](../testing/sideload-office-add-ins-for-testing.md) o suplemento no **Excel Online**. Em seguida, para depurar as funções personalizadas, use a [ferramenta de depuração nativa F12 no navegador](../testing/debug-add-ins-in-office-online.md), associado às seguintes técnicas:

- Use as instruções `console.log` no código das funções personalizadas para enviar saída ao console em tempo real.

- Use as instruções `debugger;` no código das funções personalizadas para especificar pontos de interrupção, onde a execução será pausada quando a janela F12 for aberta. Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução será pausada na instrução `debugger;`, o que permite inspecionar manualmente os valores dos parâmetros antes que a função retorne. A instrução `debugger;` não afeta o Excel Online quando a janela F12 não está aberta. Atualmente, a instrução `debugger;` não afeta o Excel para Windows.

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

Se o suplemento não for devidamente registrado, [ verifique se os certificados SSL estão configurados corretamente ](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que hospeda o aplicativo do suplemento.

## <a name="mapping-function-names-to-json-metadata"></a>Como mapear nomes de função para metadados JSON

Conforme descrito no artigo [Visão geral de funções personalizadas](custom-functions-overview.md), um projeto de funções personalizadas deve incluir um arquivo de metadados JSON com as informações necessárias que o Excel exige para registrar as funções personalizadas e disponibilizá-las aos usuários finais. Além disso, no arquivo JavaScript que define as funções personalizadas, você deve fornecer informações para especificar qual objeto de função no arquivo de metadados JSON corresponde a cada função personalizada no arquivo JavaScript.

Por exemplo, o seguinte código de exemplo define a função personalizada `add` e, em seguida, especifica que a função `add` corresponde ao objeto no arquivo de metadados JSON, em que o valor da propriedade `id` seja **ADD**.

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.

* No arquivo JavaScript, especifique os nomes das funções no camelCase. Por exemplo, o nome da função `addTenToInput` é escrito no camelCase: a primeira palavra no nome começa com uma letra minúscula e cada palavra subsequente no nome começa com uma letra maiúscula.

* No arquivo de metadados JSON, especifique o valor de cada propriedade `name` em maiúsculas. A propriedade `name` define o nome da função que os usuários finais verão no Excel. O uso de letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente aos usuários finais do Excel, onde todos os nomes de funções internos são escritos em maiúsculas.

* No arquivo de metadados JSON, especifique o valor de cada propriedade `id` em maiúsculas. Dessa maneira, fica claro qual parte da instrução `CustomFunctionMappings` no código JavaScript corresponde à propriedade `id`, no arquivo de metadados JSON, desde que o nome da função use camelCase, conforme recomendado anteriormente.

* No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos. 

* No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo. Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`. Além disso, não especifique dois valores `id` no arquivo de metadados, que tenham como diferença apenas o uso de maiúsculas e minúsculas. Por exemplo, não defina um objeto de função com um valor `id` de **add** e outro objeto de função com um valor `id` de **ADD**.

* Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente. Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.

* No arquivo JavaScript, especifique todos os mapeamentos de funções personalizadas no mesmo local. Por exemplo, o exemplo de código a seguir define duas funções personalizadas e, em seguida, especifica as informações de mapeamento para ambas.

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

    O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript.

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

## <a name="declaring-optional-parameters"></a>Como declarar parâmetros opcionais 
No Excel para Windows (versão 1812 ou posterior), é possível declarar parâmetros opcionais para suas funções personalizadas. Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes. Por exemplo, uma função `FOO` com um parâmetro obrigatório chamado `parameter1` e parâmetro opcional chamado `parameter2` seria exibida como `=FOO(parameter1, [parameter2])` no Excel.

Para tornar um parâmetro opcional, adicione `"optional": true` ao parâmetro no arquivo JSON de metadados que define a função. O exemplo a seguir mostra o provável aspecto disso para a função `=ADD(first, second, [third])`. Observe que o parâmetro `[third]` opcional segue os dois parâmetros obrigatórios. Os parâmetros obrigatórios aparecerão primeiro na interface do usuário da fórmula do Excel.

```json
{
    "id": "add",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontecerá quando os parâmetros opcionais forem indefinidos. No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`. Se o parâmetro `zipCode` estiver indefinido, o valor padrão será definido como 98052. Se o parâmetro `dayOfWeek` estiver indefinido, ele será definido como Quarta-feira.

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a>Considerações adicionais

Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM. No Excel para Windows, onde as funções personalizadas usam o [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o DOM.

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
