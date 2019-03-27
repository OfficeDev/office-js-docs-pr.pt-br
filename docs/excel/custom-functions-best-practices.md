---
ms.date: 01/08/2019
description: Saiba mais sobre as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.
title: Práticas recomendadas para funções personalizadas (versão prévia)
localization_priority: Normal
ms.openlocfilehash: ae04169044336f7e42d341c1e904090e55d568af
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871343"
---
# <a name="custom-functions-best-practices-preview"></a>Práticas recomendadas para funções personalizadas (versão prévia)

Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a>Solução de problemas

1. Quando testar o suplemento no Office para Windows, habilite o **[log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** para solucionar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução. O log de tempo de execução grava instruções `console.log` em um arquivo de log para ajudá-lo a descobrir problemas.

2. O suplemento não será carregado se uma ou mais funções personalizadas entrarem em conflito com as funções personalizadas de um suplemento registrado anteriormente. Nesse caso, você pode remover o suplemento existente ou se encontrar esse erro ao desenvolver um suplemento, você pode especificar um nome de namespace diferente em seu manifesto.

3. Para relatar problemas sobre este método de solução de problemas, envie comentários à equipe de funções personalizadas do Excel. Para fazer isso, selecione **Arquivo | Comentários | Enviar um Rosto Triste**. Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.

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

## <a name="associating-function-names-with-json-metadata"></a>Associar os nomes de função com metadados JSON

Conforme descrito no artigo [visão geral de funções personalizados](custom-functions-overview.md), um projeto de funções personalizados deve incluir um arquivo JSON de metadados e um arquivo de script (JavaScript ou TypeScript) para formar uma função completa. Para a função funcionar corretamente, será preciso associar o nome de função no arquivo de script à id listada no arquivo JSON. Esse processo é chamado de associação. Anote para incluir associações no final dos seus arquivos de código JavaScript; Caso contrário, as funções não funcionarão.

O exemplo a seguir mostra como fazer essa associação. A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.

* Use somente letras maiúsculas de uma função `name` e `id` no arquivo de metadados JSON. Não use uma combinação de casos ou somente letras minúsculas. Nesse caso, você pode acabar com dois valores que apenas variam por caso, o que causará a substituição não intencional de suas funções. Por exemplo, um objeto de função com uma `id` valor **adicionar** pode ser substituído pela declaração mais tarde no arquivo de objeto de função com uma `id` valor de **adicionar**. Além disso, a propriedade `name` define o nome da função que os usuários finais verão no Excel. O uso de letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente aos usuários finais do Excel, onde todos os nomes de funções internos são escritos em maiúsculas.

* No entanto, não é necessário colocar em maiúscula a função `name` quando associar. Por exemplo, `CustomFunctions.associate("add", add)` é equivalente a `CustomFunctions.associate("ADD", add)`.

* No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.

* No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo. Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`. 

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

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript. Observe que as propriedades `id` e `name` estão em letras maiúsculas no arquivo. 

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
    "id": "ADD",
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
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
