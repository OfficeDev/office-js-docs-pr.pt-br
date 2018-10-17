---
ms.date: 10/03/2018
description: Saiba mais sobre melhores práticas e padrões recomendados para funções personalizadas do Excel.
title: Práticas recomendadas para funções personalizadas
ms.openlocfilehash: 218e62cd074ccf3f3708bba90c938f7ddef059cb
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579818"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="4c3f6-103">Práticas recomendadas para funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="4c3f6-103">Custom functions best practices</span></span>

<span data-ttu-id="4c3f6-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="4c3f6-105">Manipulação de erro</span><span class="sxs-lookup"><span data-stu-id="4c3f6-105">Error handling</span></span>

<span data-ttu-id="4c3f6-p101">Ao construir um suplemento que define as funções personalizadas, certifique-se de incluir lógica para manipulação de erro para lidar com erros em tempo de execução. Manipulação de erro para funções personalizadas é o mesmo que [manipulação de erros para a API JavaScript do Excel, de maneira geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` manipulará quaisquer erros que ocorram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p101">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="4c3f6-109">Depuração</span><span class="sxs-lookup"><span data-stu-id="4c3f6-109">Debugging</span></span>

<span data-ttu-id="4c3f6-p102">Atualmente, o melhor método para depuração de funções personalizadas do Excel é primeiro fazer o [sideload](../testing/sideload-office-add-ins-for-testing.md) do seu suplemento no **Excel Online**. Dessa forma você pode depurar as funções personalizadas usando a [ferramenta de depuração F12 nativa do navegador](../testing/debug-add-ins-in-office-online.md) em combinação com as técnicas a seguir:</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="4c3f6-112">Use instruções `console.log` no seu código de funções personalizadas para enviar a saída para o console em tempo real.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="4c3f6-p103">Use `debugger;` instruções dentro de seu código de funções personalizadas para especificar os pontos de interrupção onde a execução fará uma pausa quando a janela F12 estiver aberta. Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução fará uma pausa sobre a instrução `debugger;` , permitindo que você inspecione manualmente os valores de parâmetro antes que a função retorne. A instrução `debugger;` não tem efeito no Excel Online quando a janela F12 não estiver aberta. Atualmente, a instrução `debugger;` não tem efeito no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p103">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open. For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns. The `debugger;` statement has no effect in Excel Online when the F12 window is not open. Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="4c3f6-117">Se seu suplemento falhar ao registrar, [verifique se os certificados SSL estão configurados corretamente](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que está hospedando o seu aplicativo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="4c3f6-118">Se você estiver testando seu suplemento no Office na área de trabalho do Windows, é possível habilitar o [log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) para depurar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="4c3f6-119">Mapeamento de nomes de função para metadados JSON</span><span class="sxs-lookup"><span data-stu-id="4c3f6-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="4c3f6-p104">Conforme descrito no artigo [Visão geral de funções personalizadas](custom-functions-overview.md) , um projeto de funções personalizadas deve incluir um arquivo de metadados JSON que forneça as informações exigidas pelo Excel para registrar as funções personalizadas e torná-las disponíveis aos usuários finais. Além disso, dentro do arquivo JavaScript que define suas funções personalizadas, você deve fornecer informações para especificar qual objeto de função no arquivo de metadados JSON corresponde a cada função personalizada no arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="4c3f6-122">Por exemplo, o exemplo de código a seguir define a função personalizada `add` e, em seguida, especifica que a função `add` corresponde ao objeto no arquivo de metadados JSON onde o valor da propriedade `id` é **ADD**.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="4c3f6-123">Tenha em mente as seguintes práticas recomendadas ao criar funções personalizadas no seu arquivo JavaScript e especificar informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="4c3f6-p105">No arquivo JavaScript, especifique nomes de função em camelCase. Por exemplo, o nome da função `addTenToInput` está escrito em camelCase: a primeira palavra no nome começa com uma letra minúscula e cada palavra subsequente no nome começa com uma letra maiúscula.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p105">In the JavaScript file, specify function names in camelCase. For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="4c3f6-p106">No arquivo de metadados JSON, especifique o valor de cada propriedade `name` em letras maiusculas. A propriedade `name` define o nome da função que os usuários finais verão no Excel. Usar letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente para usuários finais no Excel, onde todos os nomes de função interna estão em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p106">In the JSON metadata file, specify the value of each `name` property in uppercase. The `name` property defines the function name that end users will see in Excel. Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="4c3f6-p107">No arquivo de metadados JSON, especifique o valor de cada propriedade `id` em letras maiúsculas. Isso torna óbvio qual parte da instrução `CustomFunctionMappings` no seu código JavaScript corresponde à propriedade `id` no arquivo de metadados JSON (desde que o seu nome da função use camelCase, conforme recomendado anteriormente).</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p107">In the JSON metadata file, specify the value of each `id` property in uppercase. Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="4c3f6-131">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-131">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="4c3f6-p108">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` é exclusivo dentro do escopo do arquivo. Ou seja, não deve haver dois objetos function no arquivo de metadados com o mesmo valor de `id` . Além disso, não especifique dois valores `id` no arquivo de metadados que diferem somente por maiúsculas e minúsculas. Por exemplo, não defina um objeto function com um valor de `id` igual a **add** e outro objeto function com um valor de `id` igual a **ADD**.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p108">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value. Additionally, do not specify two `id` values in the metadata file that only differ by case. For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="4c3f6-p109">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON depois que ela foi mapeada para um nome de função JavaScript correspondente. Você pode alterar o nome da função que os usuários finais veem no Excel, atualizando a propriedade `name` dentro do arquivo de metadados JSON, mas você nunca deve alterar o valor de uma propriedade `id` depois que ela foi estabelecida.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p109">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="4c3f6-p110">No arquivo JavaScript, especifique todos os mapeamentos de função personalizada no mesmo local. Por exemplo, o exemplo de código a seguir define duas funções personalizadas e especifica as informações de mapeamento para ambas as funções.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-p110">In the JavaScript file, specify all custom function mappings in the same location. For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="4c3f6-140">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas neste exemplo de código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-140">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="4c3f6-141">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="4c3f6-141">Additional considerations</span></span>

<span data-ttu-id="4c3f6-142">Para criar um suplemento que possa ser executado em múltiplas plataformas (um dos locatários chaves de Suplementos do Office), você não deve acessar o Document Object Model (DOM) em funções personalizadas ou usar bibliotecas como a jQuery que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-142">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="4c3f6-143">No Excel para Windows, onde as funções personalizadas usam o  [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="4c3f6-143">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="4c3f6-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="4c3f6-144">See also</span></span>

* [<span data-ttu-id="4c3f6-145">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="4c3f6-145">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="4c3f6-146">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4c3f6-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="4c3f6-147">Runtime de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="4c3f6-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="4c3f6-148">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="4c3f6-148">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
