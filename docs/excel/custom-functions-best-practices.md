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
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="3a0ff-103">Práticas recomendadas para funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="3a0ff-103">Custom functions best practices</span></span>

<span data-ttu-id="3a0ff-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="3a0ff-105">Manipulação de erro</span><span class="sxs-lookup"><span data-stu-id="3a0ff-105">Error handling</span></span>

<span data-ttu-id="3a0ff-106">Ao criar um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erro para considerar os erros no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="3a0ff-107">O tratamento de erros de funções personalizadas é o mesmo que [tratamento de erros para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="3a0ff-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="3a0ff-108">No exemplo de código a seguir, `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="3a0ff-109">Depuração</span><span class="sxs-lookup"><span data-stu-id="3a0ff-109">Debugging</span></span>

<span data-ttu-id="3a0ff-p102">Atualmente, o melhor método para depuração de funções personalizadas do Excel é primeiro fazer o [sideload](../testing/sideload-office-add-ins-for-testing.md) do seu suplemento no **Excel Online**. Dessa forma você pode depurar as funções personalizadas usando a [ferramenta de depuração F12 nativa do navegador](../testing/debug-add-ins-in-office-online.md) em combinação com as técnicas a seguir:</span><span class="sxs-lookup"><span data-stu-id="3a0ff-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md). Use  statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="3a0ff-112">Use instruções `console.log` no seu código de funções personalizadas para enviar a saída para o console em tempo real.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="3a0ff-113">Use instruções `debugger;` no seu código de funções personalizadas para especificar pontos de interrupção onde a execução será interrompida quando a janela F12 estiver aberta.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-113">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="3a0ff-114">Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução será interrompida na instrução `debugger;`, permitindo que você inspecione manualmente os valores dos parâmetros antes que a função retorne.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-114">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="3a0ff-115">A instrução `debugger;` não tem efeito no Excel Online quando a janela F12 não está aberta.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-115">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="3a0ff-116">Atualmente, a instrução `debugger;` não tem efeito no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-116">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="3a0ff-117">Se seu suplemento falhar ao registrar, [verifique se os certificados SSL estão configurados corretamente](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que está hospedando o seu aplicativo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="3a0ff-118">Se você estiver testando seu suplemento de área de trabalho do Office 2016, é possível habilitar o [log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) para depurar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="3a0ff-119">Mapeamento de nomes de função para metadados JSON</span><span class="sxs-lookup"><span data-stu-id="3a0ff-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="3a0ff-120">Conforme descrito no artigo de [Visão geral de funções personalizadas](custom-functions-overview.md), um projeto de funções personalizadas deve incluir um arquivo de metadados JSON que fornece as informações exigidas pelo Excel para registrar as funções personalizadas e torná-las disponíveis aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-120">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="3a0ff-121">Além disso, dentro do arquivo JavaScript que define as funções personalizadas, você deve fornecer informações para especificar qual objeto de função no arquivo de metadados JSON corresponde a cada função personalizada no arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-121">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="3a0ff-122">Por exemplo, o exemplo de código a seguir define a função personalizada `add` e, em seguida, especifica que a função `add` corresponde ao objeto no arquivo de metadados JSON onde o valor da propriedade `id` é **ADD**.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="3a0ff-123">Tenha em mente as seguintes práticas recomendadas ao criar funções personalizadas no seu arquivo JavaScript e especificar informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="3a0ff-124">No arquivo JavaScript, especifique os nomes de função em camelCase.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-124">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="3a0ff-125">Por exemplo, o nome da função `addTenToInput` está escrito em camelCase: a primeira palavra no nome começa com uma letra minúscula e cada palavra subsequente no nome começa com uma letra maiúscula.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-125">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="3a0ff-126">No arquivo de metadados JSON, especifique o valor de cada propriedade`name` em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-126">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="3a0ff-127">A propriedade `name` define o nome da função que os usuários finais verão no Excel.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-127">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="3a0ff-128">Usar letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente para o usuário do Excel, pois os nomes de todas as funções internas estão em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-128">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="3a0ff-129">No arquivo de metadados JSON, especifique o valor de cada propriedade`id` em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-129">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="3a0ff-130">Isso torna óbvio qual parte da instrução `CustomFunctionMappings` no seu código JavaScript corresponde à propriedade `id` no arquivo de metadados JSON (desde que o seu nome de função use camelCase, conforme recomendado anteriormente).</span><span class="sxs-lookup"><span data-stu-id="3a0ff-130">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="3a0ff-131">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` é exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-131">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="3a0ff-132">Ou seja, não deve haver dois objetos de função no arquivo de metadados com o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-132">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="3a0ff-133">Além disso, não especifique dois valores `id` no arquivo de metadados que se distinguam apenas por letras maiúsculas ou minúsculas.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-133">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="3a0ff-134">Por exemplo, não defina um objeto de função com um `id` valor de **add** e outro objeto de função com um `id` valor de **ADD**.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-134">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="3a0ff-135">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON depois que ela tiver sido mapeada para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-135">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="3a0ff-136">Você pode alterar o nome da função que os usuários finais veem no Excel, atualizando a propriedade `name` dentro do arquivo de metadados JSON, mas você nunca deve alterar o valor de uma propriedade `id` depois que ele for estabelecido.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-136">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="3a0ff-137">No arquivo JavaScript, especifique todos os mapeamentos da função personalizada no mesmo local.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-137">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="3a0ff-138">Por exemplo, o exemplo de código a seguir define duas funções personalizadas e, em seguida, especifica a informação de mapeamento de ambas as funções.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-138">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="3a0ff-139">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas neste exemplo de código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3a0ff-139">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="3a0ff-140">Confira também</span><span class="sxs-lookup"><span data-stu-id="3a0ff-140">See also</span></span>

* [<span data-ttu-id="3a0ff-141">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="3a0ff-141">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="3a0ff-142">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a0ff-142">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3a0ff-143">Runtime de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="3a0ff-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3a0ff-144">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="3a0ff-144">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
