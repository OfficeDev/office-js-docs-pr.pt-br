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
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="07553-103">Práticas recomendadas para funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="07553-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="07553-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.</span><span class="sxs-lookup"><span data-stu-id="07553-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="07553-105">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="07553-105">Error handling</span></span>

<span data-ttu-id="07553-106">Quando criar um suplemento que define funções personalizadas, não deixe de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="07553-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="07553-107">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="07553-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="07553-108">No seguinte exemplo de código, `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="07553-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="troubleshooting"></a><span data-ttu-id="07553-109">Solução de problemas</span><span class="sxs-lookup"><span data-stu-id="07553-109">Troubleshooting</span></span>

<span data-ttu-id="07553-110">Quando testar o suplemento no Office para Windows, habilite o **[log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** para solucionar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="07553-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="07553-111">O log de tempo de execução grava instruções `console.log` em um arquivo de log para ajudá-lo a descobrir problemas.</span><span class="sxs-lookup"><span data-stu-id="07553-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="07553-112">Para relatar problemas sobre este método de solução de problemas, envie comentários à equipe de funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="07553-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="07553-113">Para fazer isso, selecione **Arquivo | Comentários | Enviar um Rosto Triste**.</span><span class="sxs-lookup"><span data-stu-id="07553-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="07553-114">Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.</span><span class="sxs-lookup"><span data-stu-id="07553-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="debugging"></a><span data-ttu-id="07553-115">Depuração</span><span class="sxs-lookup"><span data-stu-id="07553-115">Debugging</span></span>

<span data-ttu-id="07553-116">Atualmente, o método ideal para depuração de funções personalizadas do Excel consiste primeiro em [sideload](../testing/sideload-office-add-ins-for-testing.md) o suplemento no **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="07553-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="07553-117">Em seguida, para depurar as funções personalizadas, use a [ferramenta de depuração nativa F12 no navegador](../testing/debug-add-ins-in-office-online.md), associado às seguintes técnicas:</span><span class="sxs-lookup"><span data-stu-id="07553-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="07553-118">Use as instruções `console.log` no código das funções personalizadas para enviar saída ao console em tempo real.</span><span class="sxs-lookup"><span data-stu-id="07553-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="07553-119">Use as instruções `debugger;` no código das funções personalizadas para especificar pontos de interrupção, onde a execução será pausada quando a janela F12 for aberta.</span><span class="sxs-lookup"><span data-stu-id="07553-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="07553-120">Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução será pausada na instrução `debugger;`, o que permite inspecionar manualmente os valores dos parâmetros antes que a função retorne.</span><span class="sxs-lookup"><span data-stu-id="07553-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="07553-121">A instrução `debugger;` não afeta o Excel Online quando a janela F12 não está aberta.</span><span class="sxs-lookup"><span data-stu-id="07553-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="07553-122">Atualmente, a instrução `debugger;` não afeta o Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="07553-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="07553-123">Se o suplemento não for devidamente registrado, [ verifique se os certificados SSL estão configurados corretamente ](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que hospeda o aplicativo do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07553-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="07553-124">Como mapear nomes de função para metadados JSON</span><span class="sxs-lookup"><span data-stu-id="07553-124">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="07553-125">Conforme descrito no artigo [Visão geral de funções personalizadas](custom-functions-overview.md), um projeto de funções personalizadas deve incluir um arquivo de metadados JSON com as informações necessárias que o Excel exige para registrar as funções personalizadas e disponibilizá-las aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="07553-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="07553-126">Além disso, no arquivo JavaScript que define as funções personalizadas, você deve fornecer informações para especificar qual objeto de função no arquivo de metadados JSON corresponde a cada função personalizada no arquivo JavaScript.</span><span class="sxs-lookup"><span data-stu-id="07553-126">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="07553-127">Por exemplo, o seguinte código de exemplo define a função personalizada `add` e, em seguida, especifica que a função `add` corresponde ao objeto no arquivo de metadados JSON, em que o valor da propriedade `id` seja **ADD**.</span><span class="sxs-lookup"><span data-stu-id="07553-127">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="07553-128">Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="07553-128">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="07553-129">No arquivo JavaScript, especifique os nomes das funções no camelCase.</span><span class="sxs-lookup"><span data-stu-id="07553-129">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="07553-130">Por exemplo, o nome da função `addTenToInput` é escrito no camelCase: a primeira palavra no nome começa com uma letra minúscula e cada palavra subsequente no nome começa com uma letra maiúscula.</span><span class="sxs-lookup"><span data-stu-id="07553-130">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="07553-131">No arquivo de metadados JSON, especifique o valor de cada propriedade `name` em maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="07553-131">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="07553-132">A propriedade `name` define o nome da função que os usuários finais verão no Excel.</span><span class="sxs-lookup"><span data-stu-id="07553-132">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="07553-133">O uso de letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente aos usuários finais do Excel, onde todos os nomes de funções internos são escritos em maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="07553-133">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="07553-134">No arquivo de metadados JSON, especifique o valor de cada propriedade `id` em maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="07553-134">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="07553-135">Dessa maneira, fica claro qual parte da instrução `CustomFunctionMappings` no código JavaScript corresponde à propriedade `id`, no arquivo de metadados JSON, desde que o nome da função use camelCase, conforme recomendado anteriormente.</span><span class="sxs-lookup"><span data-stu-id="07553-135">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="07553-136">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="07553-136">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="07553-137">No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="07553-137">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="07553-138">Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="07553-138">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="07553-139">Além disso, não especifique dois valores `id` no arquivo de metadados, que tenham como diferença apenas o uso de maiúsculas e minúsculas.</span><span class="sxs-lookup"><span data-stu-id="07553-139">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="07553-140">Por exemplo, não defina um objeto de função com um valor `id` de **add** e outro objeto de função com um valor `id` de **ADD**.</span><span class="sxs-lookup"><span data-stu-id="07553-140">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="07553-141">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="07553-141">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="07553-142">Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.</span><span class="sxs-lookup"><span data-stu-id="07553-142">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="07553-143">No arquivo JavaScript, especifique todos os mapeamentos de funções personalizadas no mesmo local.</span><span class="sxs-lookup"><span data-stu-id="07553-143">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="07553-144">Por exemplo, o exemplo de código a seguir define duas funções personalizadas e, em seguida, especifica as informações de mapeamento para ambas.</span><span class="sxs-lookup"><span data-stu-id="07553-144">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="07553-145">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="07553-145">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="07553-146">Como declarar parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="07553-146">Declaring optional parameters</span></span> 
<span data-ttu-id="07553-147">No Excel para Windows (versão 1812 ou posterior), é possível declarar parâmetros opcionais para suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="07553-147">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="07553-148">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="07553-148">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="07553-149">Por exemplo, uma função `FOO` com um parâmetro obrigatório chamado `parameter1` e parâmetro opcional chamado `parameter2` seria exibida como `=FOO(parameter1, [parameter2])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="07553-149">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="07553-150">Para tornar um parâmetro opcional, adicione `"optional": true` ao parâmetro no arquivo JSON de metadados que define a função.</span><span class="sxs-lookup"><span data-stu-id="07553-150">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="07553-151">O exemplo a seguir mostra o provável aspecto disso para a função `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="07553-151">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="07553-152">Observe que o parâmetro `[third]` opcional segue os dois parâmetros obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="07553-152">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="07553-153">Os parâmetros obrigatórios aparecerão primeiro na interface do usuário da fórmula do Excel.</span><span class="sxs-lookup"><span data-stu-id="07553-153">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="07553-154">Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontecerá quando os parâmetros opcionais forem indefinidos.</span><span class="sxs-lookup"><span data-stu-id="07553-154">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="07553-155">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="07553-155">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="07553-156">Se o parâmetro `zipCode` estiver indefinido, o valor padrão será definido como 98052.</span><span class="sxs-lookup"><span data-stu-id="07553-156">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="07553-157">Se o parâmetro `dayOfWeek` estiver indefinido, ele será definido como Quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="07553-157">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="07553-158">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="07553-158">Additional considerations</span></span>

<span data-ttu-id="07553-159">Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="07553-159">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="07553-160">No Excel para Windows, onde as funções personalizadas usam o [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="07553-160">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="07553-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="07553-161">See also</span></span>

* [<span data-ttu-id="07553-162">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="07553-162">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="07553-163">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="07553-163">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="07553-164">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="07553-164">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="07553-165">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="07553-165">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
