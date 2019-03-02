---
ms.date: 01/08/2019
description: Saiba mais sobre as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.
title: Práticas recomendadas para funções personalizadas (versão prévia)
localization_priority: Normal
ms.openlocfilehash: 24c73ec643df073ac97dc399343a7feb0b0b4168
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359258"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="d2f18-103">Práticas recomendadas para funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="d2f18-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="d2f18-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.</span><span class="sxs-lookup"><span data-stu-id="d2f18-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="d2f18-105">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="d2f18-105">Error handling</span></span>

<span data-ttu-id="d2f18-106">Quando criar um suplemento que define funções personalizadas, não deixe de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d2f18-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="d2f18-107">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="d2f18-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="d2f18-108">No seguinte exemplo de código, `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="d2f18-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="troubleshooting"></a><span data-ttu-id="d2f18-109">Solução de problemas</span><span class="sxs-lookup"><span data-stu-id="d2f18-109">Troubleshooting</span></span>

1. <span data-ttu-id="d2f18-110">Quando testar o suplemento no Office para Windows, habilite o **[log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** para solucionar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d2f18-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="d2f18-111">O log de tempo de execução grava instruções `console.log` em um arquivo de log para ajudá-lo a descobrir problemas.</span><span class="sxs-lookup"><span data-stu-id="d2f18-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

2. <span data-ttu-id="d2f18-112">O suplemento não será carregado se uma ou mais funções personalizadas entrarem em conflito com as funções personalizadas de um suplemento registrado anteriormente.</span><span class="sxs-lookup"><span data-stu-id="d2f18-112">Your add-in will not load if one or more custom functions conflicts with a previously registered add-in's custom functions.</span></span> <span data-ttu-id="d2f18-113">Nesse caso, você pode remover o suplemento existente ou se encontrar esse erro ao desenvolver um suplemento, você pode especificar um nome de namespace diferente em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="d2f18-113">In this case, you can either remove the existing add-in, or if you encounter this error while developing an add-in, you can specify a different namespace name in your manifest.</span></span>

3. <span data-ttu-id="d2f18-114">Para relatar problemas sobre este método de solução de problemas, envie comentários à equipe de funções personalizadas do Excel.</span><span class="sxs-lookup"><span data-stu-id="d2f18-114">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="d2f18-115">Para fazer isso, selecione **Arquivo | Comentários | Enviar um Rosto Triste**.</span><span class="sxs-lookup"><span data-stu-id="d2f18-115">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="d2f18-116">Enviando um Rosto Triste, você fornece os registros necessários para entendermos o problema que você está enfrentando.</span><span class="sxs-lookup"><span data-stu-id="d2f18-116">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>


## <a name="debugging"></a><span data-ttu-id="d2f18-117">Depuração</span><span class="sxs-lookup"><span data-stu-id="d2f18-117">Debugging</span></span>

<span data-ttu-id="d2f18-118">Atualmente, o método ideal para depuração de funções personalizadas do Excel consiste primeiro em [sideload](../testing/sideload-office-add-ins-for-testing.md) o suplemento no **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="d2f18-118">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="d2f18-119">Em seguida, para depurar as funções personalizadas, use a [ferramenta de depuração nativa F12 no navegador](../testing/debug-add-ins-in-office-online.md), associado às seguintes técnicas:</span><span class="sxs-lookup"><span data-stu-id="d2f18-119">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="d2f18-120">Use as instruções `console.log` no código das funções personalizadas para enviar saída ao console em tempo real.</span><span class="sxs-lookup"><span data-stu-id="d2f18-120">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="d2f18-121">Use as instruções `debugger;` no código das funções personalizadas para especificar pontos de interrupção, onde a execução será pausada quando a janela F12 for aberta.</span><span class="sxs-lookup"><span data-stu-id="d2f18-121">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="d2f18-122">Por exemplo, se a função a seguir for executada enquanto a janela F12 estiver aberta, a execução será pausada na instrução `debugger;`, o que permite inspecionar manualmente os valores dos parâmetros antes que a função retorne.</span><span class="sxs-lookup"><span data-stu-id="d2f18-122">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="d2f18-123">A instrução `debugger;` não afeta o Excel Online quando a janela F12 não está aberta.</span><span class="sxs-lookup"><span data-stu-id="d2f18-123">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="d2f18-124">Atualmente, a instrução `debugger;` não afeta o Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="d2f18-124">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="d2f18-125">Se o suplemento não for devidamente registrado, [ verifique se os certificados SSL estão configurados corretamente ](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que hospeda o aplicativo do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d2f18-125">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="d2f18-126">Associar os nomes de função com metadados JSON</span><span class="sxs-lookup"><span data-stu-id="d2f18-126">Associating function names with JSON metadata</span></span>

<span data-ttu-id="d2f18-127">Conforme descrito no artigo [visão geral de funções personalizados](custom-functions-overview.md), um projeto de funções personalizados deve incluir um arquivo JSON de metadados e um arquivo de script (JavaScript ou TypeScript) para formar uma função completa.</span><span class="sxs-lookup"><span data-stu-id="d2f18-127">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="d2f18-128">Para a função funcionar corretamente, será preciso associar o nome de função no arquivo de script à id listada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="d2f18-128">For a function to work properly, you'll need to bind the name of the function in the script file to the id listed in the JSON file.</span></span> <span data-ttu-id="d2f18-129">Esse processo é chamado de associação.</span><span class="sxs-lookup"><span data-stu-id="d2f18-129">This process is called association.</span></span> <span data-ttu-id="d2f18-130">Anote para incluir associações no final dos seus arquivos de código JavaScript; Caso contrário, as funções não funcionarão.</span><span class="sxs-lookup"><span data-stu-id="d2f18-130">Make a note to include associations at the end of your JavaScript code files; otherwise, your functions will not work.</span></span>

<span data-ttu-id="d2f18-131">O exemplo a seguir mostra como fazer essa associação.</span><span class="sxs-lookup"><span data-stu-id="d2f18-131">The following code sample shows how to do this association.</span></span> <span data-ttu-id="d2f18-132">A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.</span><span class="sxs-lookup"><span data-stu-id="d2f18-132">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add); 
```

<span data-ttu-id="d2f18-133">Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="d2f18-133">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="d2f18-134">Use somente letras maiúsculas de uma função `name` e `id` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="d2f18-134">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="d2f18-135">Não use uma combinação de casos ou somente letras minúsculas.</span><span class="sxs-lookup"><span data-stu-id="d2f18-135">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="d2f18-136">Nesse caso, você pode acabar com dois valores que apenas variam por caso, o que causará a substituição não intencional de suas funções.</span><span class="sxs-lookup"><span data-stu-id="d2f18-136">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="d2f18-137">Por exemplo, um objeto de função com uma `id` valor **adicionar** pode ser substituído pela declaração mais tarde no arquivo de objeto de função com uma `id` valor de **adicionar**.</span><span class="sxs-lookup"><span data-stu-id="d2f18-137">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="d2f18-138">Além disso, a propriedade `name` define o nome da função que os usuários finais verão no Excel.</span><span class="sxs-lookup"><span data-stu-id="d2f18-138">Additionally, the `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="d2f18-139">O uso de letras maiúsculas para o nome de cada função personalizada fornece uma experiência consistente aos usuários finais do Excel, onde todos os nomes de funções internos são escritos em maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="d2f18-139">Using uppercase letters for the name of each custom function provides a consistent experience in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="d2f18-140">No entanto, não é necessário colocar em maiúscula a função `name` quando associar.</span><span class="sxs-lookup"><span data-stu-id="d2f18-140">However, it is not necessary to capitalize the function's `name` when associating.</span></span> <span data-ttu-id="d2f18-141">Por exemplo, `CustomFunctions.associate("add", add)` é equivalente a `CustomFunctions.associate("ADD", add)`.</span><span class="sxs-lookup"><span data-stu-id="d2f18-141">For example, `CustomFunctions.associate("add", add)` is equivalent to `CustomFunctions.associate("ADD", add)`.</span></span>

* <span data-ttu-id="d2f18-142">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="d2f18-142">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="d2f18-143">No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="d2f18-143">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="d2f18-144">Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="d2f18-144">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="d2f18-145">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="d2f18-145">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="d2f18-146">Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.</span><span class="sxs-lookup"><span data-stu-id="d2f18-146">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="d2f18-147">No arquivo JavaScript, especifique todos os mapeamentos de funções personalizadas no mesmo local.</span><span class="sxs-lookup"><span data-stu-id="d2f18-147">In the JavaScript file, specify all custom function associations in the same location.</span></span> <span data-ttu-id="d2f18-148">Por exemplo, o exemplo de código a seguir define duas funções personalizadas e, em seguida, especifica as informações de mapeamento para ambas.</span><span class="sxs-lookup"><span data-stu-id="d2f18-148">For example, the following code sample defines two custom functions and then specifies the association information for both functions.</span></span>

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

    <span data-ttu-id="d2f18-149">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d2f18-149">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="d2f18-150">Observe que as propriedades `id` e `name` estão em letras maiúsculas no arquivo.</span><span class="sxs-lookup"><span data-stu-id="d2f18-150">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="d2f18-151">Como declarar parâmetros opcionais</span><span class="sxs-lookup"><span data-stu-id="d2f18-151">Declaring optional parameters</span></span> 
<span data-ttu-id="d2f18-152">No Excel para Windows (versão 1812 ou posterior), é possível declarar parâmetros opcionais para suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d2f18-152">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="d2f18-153">Quando um usuário invoca uma função no Excel, os parâmetros opcionais são exibidos entre colchetes.</span><span class="sxs-lookup"><span data-stu-id="d2f18-153">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="d2f18-154">Por exemplo, uma função `FOO` com um parâmetro obrigatório chamado `parameter1` e parâmetro opcional chamado `parameter2` seria exibida como `=FOO(parameter1, [parameter2])` no Excel.</span><span class="sxs-lookup"><span data-stu-id="d2f18-154">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="d2f18-155">Para tornar um parâmetro opcional, adicione `"optional": true` ao parâmetro no arquivo JSON de metadados que define a função.</span><span class="sxs-lookup"><span data-stu-id="d2f18-155">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="d2f18-156">O exemplo a seguir mostra o provável aspecto disso para a função `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="d2f18-156">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="d2f18-157">Observe que o parâmetro `[third]` opcional segue os dois parâmetros obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="d2f18-157">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="d2f18-158">Os parâmetros obrigatórios aparecerão primeiro na interface do usuário da fórmula do Excel.</span><span class="sxs-lookup"><span data-stu-id="d2f18-158">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="d2f18-159">Ao definir uma função que contenha um ou mais parâmetros opcionais, especifique o que acontecerá quando os parâmetros opcionais forem indefinidos.</span><span class="sxs-lookup"><span data-stu-id="d2f18-159">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="d2f18-160">No exemplo a seguir, `zipCode` e `dayOfWeek` são dois parâmetros opcionais da função `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="d2f18-160">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="d2f18-161">Se o parâmetro `zipCode` estiver indefinido, o valor padrão será definido como 98052.</span><span class="sxs-lookup"><span data-stu-id="d2f18-161">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="d2f18-162">Se o parâmetro `dayOfWeek` estiver indefinido, ele será definido como Quarta-feira.</span><span class="sxs-lookup"><span data-stu-id="d2f18-162">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="d2f18-163">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="d2f18-163">Additional considerations</span></span>

<span data-ttu-id="d2f18-164">Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="d2f18-164">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="d2f18-165">No Excel para Windows, onde as funções personalizadas usam o [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="d2f18-165">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="d2f18-166">Confira também</span><span class="sxs-lookup"><span data-stu-id="d2f18-166">See also</span></span>

* [<span data-ttu-id="d2f18-167">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="d2f18-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="d2f18-168">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d2f18-168">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d2f18-169">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="d2f18-169">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="d2f18-170">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d2f18-170">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="d2f18-171">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="d2f18-171">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
