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
# <a name="custom-functions-best-practices"></a><span data-ttu-id="bf408-103">Práticas recomendadas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bf408-103">Custom functions best practices</span></span>

<span data-ttu-id="bf408-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="bf408-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="bf408-105">Manipulação de erros</span><span class="sxs-lookup"><span data-stu-id="bf408-105">Error handling</span></span>

<span data-ttu-id="bf408-106">Quando você cria um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erros para considerar os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="bf408-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="bf408-107">Em geral, a manipulação de erros para funções personalizadas é a mesma que [a manipulação de erros para a API JavaScript do Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="bf408-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="bf408-108">No exemplo de código a seguir, `.catch` manipulará os erros que ocorram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="bf408-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="error-logging"></a><span data-ttu-id="bf408-109">Log de erros</span><span class="sxs-lookup"><span data-stu-id="bf408-109">Error logging</span></span>

<span data-ttu-id="bf408-110">Você pode ativar o log de erros para o suplemento de funções personalizadas de várias maneiras, como:</span><span class="sxs-lookup"><span data-stu-id="bf408-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="bf408-111">[Use o log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) para depurar o arquivo de manifesto XML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf408-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="bf408-112">Use `console.log` instruções dentro do seu código de funções personalizadas para enviar a saída para o console em tempo real.</span><span class="sxs-lookup"><span data-stu-id="bf408-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="bf408-113">O log de tempo de execução está atualmente disponível apenas para a área de trabalho do Office 2016.</span><span class="sxs-lookup"><span data-stu-id="bf408-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="bf408-114">Depuração</span><span class="sxs-lookup"><span data-stu-id="bf408-114">Debugging</span></span>

<span data-ttu-id="bf408-115">Atualmente, o melhor método para depurar as funções personalizadas do Excel é usar o [Excel Online](https://www.office.com/launch/excel) e usar a ferramenta de depuração F12 nativa em seu navegador.</span><span class="sxs-lookup"><span data-stu-id="bf408-115">Currently, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser.</span></span> <span data-ttu-id="bf408-116">Outras ferramentas de depuração para funções personalizadas podem estar disponíveis no futuro.</span><span class="sxs-lookup"><span data-stu-id="bf408-116">Additional debugging tools for custom functions may be available in the future.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="bf408-117">Nomes de mapeamento</span><span class="sxs-lookup"><span data-stu-id="bf408-117">Mapping names</span></span>

<span data-ttu-id="bf408-118">Por padrão, o nome de uma função personalizada no seu arquivo JavaScript geralmente é declarado usando letras maiúsculas e corresponde exatamente ao nome da função que os usuários finais veem no Excel.</span><span class="sxs-lookup"><span data-stu-id="bf408-118">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="bf408-119">No entanto, você pode alterar isso usando o `CustomFunctionsMappings` objeto para mapear um ou mais nomes das funções do arquivo JavaScript para diferentes valores que os usuários finais verão como nomes de função no Excel.</span><span class="sxs-lookup"><span data-stu-id="bf408-119">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="bf408-120">Embora você não precisa usar `CustomFunctionsMapping`, pode ser útil se estiver usando uma sintaxe uglifier, webpack ou de importação - todas as quais têm dificuldade com nomes de função em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="bf408-120">Although you're not required to use `CustomFunctionsMapping`, it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span>
  
<span data-ttu-id="bf408-121">O exemplo de código a seguir define um único par chave-valor que mapeia o nome da função JavaScript `plusFortyTwo` para o `ADD42` nome da função na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="bf408-121">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="bf408-122">Quando o usuário final escolhe a função `ADD42` no Excel, a função `plusFortyTwo` JavaScript será executada.</span><span class="sxs-lookup"><span data-stu-id="bf408-122">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="bf408-123">O exemplo de código a seguir define dois pares chave-valor.</span><span class="sxs-lookup"><span data-stu-id="bf408-123">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="bf408-124">O primeiro par mapeia o nome da função JavaScript `plusFifty` para o `ADD50` nome da função na interface do usuário do Excel, e o segundo par mapeia o nome da função JavaScript `plusOneHundred` para o `ADD100` nome da função na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="bf408-124">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="bf408-125">Quando o usuário final escolhe a função `ADD50` no Excel, a função `plusFifty` JavaScript será executada.</span><span class="sxs-lookup"><span data-stu-id="bf408-125">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="bf408-126">Quando o usuário final escolhe a função `ADD100` no Excel, a função `plusOneHundred` JavaScript será executada.</span><span class="sxs-lookup"><span data-stu-id="bf408-126">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="bf408-127">Confira também</span><span class="sxs-lookup"><span data-stu-id="bf408-127">See also</span></span>

* [<span data-ttu-id="bf408-128">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="bf408-128">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="bf408-129">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bf408-129">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="bf408-130">Tempo de execução para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="bf408-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)