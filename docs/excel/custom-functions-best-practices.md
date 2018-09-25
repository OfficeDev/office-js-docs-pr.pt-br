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
# <a name="custom-functions-best-practices"></a><span data-ttu-id="28e10-103">Práticas recomendadas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="28e10-103">Custom functions best practices</span></span>

<span data-ttu-id="28e10-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="28e10-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="28e10-105">Manipulação de erro</span><span class="sxs-lookup"><span data-stu-id="28e10-105">Error handling</span></span>

<span data-ttu-id="28e10-106">Ao criar um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erro para considerar os erros no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="28e10-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="28e10-107">O tratamento de erros de funções personalizadas é o mesmo que [o tratamento de erros para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="28e10-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="28e10-108">No exemplo de código a seguir, `.catch` manipulará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="28e10-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="28e10-109">Depuração</span><span class="sxs-lookup"><span data-stu-id="28e10-109">Debugging</span></span>
<span data-ttu-id="28e10-110">Atualmente, o melhor método para depurar funções personalizadas do Excel é primeiro [fazer sideload](../testing/sideload-office-add-ins-for-testing.md) do seu suplemento no **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="28e10-110">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="28e10-111">Você pode depurar suas funções personalizadas usando [F12, a ferramenta de depuração nativa do seu navegador](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="28e10-111">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span> <span data-ttu-id="28e10-112">Use `console.log` instruções dentro do seu código de funções personalizadas para enviar a saída para o console em tempo real.</span><span class="sxs-lookup"><span data-stu-id="28e10-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

<span data-ttu-id="28e10-113">Se seu suplemento falhar ao registrar, [verifique se os certificados SSL estão configurados corretamente](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para o servidor Web que está hospedando o seu aplicativo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="28e10-113">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="28e10-114">Se você estiver testando seu suplemento de área de trabalho do Office 2016, é possível habilitar o [log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) para depurar problemas com o arquivo de manifesto XML do suplemento, bem como várias condições de instalação e tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="28e10-114">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span> 


## <a name="mapping-names"></a><span data-ttu-id="28e10-115">Nomes de mapeamento</span><span class="sxs-lookup"><span data-stu-id="28e10-115">Mapping names</span></span>

<span data-ttu-id="28e10-116">Por padrão, o nome de uma função personalizada no seu arquivo JavaScript geralmente é declarado usando letras maiúsculas e corresponde exatamente ao nome da função que os usuários finais veem no Excel.</span><span class="sxs-lookup"><span data-stu-id="28e10-116">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="28e10-117">No entanto, você pode alterar isso usando o `CustomFunctionsMappings` objeto para mapear um ou mais nomes das funções do arquivo JavaScript para diferentes valores que os usuários finais verão como nomes de função no Excel.</span><span class="sxs-lookup"><span data-stu-id="28e10-117">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="28e10-118">Isso é útil se você estiver usando um uglifier, webpack ou sintaxe de importação - todas eles têm dificuldade com nomes de função em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="28e10-118">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="28e10-119">`CustomFunctionsMappings` é opcional, possivelmente, para projetos que usam JavaScript, mas deve ser usado se o seu projeto usa TypeScript.</span><span class="sxs-lookup"><span data-stu-id="28e10-119">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="28e10-120">O exemplo de código a seguir define um único par chave-valor que mapeia o nome da função JavaScript `plusFortyTwo` para o `ADD42` nome da função na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="28e10-120">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="28e10-121">Quando o usuário final escolhe a função `ADD42` no Excel, a função `plusFortyTwo` JavaScript será executada.</span><span class="sxs-lookup"><span data-stu-id="28e10-121">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="28e10-122">O exemplo de código a seguir define dois pares chave-valor.</span><span class="sxs-lookup"><span data-stu-id="28e10-122">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="28e10-123">O primeiro par mapeia o nome da função JavaScript `plusFifty` para o `ADD50` nome da função na interface do usuário do Excel, e o segundo par mapeia o nome da função JavaScript `plusOneHundred` para o `ADD100` nome da função na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="28e10-123">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="28e10-124">Quando o usuário final escolhe a função `ADD50` no Excel, a função `plusFifty` JavaScript será executada.</span><span class="sxs-lookup"><span data-stu-id="28e10-124">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="28e10-125">Quando o usuário final escolhe a função `ADD100` no Excel, a função `plusOneHundred` JavaScript será executada.</span><span class="sxs-lookup"><span data-stu-id="28e10-125">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="28e10-126">Confira também</span><span class="sxs-lookup"><span data-stu-id="28e10-126">See also</span></span>

- [<span data-ttu-id="28e10-127">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="28e10-127">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="28e10-128">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="28e10-128">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="28e10-129">Tempo de execução para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="28e10-129">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
