---
ms.date: 01/08/2020
description: Criar uma função personalizada no Excel para o Suplemento do Office.
title: Criar funções personalizadas no Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 97037f201a237cdc6dae551552a0a1609a58b34c
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839870"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="b815e-103">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="b815e-103">Create custom functions in Excel</span></span>

<span data-ttu-id="b815e-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="b815e-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="b815e-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="b815e-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="b815e-106">A imagem animada a seguir mostra a sua pasta de trabalho solicitando uma função que você criou com o JavaScript ou o Typescript.</span><span class="sxs-lookup"><span data-stu-id="b815e-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="b815e-107">Neste exemplo, a função personalizada `=MYFUNCTION.SPHEREVOLUME` calcula o volume de uma esfera.</span><span class="sxs-lookup"><span data-stu-id="b815e-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="b815e-108">O código a seguir define a função personalizada `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="b815e-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!TIP]
> <span data-ttu-id="b815e-109">Se seu suplemento de função personalizada usará um painel de tarefas ou um botão da faixa de opções, além de executar o código de função personalizada, você precisará configurar um tempo de execução de JavaScript compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b815e-109">If your custom function add-in will use a task pane or a ribbon button, in addition to running custom function code, you will need to set up a shared JavaScript runtime.</span></span> <span data-ttu-id="b815e-110">Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="b815e-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="b815e-111">Como uma função personalizada é definida em código</span><span class="sxs-lookup"><span data-stu-id="b815e-111">How a custom function is defined in code</span></span>

<span data-ttu-id="b815e-112">Se você usar o [Gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento funções personalizadas do Excel, ele criará os arquivos que controlam as funções e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b815e-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="b815e-113">Vamos nos concentrar em arquivos que são importantes para funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="b815e-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="b815e-114">File</span><span class="sxs-lookup"><span data-stu-id="b815e-114">File</span></span> | <span data-ttu-id="b815e-115">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="b815e-115">File format</span></span> | <span data-ttu-id="b815e-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="b815e-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="b815e-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="b815e-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="b815e-118">ou</span><span class="sxs-lookup"><span data-stu-id="b815e-118">or</span></span><br/><span data-ttu-id="b815e-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="b815e-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="b815e-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="b815e-120">JavaScript</span></span><br/><span data-ttu-id="b815e-121">ou</span><span class="sxs-lookup"><span data-stu-id="b815e-121">or</span></span><br/><span data-ttu-id="b815e-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="b815e-122">TypeScript</span></span> | <span data-ttu-id="b815e-123">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b815e-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="b815e-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="b815e-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="b815e-125">HTML</span><span class="sxs-lookup"><span data-stu-id="b815e-125">HTML</span></span> | <span data-ttu-id="b815e-126">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b815e-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="b815e-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="b815e-127">**./manifest.xml**</span></span> | <span data-ttu-id="b815e-128">XML</span><span class="sxs-lookup"><span data-stu-id="b815e-128">XML</span></span> | <span data-ttu-id="b815e-129">Especifica o local de vários arquivos que a sua função personalizada usa, como as funções personalizadas JavaScript, JSON e arquivos HTML.</span><span class="sxs-lookup"><span data-stu-id="b815e-129">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="b815e-130">Ele também lista os locais de arquivos do painel de tarefas, os arquivos de comando e especifica o tempo de execução que suas funções personalizadas devem usar.</span><span class="sxs-lookup"><span data-stu-id="b815e-130">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="b815e-131">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="b815e-131">Script file</span></span>

<span data-ttu-id="b815e-132">O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contém o código que define funções e comentários que definem a função.</span><span class="sxs-lookup"><span data-stu-id="b815e-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="b815e-133">O código a seguir define a função personalizada `add`.</span><span class="sxs-lookup"><span data-stu-id="b815e-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="b815e-134">Os comentários do código são usados para gerar um arquivo de metadados JSON que descreve a função personalizada ao Excel.</span><span class="sxs-lookup"><span data-stu-id="b815e-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="b815e-135">O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="b815e-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="b815e-136">Em seguida, dois parâmetros são declarados, `first` e `second`, seguidos por suas propriedades de `description`.</span><span class="sxs-lookup"><span data-stu-id="b815e-136">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="b815e-137">Por fim, uma `returns` descrição é fornecida.</span><span class="sxs-lookup"><span data-stu-id="b815e-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="b815e-138">Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Gerar automaticamente os metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="b815e-138">For more information about what comments are required for your custom function, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

### <a name="manifest-file"></a><span data-ttu-id="b815e-139">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="b815e-139">Manifest file</span></span>

<span data-ttu-id="b815e-140">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto que o gerador de Yo Office cria) faz várias coisas:</span><span class="sxs-lookup"><span data-stu-id="b815e-140">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="b815e-141">Define o espaço de nomes das suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b815e-141">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="b815e-142">Um namespace se direciona para suas funções personalizadas para ajudar os clientes a identificar suas funções como parte do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b815e-142">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="b815e-143">Usa os elementos `<ExtensionPoint>` e `<Resources>` que são exclusivos de um manifesto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b815e-143">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="b815e-144">Esses elementos contêm informações sobre os locais dos arquivos JavaScript, JSON e HTML.</span><span class="sxs-lookup"><span data-stu-id="b815e-144">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="b815e-145">Especifica o tempo de execução a ser usado para a sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="b815e-145">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="b815e-146">Recomendamos sempre usar um tempo de execução compartilhado, a menos que você tenha uma necessidade específica para outro tempo de execução, porque um tempo de execução compartilhado permite o compartilhamento de dados entre funções e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b815e-146">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span> <span data-ttu-id="b815e-147">Observe que usar um tempo de execução compartilhado significa que seu suplemento usará o Internet Explorer 11, não o Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="b815e-147">Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.</span></span>

<span data-ttu-id="b815e-148">Se você estiver usando o gerador do Yo Office para criar arquivos, recomendamos ajustar o manifesto para usar o tempo de execução compartilhado, uma vez que esse não é o padrão para esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="b815e-148">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="b815e-149">Para alterar o manifesto, siga as instruções no [Configurar seu suplemento do Excel para usar um de tempo de execução JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="b815e-149">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="b815e-150">Para ver um manifesto funcional completo de um suplemento de amostra, consulte [esse repositório do GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="b815e-150">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="b815e-151">Coautoria</span><span class="sxs-lookup"><span data-stu-id="b815e-151">Coauthoring</span></span>

<span data-ttu-id="b815e-152">O Excel na Web e o Windows conectado a uma assinatura do Microsoft 365 permitem que você se conecte ao Excel.</span><span class="sxs-lookup"><span data-stu-id="b815e-152">Excel on the web and Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="b815e-153">Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="b815e-153">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="b815e-154">Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.</span><span class="sxs-lookup"><span data-stu-id="b815e-154">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="b815e-155">Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="b815e-155">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="next-steps"></a><span data-ttu-id="b815e-156">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="b815e-156">Next steps</span></span>

<span data-ttu-id="b815e-157">Quer experimentar funções personalizadas?</span><span class="sxs-lookup"><span data-stu-id="b815e-157">Want to try out custom functions?</span></span> <span data-ttu-id="b815e-158">Confira o simples [início rápido das funções personalizadas](../quickstarts/excel-custom-functions-quickstart.md) ou o mais detalhado [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md), caso ainda não tenha.</span><span class="sxs-lookup"><span data-stu-id="b815e-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="b815e-159">Outra maneira fácil de experimentar as funções personalizadas é usar o [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), que é um suplemento que permite com que você experimente as funções personalizadas diretamente no Excel.</span><span class="sxs-lookup"><span data-stu-id="b815e-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="b815e-160">Você pode experimentar criar a sua própria função personalizada ou usar os exemplos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="b815e-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="b815e-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="b815e-161">See also</span></span> 
* [<span data-ttu-id="b815e-162">Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="b815e-162">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
* [<span data-ttu-id="b815e-163">Conjuntos de requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b815e-163">Custom functions requirement sets</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="b815e-164">Diretrizes de nomenclatura de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b815e-164">Custom functions naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="b815e-165">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="b815e-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="b815e-166">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="b815e-166">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
