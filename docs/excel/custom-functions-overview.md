---
ms.date: 05/17/2020
description: Criar uma função personalizada do Excel para seu suplemento do Office
title: Criar funções personalizadas no Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: dabb196bc4b55bd4852f9c857767dcabd3063045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44276005"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="bc164-103">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="bc164-103">Create custom functions in Excel</span></span>

<span data-ttu-id="bc164-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="bc164-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="bc164-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="bc164-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="bc164-106">A imagem animada a seguir mostra a sua pasta de trabalho solicitando uma função que você criou com o JavaScript ou o Typescript.</span><span class="sxs-lookup"><span data-stu-id="bc164-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="bc164-107">Neste exemplo, a função personalizada `=MYFUNCTION.SPHEREVOLUME` calcula o volume de uma esfera.</span><span class="sxs-lookup"><span data-stu-id="bc164-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="bc164-108">O código a seguir define a função personalizada `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="bc164-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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

> [!NOTE]
> <span data-ttu-id="bc164-109">A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bc164-109">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="bc164-110">Como uma função personalizada é definida em código</span><span class="sxs-lookup"><span data-stu-id="bc164-110">How a custom function is defined in code</span></span>

<span data-ttu-id="bc164-111">Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento de funções personalizadas do Excel, ele criará arquivos que controlam as funções e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="bc164-111">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="bc164-112">Vamos nos concentrar em arquivos que são importantes para funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="bc164-112">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="bc164-113">File</span><span class="sxs-lookup"><span data-stu-id="bc164-113">File</span></span> | <span data-ttu-id="bc164-114">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="bc164-114">File format</span></span> | <span data-ttu-id="bc164-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="bc164-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="bc164-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="bc164-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="bc164-117">ou</span><span class="sxs-lookup"><span data-stu-id="bc164-117">or</span></span><br/><span data-ttu-id="bc164-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="bc164-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="bc164-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="bc164-119">JavaScript</span></span><br/><span data-ttu-id="bc164-120">ou</span><span class="sxs-lookup"><span data-stu-id="bc164-120">or</span></span><br/><span data-ttu-id="bc164-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="bc164-121">TypeScript</span></span> | <span data-ttu-id="bc164-122">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bc164-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="bc164-123">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="bc164-123">**./src/functions/functions.html**</span></span> | <span data-ttu-id="bc164-124">HTML</span><span class="sxs-lookup"><span data-stu-id="bc164-124">HTML</span></span> | <span data-ttu-id="bc164-125">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bc164-125">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="bc164-126">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="bc164-126">**./manifest.xml**</span></span> | <span data-ttu-id="bc164-127">XML</span><span class="sxs-lookup"><span data-stu-id="bc164-127">XML</span></span> | <span data-ttu-id="bc164-128">Especifica o local de vários arquivos que sua função personalizada usa, como as funções personalizadas JavaScript, JSON e arquivos HTML.</span><span class="sxs-lookup"><span data-stu-id="bc164-128">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="bc164-129">Ele também lista os locais dos arquivos de painel de tarefas, os arquivos de comando e especifica o tempo de execução que suas funções personalizadas devem usar.</span><span class="sxs-lookup"><span data-stu-id="bc164-129">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="bc164-130">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="bc164-130">Script file</span></span>

<span data-ttu-id="bc164-131">O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contém o código que define funções e comentários que definem a função.</span><span class="sxs-lookup"><span data-stu-id="bc164-131">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="bc164-132">O código a seguir define a função personalizada `add`.</span><span class="sxs-lookup"><span data-stu-id="bc164-132">The following code defines the custom function `add`.</span></span> <span data-ttu-id="bc164-133">Os comentários do código são usados para gerar um arquivo de metadados JSON que descreve a função personalizada ao Excel.</span><span class="sxs-lookup"><span data-stu-id="bc164-133">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="bc164-134">O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="bc164-134">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="bc164-135">Em seguida, dois parâmetros são declarados `first` e `second` , em seguida, suas `description` Propriedades.</span><span class="sxs-lookup"><span data-stu-id="bc164-135">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="bc164-136">Por fim, uma `returns` descrição é fornecida.</span><span class="sxs-lookup"><span data-stu-id="bc164-136">Finally, a `returns` description is given.</span></span> <span data-ttu-id="bc164-137">Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="bc164-137">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

### <a name="manifest-file"></a><span data-ttu-id="bc164-138">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="bc164-138">Manifest file</span></span>

<span data-ttu-id="bc164-139">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto criado pelo gerador do Office Yo) faz várias coisas:</span><span class="sxs-lookup"><span data-stu-id="bc164-139">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="bc164-140">Define o namespace para suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bc164-140">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="bc164-141">Um namespace se precede às suas funções personalizadas para ajudar os clientes a identificar suas funções como parte do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="bc164-141">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="bc164-142">Usos `<ExtensionPoint>` e `<Resources>` elementos exclusivos de um manifesto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bc164-142">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="bc164-143">Esses elementos contêm informações sobre os locais dos arquivos JavaScript, JSON e HTML.</span><span class="sxs-lookup"><span data-stu-id="bc164-143">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="bc164-144">Especifica o tempo de execução a ser usado para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="bc164-144">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="bc164-145">Recomendamos sempre usar um tempo de execução compartilhado, a menos que você tenha uma necessidade específica de outro tempo de execução, pois um tempo de execução compartilhado permite o compartilhamento de dados entre funções e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="bc164-145">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span>

<span data-ttu-id="bc164-146">Se você estiver usando o gerador de Yo Office para criar arquivos, recomendamos ajustar seu manifesto para usar um tempo de execução compartilhado, pois esse não é o padrão para esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="bc164-146">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="bc164-147">Para alterar o manifesto, siga as instruções em [configurar seu suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](./configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="bc164-147">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](./configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="bc164-148">Para ver um manifesto de trabalho completo de um suplemento de exemplo, confira [o repositório do GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="bc164-148">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="bc164-149">Coautoria</span><span class="sxs-lookup"><span data-stu-id="bc164-149">Coauthoring</span></span>

<span data-ttu-id="bc164-150">O Excel na Web e o Windows conectado a uma assinatura do Office 365 permitem que você coautor no Excel.</span><span class="sxs-lookup"><span data-stu-id="bc164-150">Excel on the web and Windows connected to an Office 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="bc164-151">Se sua pasta de trabalho usa uma função personalizada, seu colega de coautoria é solicitado a carregar o suplemento da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="bc164-151">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="bc164-152">Depois que você carregar o suplemento, a função personalizada compartilhará os resultados por meio de coautoria.</span><span class="sxs-lookup"><span data-stu-id="bc164-152">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="bc164-153">Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="bc164-153">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="bc164-154">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="bc164-154">Known issues</span></span>

<span data-ttu-id="bc164-155">Veja os problemas conhecidos no nosso [GitHub de funções do Excel personalizado repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="bc164-155">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="bc164-156">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="bc164-156">Next steps</span></span>

<span data-ttu-id="bc164-157">Quer experimentar funções personalizadas?</span><span class="sxs-lookup"><span data-stu-id="bc164-157">Want to try out custom functions?</span></span> <span data-ttu-id="bc164-158">Confira o simples [início rápido das funções personalizadas](../quickstarts/excel-custom-functions-quickstart.md) ou o mais detalhado [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md), caso ainda não tenha.</span><span class="sxs-lookup"><span data-stu-id="bc164-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="bc164-159">Outra maneira fácil de experimentar as funções personalizadas é usar o [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), que é um suplemento que permite com que você experimente as funções personalizadas diretamente no Excel.</span><span class="sxs-lookup"><span data-stu-id="bc164-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="bc164-160">Você pode experimentar criar a sua própria função personalizada ou usar os exemplos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="bc164-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="bc164-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="bc164-161">See also</span></span> 
* [<span data-ttu-id="bc164-162">Requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc164-162">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="bc164-163">Diretrizes de nomenclatura</span><span class="sxs-lookup"><span data-stu-id="bc164-163">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="bc164-164">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="bc164-164">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
