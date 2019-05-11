---
ms.date: 05/08/2019
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel
localization_priority: Priority
ms.openlocfilehash: d939d91e2c3fad239436621ae2704309f4f0f868
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952128"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="2bbf6-103">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="2bbf6-103">Create custom functions in Excel</span></span> 

<span data-ttu-id="2bbf6-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="2bbf6-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="2bbf6-106">Este artigo descreve como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="2bbf6-107">A imagem animada a seguir mostra a sua pasta de trabalho solicitando uma função que você criou com o JavaScript ou o Typescript.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-107">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="2bbf6-108">Neste exemplo, a função personalizada `=MYFUNCTION.SPHEREVOLUME` calcula o volume de uma esfera.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-108">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolume.gif" />

<span data-ttu-id="2bbf6-109">O código a seguir define a função personalizada `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-109">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
CustomFunctions.associate("SPHEREVOLUME", sphereVolume)
```

> [!NOTE]
> <span data-ttu-id="2bbf6-110">A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="2bbf6-111">Como uma função personalizada é definida em código</span><span class="sxs-lookup"><span data-stu-id="2bbf6-111">How a custom function is defined in code</span></span>

<span data-ttu-id="2bbf6-112">Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas em um projeto do Excel, você encontrará que cria os arquivos que controlam as funções, o painel de tarefas e o suplemento geral.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="2bbf6-113">Vamos nos concentrar em arquivos que são importantes para funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="2bbf6-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="2bbf6-114">File</span><span class="sxs-lookup"><span data-stu-id="2bbf6-114">File</span></span> | <span data-ttu-id="2bbf6-115">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="2bbf6-115">File format</span></span> | <span data-ttu-id="2bbf6-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bbf6-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="2bbf6-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="2bbf6-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="2bbf6-118">ou</span><span class="sxs-lookup"><span data-stu-id="2bbf6-118">or</span></span><br/><span data-ttu-id="2bbf6-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="2bbf6-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="2bbf6-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="2bbf6-120">JavaScript</span></span><br/><span data-ttu-id="2bbf6-121">ou</span><span class="sxs-lookup"><span data-stu-id="2bbf6-121">or</span></span><br/><span data-ttu-id="2bbf6-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="2bbf6-122">TypeScript</span></span> | <span data-ttu-id="2bbf6-123">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="2bbf6-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="2bbf6-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="2bbf6-125">HTML</span><span class="sxs-lookup"><span data-stu-id="2bbf6-125">HTML</span></span> | <span data-ttu-id="2bbf6-126">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="2bbf6-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="2bbf6-127">**./manifest.xml**</span></span> | <span data-ttu-id="2bbf6-128">XML</span><span class="sxs-lookup"><span data-stu-id="2bbf6-128">XML</span></span> | <span data-ttu-id="2bbf6-129">Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos JavaScript e HTML listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="2bbf6-130">Também lista os locais de outros arquivos, que o suplemento pode fazer uso, como os arquivos do painel de tarefas e arquivos de comando.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="2bbf6-131">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="2bbf6-131">Script file</span></span>

<span data-ttu-id="2bbf6-132">O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contém o código que define funções personalizadas, comentários que definem a função e associa os nomes das funções personalizadas a objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="2bbf6-133">O código a seguir define a função personalizada `add`.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="2bbf6-134">Os comentários do código são usados para gerar um arquivo de metadados JSON que descreve a função personalizada ao Excel.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="2bbf6-135">O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="2bbf6-136">Além disso, observe que dois parâmetros foram declarados, `first` e `second`, que é seguido por suas `description` propriedades.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-136">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="2bbf6-137">Por fim, uma `returns` descrição é fornecida.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="2bbf6-138">Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="2bbf6-138">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="2bbf6-139">O seguinte código também solicita `CustomFunctions.associate("ADD", add)` para associar a função `add()` com o seu ID no arquivo de metadados JSON `ADD`.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-139">The following code also calls `CustomFunctions.associate("ADD", add)` to associate the function `add()` with its ID in the JSON metadata file `ADD`.</span></span> <span data-ttu-id="2bbf6-140">Para mais informações sobre a associação de funções, confira as [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="2bbf6-140">For more information about associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="2bbf6-141">Note que o arquivo **functions.html**, que governa o carregamento do tempo de execução das funções personalizadas, deve vincular-se à CDN atual para as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-141">Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions.</span></span> <span data-ttu-id="2bbf6-142">Projetos preparados com a versão atual do gerador Yo Office referenciam a CDN correta.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-142">Projects prepared with the current version of the Yo Office generator reference the correct CDN.</span></span> <span data-ttu-id="2bbf6-143">Se você estiver readaptando um projeto anterior de função personalizada de março de 2019 ou anteriormente, você precisará copiar no código abaixo para a página **functions.html**.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-143">If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a><span data-ttu-id="2bbf6-144">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="2bbf6-144">Manifest file</span></span>

<span data-ttu-id="2bbf6-145">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-145">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="2bbf6-146">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-146">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="2bbf6-147">Se estiver usando o gerador Yo Office, seus arquivos de funções personalizadas gerados conterão um arquivo de manifesto mais complexo, que você pode comparar neste [repositório do Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="2bbf6-147">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="2bbf6-148">As URLs especificadas no arquivo de manifesto para as funções personalizadas JavaScript e JSON e arquivos HTML devem estar publicamente acessíveis e ter o mesmo subdomínio.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-148">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="2bbf6-149">Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-149">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="2bbf6-150">O namespace da função vem antes do nome da função e são separados por um ponto.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-150">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="2bbf6-151">Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-151">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="2bbf6-152">O namespace deve ser usado como identificador para o as sua empresa ou suplemento.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-152">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="2bbf6-153">Um namespace pode conter apenas caracteres alfanuméricos e períodos.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-153">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="2bbf6-154">Coautoria</span><span class="sxs-lookup"><span data-stu-id="2bbf6-154">Coauthoring</span></span>

<span data-ttu-id="2bbf6-155">O Excel Online e o Excel no Windows com uma assinatura do Office 365 permitem editar documentos em coautoria, e esse recurso funciona com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-155">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="2bbf6-156">Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-156">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="2bbf6-157">Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-157">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="2bbf6-158">Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="2bbf6-158">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="2bbf6-159">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="2bbf6-159">Known issues</span></span>

<span data-ttu-id="2bbf6-160">Veja os problemas conhecidos no nosso [GitHub de funções do Excel personalizado repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="2bbf6-160">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="2bbf6-161">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="2bbf6-161">Next steps</span></span>

<span data-ttu-id="2bbf6-162">Quer experimentar funções personalizadas?</span><span class="sxs-lookup"><span data-stu-id="2bbf6-162">Want to try out custom functions?</span></span> <span data-ttu-id="2bbf6-163">Confira o simples [início rápido das funções personalizadas](../quickstarts/excel-custom-functions-quickstart.md) ou o mais detalhado [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md), caso ainda não tenha.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-163">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span> 

<span data-ttu-id="2bbf6-164">Outra maneira fácil de experimentar as funções personalizadas é usar o [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), que é um suplemento que permite com que você experimente as funções personalizadas diretamente no Excel.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-164">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="2bbf6-165">Você pode experimentar criar a sua própria função personalizada ou usar os exemplos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="2bbf6-165">You can try out creating your own custom function or play with the provided samples.</span></span>

<span data-ttu-id="2bbf6-166">Pronto para ler mais sobre os recursos de funções personalizadas?</span><span class="sxs-lookup"><span data-stu-id="2bbf6-166">Ready to read more about the capabilities custom functions?</span></span> <span data-ttu-id="2bbf6-167">Saiba mais sobre a visão geral da [arquitetura de funções personalizadas](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="2bbf6-167">Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2bbf6-168">Confira também</span><span class="sxs-lookup"><span data-stu-id="2bbf6-168">See also</span></span> 
* [<span data-ttu-id="2bbf6-169">Requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="2bbf6-169">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="2bbf6-170">Diretrizes de nomenclatura</span><span class="sxs-lookup"><span data-stu-id="2bbf6-170">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="2bbf6-171">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="2bbf6-171">Best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="2bbf6-172">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="2bbf6-172">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
