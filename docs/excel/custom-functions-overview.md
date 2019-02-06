---
ms.date: 01/30/2019
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (versão prévia)
localization_priority: Priority
ms.openlocfilehash: 7df07d7abda182c5b9413047f74a5cc31f7d4057
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742356"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="f144e-103">Criar funções personalizadas no Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="f144e-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="f144e-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="f144e-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="f144e-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="f144e-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="f144e-106">Este artigo descreve como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f144e-107">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="f144e-108">A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par dos números que o usuário especifica como parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="f144e-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="f144e-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="f144e-110">A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f144e-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="f144e-111">Componentes de um projeto de suplemento de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f144e-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="f144e-112">Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas do Excel em um projeto, você verá os seguintes arquivos no projeto que o gerador cria:</span><span class="sxs-lookup"><span data-stu-id="f144e-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="f144e-113">File</span><span class="sxs-lookup"><span data-stu-id="f144e-113">File</span></span> | <span data-ttu-id="f144e-114">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="f144e-114">File format</span></span> | <span data-ttu-id="f144e-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="f144e-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="f144e-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="f144e-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="f144e-117">ou</span><span class="sxs-lookup"><span data-stu-id="f144e-117">or</span></span><br/><span data-ttu-id="f144e-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="f144e-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="f144e-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f144e-119">JavaScript</span></span><br/><span data-ttu-id="f144e-120">ou</span><span class="sxs-lookup"><span data-stu-id="f144e-120">or</span></span><br/><span data-ttu-id="f144e-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f144e-121">TypeScript</span></span> | <span data-ttu-id="f144e-122">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f144e-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="f144e-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="f144e-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="f144e-124">JSON</span><span class="sxs-lookup"><span data-stu-id="f144e-124">JSON</span></span> | <span data-ttu-id="f144e-125">Contém metadados que descrevem funções personalizadas e permitem que o Excel registre funções personalizadas e as disponibilize para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="f144e-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="f144e-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="f144e-126">**./index.html**</span></span> | <span data-ttu-id="f144e-127">HTML</span><span class="sxs-lookup"><span data-stu-id="f144e-127">HTML</span></span> | <span data-ttu-id="f144e-128">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f144e-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="f144e-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="f144e-129">**./manifest.xml**</span></span> | <span data-ttu-id="f144e-130">XML</span><span class="sxs-lookup"><span data-stu-id="f144e-130">XML</span></span> | <span data-ttu-id="f144e-131">Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="f144e-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="f144e-132">As seções a seguir fornecem mais informações sobre esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="f144e-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="f144e-133">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="f144e-133">Script file</span></span>

<span data-ttu-id="f144e-134">O arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** no projeto gerador que o Yo Office cria) contém o código que define funções personalizadas e mapeia os nomes da funções personalizadas aos objetos em [arquivos de metadados JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="f144e-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="f144e-135">Por exemplo, o código a seguir define funções personalizadas `add` e `increment` e especifica as informações de mapeamento para as duas funções.</span><span class="sxs-lookup"><span data-stu-id="f144e-135">For example, the following code defines the custom functions `add` and `increment` and then specifies association information for both functions.</span></span> <span data-ttu-id="f144e-136">A `add` função está associada com o objeto no arquivo nos metadados JSON onde o valor da propriedade `id` é **Adicionar**e a `increment` função é associada com o objeto no arquivo metadados onde o valor da propriedade`id`é **INCREMENTO**.</span><span class="sxs-lookup"><span data-stu-id="f144e-136">The `add` function is associated with the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is associated with the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="f144e-137">Ver [Práticas recomendadas de funções personalizados](custom-functions-best-practices.md#associating-function-names-with-json-metadata) para saber mais como associar os nomes de função no arquivo de script para objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-137">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about associating function names in the script file to objects in the JSON metadata file.</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a><span data-ttu-id="f144e-138">Arquivo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="f144e-138">JSON metadata file</span></span>

<span data-ttu-id="f144e-139">O arquivo de metadados de funções personalizadas (**./config/customfunctions.json** no projeto gerador que o Yo Office cria) fornece informações exigidas pelo Excel para registrar funções personalizadas e disponibilizá-las aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="f144e-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="f144e-140">Funções personalizadas são registradas quando um usuário usar um suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="f144e-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="f144e-141">Depois disso, eles estão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento foi inicialmente executado.)</span><span class="sxs-lookup"><span data-stu-id="f144e-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="f144e-142">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="f144e-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="f144e-143">O seguinte código em **customfunctions.json** especifica os metadados para a função `add` e a função `increment` descritas anteriormente.</span><span class="sxs-lookup"><span data-stu-id="f144e-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="f144e-144">A tabela que segue o código fornece informações detalhadas sobre as propriedades individuais nesse objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="f144e-145">Ver [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#associating-function-names-with-json-metadata) para saber mais sobre como especificar o valor das propriedades`id` e `name` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-145">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
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
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

<span data-ttu-id="f144e-146">A tabela a seguir lista as propriedades normalmente presentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="f144e-147">Para saber mais sobre o arquivo de metadados JSON, confira [Metadados de funções personalizadas](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="f144e-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="f144e-148">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f144e-148">Property</span></span>  | <span data-ttu-id="f144e-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="f144e-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="f144e-150">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-150">A unique ID for the function.</span></span> <span data-ttu-id="f144e-151">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="f144e-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="f144e-152">Nome da função que o usuário final vê no Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="f144e-153">No Excel, o nome de função será prefixado pelo namespace de funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="f144e-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="f144e-154">A URL da página é exibida quando um usuário solicitar ajuda.</span><span class="sxs-lookup"><span data-stu-id="f144e-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="f144e-155">Descreve o que faz a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-155">Describes what the function does.</span></span> <span data-ttu-id="f144e-156">Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu de preenchimento automático do Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="f144e-157">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="f144e-158">Para obter informações detalhadas sobre esse objeto, consulte [resultado](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="f144e-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="f144e-159">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="f144e-160">Para obter informações detalhadas sobre esse objeto, consulte [parâmetros](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="f144e-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="f144e-161">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="f144e-162">Confira mais informações sobre como essa propriedade pode ser usada em [funções Streaming](#streaming-functions) e [Cancelar uma função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="f144e-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function).</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="f144e-163">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="f144e-163">Manifest file</span></span>

<span data-ttu-id="f144e-164">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="f144e-165">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f144e-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="f144e-166">Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f144e-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="f144e-167">O namespace da função vem antes do nome da função e são separados por um ponto.</span><span class="sxs-lookup"><span data-stu-id="f144e-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="f144e-168">Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="f144e-169">O namespace deve ser usado como identificador para o as sua empresa ou suplemento.</span><span class="sxs-lookup"><span data-stu-id="f144e-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="f144e-170">Um namespace pode conter apenas caracteres alfanuméricos e períodos.</span><span class="sxs-lookup"><span data-stu-id="f144e-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="f144e-171">Funções que retornam os dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="f144e-171">Functions that return data from external sources</span></span>

<span data-ttu-id="f144e-172">Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="f144e-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="f144e-173">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="f144e-174">Resolva a promessa com o uso da função retorno de chamada de valor final.</span><span class="sxs-lookup"><span data-stu-id="f144e-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="f144e-175">Exibição de funções personalizados mostra um `#GETTING_DATA` resultado temporário na célula enquanto o Excel espera do resultado final.</span><span class="sxs-lookup"><span data-stu-id="f144e-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="f144e-176">Os usuários podem interagir normalmente com o restante da planilha aguardando o resultado.</span><span class="sxs-lookup"><span data-stu-id="f144e-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="f144e-177">No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro.</span><span class="sxs-lookup"><span data-stu-id="f144e-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="f144e-178">Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa [XHR](custom-functions-runtime.md#xhr-example) para chamar um serviço web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="f144e-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="f144e-179">Funções Streaming</span><span class="sxs-lookup"><span data-stu-id="f144e-179">Streaming functions</span></span>

<span data-ttu-id="f144e-180">Funções personalizadas de streaming permitem a saída de dados das células repetidamente ao longo do tempo, sem a necessidade de um usuário explicitamente solicitar a atualização de dados.</span><span class="sxs-lookup"><span data-stu-id="f144e-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="f144e-181">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="f144e-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="f144e-182">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="f144e-182">Note the following about this code:</span></span>

- <span data-ttu-id="f144e-183">Cada novo valor usando o Excel automaticamente exibirá o `setResult` retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f144e-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="f144e-184">O segundo parâmetro de entrada, `handler`, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="f144e-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="f144e-185">O `onCanceled` retorno de chamada define a função que é executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="f144e-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="f144e-186">Implemente um identificador de cancelamento assim para qualquer função de streaming.</span><span class="sxs-lookup"><span data-stu-id="f144e-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="f144e-187">Para saber mais, confira [Cancelar uma função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="f144e-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

<span data-ttu-id="f144e-188">Quando você especifica os metadados para uma função streaming no arquivo de metadados JSON, você deve definir as propriedades `"cancelable": true` e `"stream": true` no `options` objeto, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f144e-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="f144e-189">Cancelar uma função</span><span class="sxs-lookup"><span data-stu-id="f144e-189">Canceling a function</span></span>

<span data-ttu-id="f144e-190">Em algumas situações, talvez seja necessário cancelar a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU.</span><span class="sxs-lookup"><span data-stu-id="f144e-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="f144e-191">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="f144e-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="f144e-192">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="f144e-192">When the user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="f144e-193">Quando é alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-193">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="f144e-194">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="f144e-194">In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="f144e-195">Quando o usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="f144e-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="f144e-196">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="f144e-196">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="f144e-197">Para habilitar o recurso cancelar uma função, implemente um identificador de cancelamento dentro da função JavaScript e especifique a propriedade `"cancelable": true` dentro do `options` objeto nos metadados JSON que descreve a função.</span><span class="sxs-lookup"><span data-stu-id="f144e-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="f144e-198">Amostras de código na seção anterior neste artigo fornecem um exemplo dessas técnicas.</span><span class="sxs-lookup"><span data-stu-id="f144e-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="f144e-199">Salvar e compartilhar estado</span><span class="sxs-lookup"><span data-stu-id="f144e-199">Saving and sharing state</span></span>

<span data-ttu-id="f144e-200">Funções personalizadas podem salvar os dados em variáveis, que podem ser usadas em chamadas subsequentes.</span><span class="sxs-lookup"><span data-stu-id="f144e-200">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="f144e-201">O estado salvo é útil quando os usuários solicitam a mesma função personalizada usando mais de uma célula, porque todas as ocorrências da função podem acessar o estado.</span><span class="sxs-lookup"><span data-stu-id="f144e-201">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="f144e-202">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="f144e-202">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="f144e-203">O código a seguir mostra uma implementação da função de streaming de temperatura que salva o estado globalmente.</span><span class="sxs-lookup"><span data-stu-id="f144e-203">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="f144e-204">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="f144e-204">Note the following about this code:</span></span>

- <span data-ttu-id="f144e-205">A função `streamTemperature` atualiza o valor de temperatura exibido na célula a cada segundo e ele usa a variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="f144e-205">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="f144e-206">Como `streamTemperature` é uma função de streaming, ela implementa um identificador de cancelamento que será executado quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="f144e-206">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="f144e-207">Se um usuário ligar a função`streamTemperature` de várias células no Excel, a função `streamTemperature` lê os dados a partir da mesma`savedTemperatures` variável toda vez que ela for executada.</span><span class="sxs-lookup"><span data-stu-id="f144e-207">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="f144e-208">`refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo e armazena o resultado na variável`savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="f144e-208">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="f144e-209">Como a função`refreshTemperature` não é exibida para os usuários finais no Excel, não é necessário ser registrado no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-209">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="f144e-210">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="f144e-210">Working with ranges of data</span></span>

<span data-ttu-id="f144e-211">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="f144e-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="f144e-212">Em JavaScript, um intervalo de dados é representado como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="f144e-212">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="f144e-213">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-213">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f144e-214">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="f144e-214">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="f144e-215">Observe que, nos metadados JSON dessa função, você deve definir o parâmetro `type` propriedade para `matrix`.</span><span class="sxs-lookup"><span data-stu-id="f144e-215">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="f144e-216">Determinar quais células chamadas de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="f144e-216">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="f144e-217">Em alguns casos, você precisará obter o endereço da célula invocada na sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f144e-217">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f144e-218">Isso pode ser útil para os seguintes tipos de cenários:</span><span class="sxs-lookup"><span data-stu-id="f144e-218">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="f144e-219">Formatação de intervalos: Use o endereço da célula como a chave para armazenar informações em [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="f144e-219">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="f144e-220">Em seguida, use [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="f144e-220">Then, use [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="f144e-221">Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `AsyncStorage` usando `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="f144e-221">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="f144e-222">Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.</span><span class="sxs-lookup"><span data-stu-id="f144e-222">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="f144e-223">As informações sobre o endereço de uma célula serão expostas somente se `requiresAddress` estiver marcado como `true` no arquivo de metadados JSON da função.</span><span class="sxs-lookup"><span data-stu-id="f144e-223">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="f144e-224">A seguir, um exemplo disso:</span><span class="sxs-lookup"><span data-stu-id="f144e-224">The following sample gives an example of this:</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="f144e-225">No arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts**), também será necessário adicionar uma função `getAddress` para encontrar o endereço de uma célula.</span><span class="sxs-lookup"><span data-stu-id="f144e-225">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="f144e-226">Essa função pode ter parâmetros, conforme mostrado no exemplo a seguir como `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="f144e-226">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="f144e-227">O último parâmetro sempre será `invocationContext`, um objeto com o local da célula que o Excel passa quando `requiresAddress` é marcado como `true` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f144e-227">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="f144e-228">Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="f144e-228">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="f144e-229">Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="f144e-229">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="handling-errors"></a><span data-ttu-id="f144e-230">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="f144e-230">Handling errors</span></span>

<span data-ttu-id="f144e-231">Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="f144e-231">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="f144e-232">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="f144e-232">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="f144e-233">No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="f144e-233">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="f144e-234">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="f144e-234">Known issues</span></span>

- <span data-ttu-id="f144e-235">As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-235">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="f144e-236">Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.</span><span class="sxs-lookup"><span data-stu-id="f144e-236">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="f144e-237">Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não serão aceitas.</span><span class="sxs-lookup"><span data-stu-id="f144e-237">Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="f144e-238">Implantação por meio do Portal de administração do Office 365 e AppSource ainda não estão habilitados.</span><span class="sxs-lookup"><span data-stu-id="f144e-238">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="f144e-239">Funções personalizadas no Excel Online podem deixar de funcionar durante uma sessão após um período de inatividade.</span><span class="sxs-lookup"><span data-stu-id="f144e-239">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="f144e-240">Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="f144e-240">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="f144e-241">Você pode ver o resultado temporário **# OBTENDO_DADOS** nas células de uma planilha, se você tiver vários suplementos em execução no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="f144e-241">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="f144e-242">Feche todas as janelas do Excel e reinicie o Excel.</span><span class="sxs-lookup"><span data-stu-id="f144e-242">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="f144e-243">Ferramentas de depuração especificas para funções personalizadas podem estar disponíveis no futuro.</span><span class="sxs-lookup"><span data-stu-id="f144e-243">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="f144e-244">Enquanto isso, você pode depurar no Excel Online usando ferramentas de desenvolvedor F12.</span><span class="sxs-lookup"><span data-stu-id="f144e-244">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="f144e-245">Ver mais detalhes em [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="f144e-245">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>
- <span data-ttu-id="f144e-246">Na versão de 32 bits do Office 365 para Insiders 1901 de *dezembro* (compilação 11128.20000), as funções personalizadas podem não funcionar corretamente.</span><span class="sxs-lookup"><span data-stu-id="f144e-246">In the 32 bit version of the Office 365 *December* Insiders Version 1901 (Build 11128.20000),  Custom Functions may not work properly.</span></span> <span data-ttu-id="f144e-247">Em alguns casos, você pode solucionar esse erro baixando o arquivo em https://github.com/OfficeDev/Excel-Custom-Functions/blob/december-insiders-workaround/excel-udf-host.win32.bundle.</span><span class="sxs-lookup"><span data-stu-id="f144e-247">In some cases you can workaround this bug by downloading the file at https://github.com/OfficeDev/Excel-Custom-Functions/blob/december-insiders-workaround/excel-udf-host.win32.bundle.</span></span> <span data-ttu-id="f144e-248">Em seguida, copie a pasta “C:\ Arquivos de Programas (x86)\Microsoft Office\root\Office16”.</span><span class="sxs-lookup"><span data-stu-id="f144e-248">Then, copy it your "C:\Program Files (x86)\Microsoft Office\root\Office16" folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="f144e-249">Confira também</span><span class="sxs-lookup"><span data-stu-id="f144e-249">See also</span></span>

* [<span data-ttu-id="f144e-250">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f144e-250">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f144e-251">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="f144e-251">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="f144e-252">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="f144e-252">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="f144e-253">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f144e-253">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="f144e-254">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="f144e-254">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

