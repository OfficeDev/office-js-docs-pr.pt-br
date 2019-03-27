---
ms.date: 03/19/2019
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (versão prévia)
localization_priority: Priority
ms.openlocfilehash: ac3410267da415c4d567092da2e653fcffd10b72
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870447"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="04160-103">Criar funções personalizadas no Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="04160-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="04160-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="04160-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="04160-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="04160-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="04160-106">Este artigo descreve como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="04160-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="04160-107">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="04160-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="04160-108">A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par dos números que o usuário especifica como parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="04160-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="04160-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="04160-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="04160-110">A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="04160-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="04160-111">Componentes de um projeto de suplemento de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="04160-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="04160-112">Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas do Excel em um projeto, você verá os seguintes arquivos no projeto que o gerador cria:</span><span class="sxs-lookup"><span data-stu-id="04160-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="04160-113">File</span><span class="sxs-lookup"><span data-stu-id="04160-113">File</span></span> | <span data-ttu-id="04160-114">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="04160-114">File format</span></span> | <span data-ttu-id="04160-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="04160-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="04160-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="04160-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="04160-117">ou</span><span class="sxs-lookup"><span data-stu-id="04160-117">or</span></span><br/><span data-ttu-id="04160-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="04160-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="04160-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="04160-119">JavaScript</span></span><br/><span data-ttu-id="04160-120">ou</span><span class="sxs-lookup"><span data-stu-id="04160-120">or</span></span><br/><span data-ttu-id="04160-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="04160-121">TypeScript</span></span> | <span data-ttu-id="04160-122">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="04160-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="04160-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="04160-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="04160-124">JSON</span><span class="sxs-lookup"><span data-stu-id="04160-124">JSON</span></span> | <span data-ttu-id="04160-125">Contém metadados que descrevem funções personalizadas e permitem que o Excel registre funções personalizadas e as disponibilize para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="04160-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="04160-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="04160-126">**./index.html**</span></span> | <span data-ttu-id="04160-127">HTML</span><span class="sxs-lookup"><span data-stu-id="04160-127">HTML</span></span> | <span data-ttu-id="04160-128">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="04160-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="04160-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="04160-129">**./manifest.xml**</span></span> | <span data-ttu-id="04160-130">XML</span><span class="sxs-lookup"><span data-stu-id="04160-130">XML</span></span> | <span data-ttu-id="04160-131">Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="04160-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="04160-132">As seções a seguir fornecem mais informações sobre esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="04160-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="04160-133">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="04160-133">Script file</span></span>

<span data-ttu-id="04160-134">O arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** no projeto gerador que o Yo Office cria) contém o código que define funções personalizadas e mapeia os nomes da funções personalizadas aos objetos em [arquivos de metadados JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="04160-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="04160-135">Por exemplo, o código a seguir define funções personalizadas `add` e `increment` e especifica as informações de mapeamento para as duas funções.</span><span class="sxs-lookup"><span data-stu-id="04160-135">For example, the following code defines the custom functions `add` and `increment` and then specifies association information for both functions.</span></span> <span data-ttu-id="04160-136">A `add` função está associada com o objeto no arquivo nos metadados JSON onde o valor da propriedade `id` é **Adicionar**e a `increment` função é associada com o objeto no arquivo metadados onde o valor da propriedade`id`é **INCREMENTO**.</span><span class="sxs-lookup"><span data-stu-id="04160-136">The `add` function is associated with the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is associated with the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="04160-137">Ver [Práticas recomendadas de funções personalizados](custom-functions-best-practices.md#associating-function-names-with-json-metadata) para saber mais como associar os nomes de função no arquivo de script para objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-137">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about associating function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="04160-138">Arquivo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="04160-138">JSON metadata file</span></span>

<span data-ttu-id="04160-139">O arquivo de metadados de funções personalizadas (**./config/customfunctions.json** no projeto gerador que o Yo Office cria) fornece informações exigidas pelo Excel para registrar funções personalizadas e disponibilizá-las aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="04160-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="04160-140">Funções personalizadas são registradas quando um usuário usar um suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="04160-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="04160-141">Depois disso, eles estão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento foi inicialmente executado.)</span><span class="sxs-lookup"><span data-stu-id="04160-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="04160-142">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="04160-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="04160-143">O seguinte código em **customfunctions.json** especifica os metadados para a função `add` e a função `increment` descritas anteriormente.</span><span class="sxs-lookup"><span data-stu-id="04160-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="04160-144">A tabela que segue o código fornece informações detalhadas sobre as propriedades individuais nesse objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="04160-145">Ver [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#associating-function-names-with-json-metadata) para saber mais sobre como especificar o valor das propriedades`id` e `name` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-145">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="04160-146">A tabela a seguir lista as propriedades normalmente presentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="04160-147">Para saber mais sobre o arquivo de metadados JSON, confira [Metadados de funções personalizadas](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="04160-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="04160-148">Propriedade</span><span class="sxs-lookup"><span data-stu-id="04160-148">Property</span></span>  | <span data-ttu-id="04160-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="04160-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="04160-150">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="04160-150">A unique ID for the function.</span></span> <span data-ttu-id="04160-151">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="04160-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="04160-152">Nome da função que o usuário final vê no Excel.</span><span class="sxs-lookup"><span data-stu-id="04160-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="04160-153">No Excel, o nome de função será prefixado pelo namespace de funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="04160-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="04160-154">A URL da página é exibida quando um usuário solicitar ajuda.</span><span class="sxs-lookup"><span data-stu-id="04160-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="04160-155">Descreve o que faz a função.</span><span class="sxs-lookup"><span data-stu-id="04160-155">Describes what the function does.</span></span> <span data-ttu-id="04160-156">Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu de preenchimento automático do Excel.</span><span class="sxs-lookup"><span data-stu-id="04160-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="04160-157">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="04160-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="04160-158">Para obter informações detalhadas sobre esse objeto, consulte [resultado](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="04160-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="04160-159">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="04160-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="04160-160">Para obter informações detalhadas sobre esse objeto, consulte [parâmetros](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="04160-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="04160-161">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="04160-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="04160-162">Confira mais informações sobre como essa propriedade pode ser usada em [funções de Streaming](#streaming-functions) e [cancelar uma função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="04160-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [canceling a function](#canceling-a-function).</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="04160-163">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="04160-163">Manifest file</span></span>

<span data-ttu-id="04160-164">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="04160-165">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="04160-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
> <span data-ttu-id="04160-166">Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="04160-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="04160-167">O namespace da função vem antes do nome da função e são separados por um ponto.</span><span class="sxs-lookup"><span data-stu-id="04160-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="04160-168">Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="04160-169">O namespace deve ser usado como identificador para o as sua empresa ou suplemento.</span><span class="sxs-lookup"><span data-stu-id="04160-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="04160-170">Um namespace pode conter apenas caracteres alfanuméricos e períodos.</span><span class="sxs-lookup"><span data-stu-id="04160-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="04160-171">Como declarar uma função volátil</span><span class="sxs-lookup"><span data-stu-id="04160-171">Declaring a volatile function</span></span>

<span data-ttu-id="04160-172">As [funções voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) são funções nas quais o valor muda de momento a momento, mesmo que nenhum dos argumentos da função tenha mudado.</span><span class="sxs-lookup"><span data-stu-id="04160-172">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="04160-173">Essas funções são recalculadas sempre que o Excel recalcular.</span><span class="sxs-lookup"><span data-stu-id="04160-173">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="04160-174">Por exemplo, imagine uma célula que chame a função `NOW`.</span><span class="sxs-lookup"><span data-stu-id="04160-174">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="04160-175">Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.</span><span class="sxs-lookup"><span data-stu-id="04160-175">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="04160-176">O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="04160-176">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="04160-177">Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="04160-177">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="04160-178">As funções personalizadas permitem que você crie suas próprias funções voláteis, que podem ser úteis ao lidar com datas, horas, números aleatórios e modelagem.</span><span class="sxs-lookup"><span data-stu-id="04160-178">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="04160-179">Por exemplo, as simulações de Monte Carlo exigem a geração de entradas aleatórias para determinar uma solução ideal.</span><span class="sxs-lookup"><span data-stu-id="04160-179">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="04160-180">Para declarar uma função volátil, adicione `"volatile": true` no objeto `options` para a função no arquivo JSON de metadados, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="04160-180">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="04160-181">Observe que uma função não pode ser marcada como `"streaming": true` e `"volatile": true`; em casos em que ambas estejam marcadas com `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="04160-181">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="04160-182">Salvar e compartilhar estado</span><span class="sxs-lookup"><span data-stu-id="04160-182">Saving and sharing state</span></span>

<span data-ttu-id="04160-183">Funções personalizadas podem salvar os dados em variáveis, que podem ser usadas em chamadas subsequentes.</span><span class="sxs-lookup"><span data-stu-id="04160-183">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="04160-184">O estado salvo é útil quando os usuários solicitam a mesma função personalizada usando mais de uma célula, porque todas as ocorrências da função podem acessar o estado.</span><span class="sxs-lookup"><span data-stu-id="04160-184">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="04160-185">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="04160-185">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="04160-186">O código a seguir mostra uma implementação da função de streaming de temperatura que salva o estado globalmente.</span><span class="sxs-lookup"><span data-stu-id="04160-186">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="04160-187">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="04160-187">Note the following about this code:</span></span>

- <span data-ttu-id="04160-188">A função `streamTemperature` atualiza o valor de temperatura exibido na célula a cada segundo e ele usa a variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="04160-188">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="04160-189">Como `streamTemperature` é uma função de streaming, ela implementa um identificador de cancelamento que será executado quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="04160-189">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="04160-190">Se um usuário ligar a função`streamTemperature` de várias células no Excel, a função `streamTemperature` lê os dados a partir da mesma`savedTemperatures` variável toda vez que ela for executada.</span><span class="sxs-lookup"><span data-stu-id="04160-190">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="04160-191">`refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo e armazena o resultado na variável`savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="04160-191">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="04160-192">Como a função`refreshTemperature` não é exibida para os usuários finais no Excel, não é necessário ser registrado no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-192">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="04160-193">Coautoria</span><span class="sxs-lookup"><span data-stu-id="04160-193">Coauthoring</span></span>

<span data-ttu-id="04160-194">O Excel Online e o Excel para Windows com uma assinatura do Office 365 permitem editar documentos em coautoria, e esse recurso funciona com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="04160-194">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="04160-195">Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="04160-195">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="04160-196">Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.</span><span class="sxs-lookup"><span data-stu-id="04160-196">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="04160-197">Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="04160-197">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="04160-198">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="04160-198">Working with ranges of data</span></span>

<span data-ttu-id="04160-199">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="04160-199">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="04160-200">Em JavaScript, um intervalo de dados é representado como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="04160-200">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="04160-201">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="04160-201">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="04160-202">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="04160-202">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="04160-203">Observe que, nos metadados JSON dessa função, você deve definir o parâmetro `type` propriedade para `matrix`.</span><span class="sxs-lookup"><span data-stu-id="04160-203">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="04160-204">Determinar quais células chamadas de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="04160-204">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="04160-205">Em alguns casos, você precisará obter o endereço da célula invocada na sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="04160-205">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="04160-206">Isso pode ser útil para os seguintes tipos de cenários:</span><span class="sxs-lookup"><span data-stu-id="04160-206">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="04160-207">Formatação de intervalos: Use o endereço da célula como a chave para armazenar informações em [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="04160-207">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="04160-208">Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="04160-208">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="04160-209">Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `AsyncStorage` usando `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="04160-209">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="04160-210">Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.</span><span class="sxs-lookup"><span data-stu-id="04160-210">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="04160-211">As informações sobre o endereço de uma célula serão expostas somente se `requiresAddress` estiver marcado como `true` no arquivo de metadados JSON da função.</span><span class="sxs-lookup"><span data-stu-id="04160-211">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="04160-212">A seguir, um exemplo disso:</span><span class="sxs-lookup"><span data-stu-id="04160-212">The following sample gives an example of this:</span></span>

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

<span data-ttu-id="04160-213">No arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts**), também será necessário adicionar uma função `getAddress` para encontrar o endereço de uma célula.</span><span class="sxs-lookup"><span data-stu-id="04160-213">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="04160-214">Essa função pode ter parâmetros, conforme mostrado no exemplo a seguir como `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="04160-214">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="04160-215">O último parâmetro sempre será `invocationContext`, um objeto com o local da célula que o Excel passa quando `requiresAddress` é marcado como `true` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="04160-215">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="04160-216">Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="04160-216">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="04160-217">Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="04160-217">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="04160-218">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="04160-218">Known issues</span></span>

<span data-ttu-id="04160-219">Veja os problemas conhecidos no nosso [GitHub de funções do Excel personalizado repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="04160-219">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="04160-220">Confira também</span><span class="sxs-lookup"><span data-stu-id="04160-220">See also</span></span>

* [<span data-ttu-id="04160-221">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="04160-221">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="04160-222">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="04160-222">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="04160-223">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="04160-223">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="04160-224">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="04160-224">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="04160-225">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="04160-225">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
