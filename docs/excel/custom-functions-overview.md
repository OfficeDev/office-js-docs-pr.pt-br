---
ms.date: 09/27/2018
description: Criar uma função personalizada no Excel usando o JavaScript.
title: Criar funções personalizadas no Excel (visualização)
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348798"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="8545b-103">Criar funções personalizadas no Excel (visualização)</span><span class="sxs-lookup"><span data-stu-id="8545b-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="8545b-p101">As funções personalizadas permitem que desenvolvedores adicionem novas funções ao Excel definindo essas funções em JavaScript, como parte de um suplemento. Os usuários podem acessar as funções personalizadas como fazem com qualquer função nativa no Excel, como `SUM()`. Este artigo descreve como criar funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`). This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="8545b-107">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="8545b-108">A `CONTOSO.ADD42` função personalizada foi projetada para adicionar 42 ao par de números que o usuário especifica como parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="8545b-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="8545b-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="8545b-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="8545b-110">A seção de [Problemas conhecidos](#known-issues) mais adiante neste artigo especifica as limitações atuais das funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8545b-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="8545b-111">Componentes de um projeto de suplemento de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8545b-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="8545b-112">Se você usar o [gerador de Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento de funções personalizadas do Excel, verá os seguintes arquivos no projeto criado pelo gerador:</span><span class="sxs-lookup"><span data-stu-id="8545b-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="8545b-113">Arquivo</span><span class="sxs-lookup"><span data-stu-id="8545b-113">File</span></span> | <span data-ttu-id="8545b-114">Formato do arquivo</span><span class="sxs-lookup"><span data-stu-id="8545b-114">File format</span></span> | <span data-ttu-id="8545b-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="8545b-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="8545b-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="8545b-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="8545b-117">ou</span><span class="sxs-lookup"><span data-stu-id="8545b-117">or</span></span><br/><span data-ttu-id="8545b-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="8545b-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="8545b-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8545b-119">JavaScript</span></span><br/><span data-ttu-id="8545b-120">ou</span><span class="sxs-lookup"><span data-stu-id="8545b-120">or</span></span><br/><span data-ttu-id="8545b-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8545b-121">TypeScript</span></span> | <span data-ttu-id="8545b-122">Contém o código que define as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8545b-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="8545b-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="8545b-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="8545b-124">JSON</span><span class="sxs-lookup"><span data-stu-id="8545b-124">JSON</span></span> | <span data-ttu-id="8545b-125">Contém metadados que descrevem as funções personalizadas e permitem que o Excel registre as funções personalizadas e as disponibilize para o usuário final.</span><span class="sxs-lookup"><span data-stu-id="8545b-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="8545b-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="8545b-126">**./index.html**</span></span> | <span data-ttu-id="8545b-127">HTML</span><span class="sxs-lookup"><span data-stu-id="8545b-127">HTML</span></span> | <span data-ttu-id="8545b-128">Fornece uma referência de &lt;script&gt; ao arquivo JavaScript que define as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8545b-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="8545b-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="8545b-129">**Manifest.XML**</span></span> | <span data-ttu-id="8545b-130">XML</span><span class="sxs-lookup"><span data-stu-id="8545b-130">XML</span></span> | <span data-ttu-id="8545b-131">Especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="8545b-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="8545b-132">As seções a seguir fornecem mais informações sobre esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="8545b-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="8545b-133">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="8545b-133">Script file</span></span> 

<span data-ttu-id="8545b-134">O arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** no projeto criado pelo gerador Yo Office) contém o código que define as funções personalizadas e mapeia os nomes das funções personalizadas para os objetos no [arquivo de metadados JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="8545b-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="8545b-135">Por exemplo, o código a seguir define as funções personalizadas `add` e `increment` e, em seguida, especifica o mapeamento de ambas as funções.</span><span class="sxs-lookup"><span data-stu-id="8545b-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="8545b-136">A função `add` é mapeada para o objeto no arquivo de metadados JSON onde o valor da propriedade `id` é **ADD** e a função `increment` é mapeada para o objeto no arquivo de metadados onde o valor da propriedade  `id` é **INCREMENT**.</span><span class="sxs-lookup"><span data-stu-id="8545b-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="8545b-137">Confira o artigo [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para obter mais informações sobre o mapeamento de nomes de função no arquivo de script para objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="8545b-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a><span data-ttu-id="8545b-138">Arquivo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="8545b-138">JSON metadata file</span></span> 

<span data-ttu-id="8545b-139">O arquivo de metadados de funções personalizadas (**./config/customfunctions.json** no projeto criado pelo gerador Yo Office) fornece as informações que o Excel precisa para registrar as funções personalizadas e disponibilizá-las para o usuário final.</span><span class="sxs-lookup"><span data-stu-id="8545b-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="8545b-140">As funções personalizadas são registradas quando um usuário executa o suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="8545b-140">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="8545b-141">Depois disso, elas estarão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento inicialmente executou.)</span><span class="sxs-lookup"><span data-stu-id="8545b-141">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="8545b-142">As configurações do servidor que hospeda o arquivo JSON devem ter [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="8545b-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="8545b-143">O código a seguir em **customfunctions.json** especifica os metadados para as funções `add` e `increment` descritas anteriormente.</span><span class="sxs-lookup"><span data-stu-id="8545b-143">The following code in **customfunctions.json** specifies the metadata for the `add` function that was described previously in this article.</span></span> <span data-ttu-id="8545b-144">A tabela que segue o exemplo de código a seguir fornece informações detalhadas sobre as propriedades individuais dentro desse objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="8545b-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="8545b-145">Confira o artigo [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para obter mais informações sobre como especificar o valor das propriedades `id` e `name` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="8545b-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="8545b-146">A tabela a seguir lista as propriedades que geralmente estão presentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="8545b-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="8545b-147">Para obter informações mais detalhadas sobre o arquivo de metadados JSON, confira [Metadados de funções personalizadas](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="8545b-147">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="8545b-148">Propriedade</span><span class="sxs-lookup"><span data-stu-id="8545b-148">Property</span></span>  | <span data-ttu-id="8545b-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="8545b-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="8545b-150">Uma ID exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="8545b-150">A unique ID for the group.</span></span> <span data-ttu-id="8545b-151">Esse ID não deve ser alterado depois de definido.</span><span class="sxs-lookup"><span data-stu-id="8545b-151">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="8545b-152">Nome da função que o usuário final vê no Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="8545b-153">No Excel, esse nome de função será prefixado pelo namespace de funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="8545b-153">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="8545b-154">URL da página que é exibida quando o usuário solicita ajuda.</span><span class="sxs-lookup"><span data-stu-id="8545b-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="8545b-155">Descreve o que a função faz.</span><span class="sxs-lookup"><span data-stu-id="8545b-155">Describes what the function does.</span></span> <span data-ttu-id="8545b-156">Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu Preenchimento Automático dentro do Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="8545b-157">Objeto que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="8545b-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="8545b-158">O valor da propriedade filho `type` pode ser uma **sequência de caracteres**, **número** ou **booleano**.</span><span class="sxs-lookup"><span data-stu-id="8545b-158">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="8545b-159">O valor da propriedade filho `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado).</span><span class="sxs-lookup"><span data-stu-id="8545b-159">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="8545b-160">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="8545b-160">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="8545b-161">As propriedades filho `name` e `description` são usadas no Intellisense do Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-161">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="8545b-162">O valor da propriedade filho `type` pode ser **string**, **number** ou **boolean**.</span><span class="sxs-lookup"><span data-stu-id="8545b-162">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="8545b-163">O valor da propriedade filho `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado).</span><span class="sxs-lookup"><span data-stu-id="8545b-163">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `options` | <span data-ttu-id="8545b-164">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="8545b-164">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="8545b-165">Para obter mais informações sobre como essa propriedade pode ser usada, consulte [Funções de fluxo contínuo](#streamed-functions) e [Cancelamento de uma função](#canceling-a-function) mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="8545b-165">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="8545b-166">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="8545b-166">Manifest file</span></span>

<span data-ttu-id="8545b-167">O arquivo de manifesto XML de um suplemento que define funções personalizadas (**./manifest.xml** no projeto criado pelo gerador Yo Office) especifica o namespace de todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML.</span><span class="sxs-lookup"><span data-stu-id="8545b-167">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="8545b-168">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8545b-168">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="8545b-169">As funções do Excel recebem o prefixo do namespace especificado em seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8545b-169">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="8545b-170">Um namespace de uma função vem antes do nome da função e são separados por um período.</span><span class="sxs-lookup"><span data-stu-id="8545b-170">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="8545b-171">Por exemplo, para chamar a função `ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque CONTOSO é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="8545b-171">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="8545b-172">O namespace deve ser usado como identificador da sua empresa ou do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8545b-172">The prefix is intended to be used as an identifier for your add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="8545b-173">Funções que retornam dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="8545b-173">Functions that return data from external sources</span></span>

<span data-ttu-id="8545b-174">Se uma função personalizada recupera dados de uma fonte externa, como web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="8545b-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="8545b-175">Retornar um Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="8545b-176">Resolver a Promise com o valor final usando a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8545b-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="8545b-177">Funções personalizadas exibem um resultado temporário `#GETTING_DATA` na célula enquanto o Excel aguarda o resultado final.</span><span class="sxs-lookup"><span data-stu-id="8545b-177">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="8545b-178">Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.</span><span class="sxs-lookup"><span data-stu-id="8545b-178">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="8545b-179">No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro.</span><span class="sxs-lookup"><span data-stu-id="8545b-179">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="8545b-180">Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa [XHR](custom-functions-runtime.md#xhr) para chamar um serviço Web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="8545b-180">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="8545b-181">Funções de fluxo contínuo</span><span class="sxs-lookup"><span data-stu-id="8545b-181">Streamed functions</span></span>

<span data-ttu-id="8545b-182">Funções personalizadas de fluxo contínuo permitem que você atribua dados para as células repetidamente ao longo do tempo, sem que o usuário precise solicitar explicitamente uma atualização de dados.</span><span class="sxs-lookup"><span data-stu-id="8545b-182">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="8545b-183">O exemplo de código a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="8545b-183">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="8545b-184">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="8545b-184">Note the following about this code:</span></span>

- <span data-ttu-id="8545b-185">O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.</span><span class="sxs-lookup"><span data-stu-id="8545b-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="8545b-186">O segundo parâmetro de entrada, `handler`, não é exibido para o usuário final no Excel, quando ele seleciona a função do menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="8545b-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="8545b-187">O retorno de chamada `onCanceled` define a função que é executada quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="8545b-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="8545b-188">Você deve implementar um manipulador de cancelamento como este para qualquer função de fluxo contínuo.</span><span class="sxs-lookup"><span data-stu-id="8545b-188">You must implement a cancellation handler like this for any streamed function.</span></span> <span data-ttu-id="8545b-189">Para obter mais informações, confira [Cancelamento de função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="8545b-189">For more information, see [Canceling a function](#canceling-a-function).</span></span> 

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

<span data-ttu-id="8545b-190">Quando você especifica os metadados de uma função de fluxo contínuo no arquivo de metadados JSON, deve definir as propriedades `"cancelable": true` e `"stream": true` dentro do objeto `options`, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="8545b-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="8545b-191">Cancelamento de função</span><span class="sxs-lookup"><span data-stu-id="8545b-191">Canceling a function</span></span>

<span data-ttu-id="8545b-192">Em alguns casos, talvez seja necessário cancelar a execução de uma função personalizada em fluxo contínuo para reduzir seu consumo de largura de banda, a memória de trabalho e a carga da CPU.</span><span class="sxs-lookup"><span data-stu-id="8545b-192">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="8545b-193">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="8545b-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="8545b-194">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="8545b-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="8545b-195">Quando um dos argumentos (entradas) para a função é alterado.</span><span class="sxs-lookup"><span data-stu-id="8545b-195">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="8545b-196">Nesse caso, uma nova chamada de função é disparada após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="8545b-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="8545b-197">Quando o usuário aciona o recálculo manualmente.</span><span class="sxs-lookup"><span data-stu-id="8545b-197">When the user triggers recalculation manually.</span></span> <span data-ttu-id="8545b-198">Nesse caso, uma nova chamada de função é disparada após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="8545b-198">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="8545b-199">Para habilitar a capacidade de cancelar uma função, é preciso implementar um manipulador de cancelamento dentro da função JavaScript e especificar a propriedade `"cancelable": true` dentro do objeto `options` nos metadados JSON que descrevem a função.</span><span class="sxs-lookup"><span data-stu-id="8545b-199">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="8545b-200">Os exemplos de código na seção anterior deste artigo fornecem um exemplo dessas técnicas.</span><span class="sxs-lookup"><span data-stu-id="8545b-200">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="8545b-201">Compartilhamento e gravação de estado</span><span class="sxs-lookup"><span data-stu-id="8545b-201">Saving and sharing state</span></span>

<span data-ttu-id="8545b-202">Funções personalizadas podem salvar os dados em variáveis JavaScript globais.</span><span class="sxs-lookup"><span data-stu-id="8545b-202">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="8545b-203">Em chamadas subsequentes, sua função personalizada pode usar valores salvos nessas variáveis.</span><span class="sxs-lookup"><span data-stu-id="8545b-203">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="8545b-204">O estado salvo é útil quando os usuários adicionam a mesma função personalizada a mais de uma célula, porque todas as instâncias da função podem compartilhar o estado.</span><span class="sxs-lookup"><span data-stu-id="8545b-204">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="8545b-205">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="8545b-205">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="8545b-206">O exemplo de código a seguir mostra uma implementação da função de fluxo contínuo de temperatura que salva o estado de forma global.</span><span class="sxs-lookup"><span data-stu-id="8545b-206">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="8545b-207">Observe o seguinte sobre esse código:</span><span class="sxs-lookup"><span data-stu-id="8545b-207">Note the following about this code:</span></span>

- <span data-ttu-id="8545b-208">`refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="8545b-208">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="8545b-209">Novas temperaturas são salvas na variável `savedTemperatures`, mas não o valor da célula não é atualizado diretamente.</span><span class="sxs-lookup"><span data-stu-id="8545b-209">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="8545b-210">Não deve ser chamada diretamente de uma célula da planilha, *por isso não está registrada no arquivo JSON*.</span><span class="sxs-lookup"><span data-stu-id="8545b-210">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="8545b-211">`streamTemperature` atualiza os valores de temperatura exibidos na célula a cada segundo e usa variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="8545b-211">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="8545b-212">Deve ser registrada no arquivo JSON e nomeada com todas as letras maiúsculas, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="8545b-212">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="8545b-213">Os usuários podem chamar `streamTemperature` de várias células na interface de usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-213">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="8545b-214">Cada chamada lê dados da mesma variável `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="8545b-214">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="8545b-215">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="8545b-215">Working with ranges of data</span></span>

<span data-ttu-id="8545b-216">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou ela pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="8545b-216">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="8545b-217">No JavaScript, um intervalo de dados é representado como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="8545b-217">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="8545b-218">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-218">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="8545b-219">A função a seguir aceita o parâmetro `values`, que é do tipo `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="8545b-219">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="8545b-220">Observe que nos metadados JSON para esta função, você definiria a propriedade `type` do parâmetro como `matrix`.</span><span class="sxs-lookup"><span data-stu-id="8545b-220">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="8545b-221">Lidar com erros</span><span class="sxs-lookup"><span data-stu-id="8545b-221">Handling errors</span></span>

<span data-ttu-id="8545b-222">Quando você criar um suplemento que define funções personalizadas, certifique-se de incluir a lógica de manipulação de erro para considerar os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="8545b-222">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="8545b-223">O tratamento de erros de funções personalizadas é o mesmo que [tratamento de erros para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="8545b-223">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="8545b-224">No exemplo de código a seguir, `.catch` manipulará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="8545b-224">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="8545b-225">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="8545b-225">Known issues</span></span>

- <span data-ttu-id="8545b-226">As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="8545b-227">Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.</span><span class="sxs-lookup"><span data-stu-id="8545b-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="8545b-228">Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="8545b-228">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="8545b-229">A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.</span><span class="sxs-lookup"><span data-stu-id="8545b-229">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="8545b-230">Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade.</span><span class="sxs-lookup"><span data-stu-id="8545b-230">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="8545b-231">Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="8545b-231">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="8545b-232">Se você tiver vários suplementos em execução no Excel para Windows, você poderá ver o resultado temporário **#GETTING_DATA** dentro de células de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="8545b-232">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="8545b-233">Feche todas as janelas do Excel e reinicie o Excel.</span><span class="sxs-lookup"><span data-stu-id="8545b-233">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="8545b-234">Outras ferramentas de depuração para funções personalizadas podem estar disponíveis no futuro.</span><span class="sxs-lookup"><span data-stu-id="8545b-234">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="8545b-235">Enquanto isso, você pode depurar no Excel Online usando as ferramentas de desenvolvedor F12.</span><span class="sxs-lookup"><span data-stu-id="8545b-235">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="8545b-236">Consulte mais detalhes em [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="8545b-236">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="8545b-237">Log de mudanças</span><span class="sxs-lookup"><span data-stu-id="8545b-237">Changelog</span></span>

- <span data-ttu-id="8545b-238">**7 de novembro de 2017**: Enviados\* exemplos e versão prévia de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8545b-238">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="8545b-239">**20 de Novembro de 2017**: Correção de bug de compatibilidade para quem usa o build 8801 e posteriores</span><span class="sxs-lookup"><span data-stu-id="8545b-239">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="8545b-240">**28 de novembro de 2017**: Enviado\* suporte para cancelamento em funções assíncronas (requer alteração para funções de fluxo contínuo)</span><span class="sxs-lookup"><span data-stu-id="8545b-240">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="8545b-241">**7 de maio de 2018**: Enviado\* o suporte para Mac, Excel Online e funções síncronas executadas no processo</span><span class="sxs-lookup"><span data-stu-id="8545b-241">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="8545b-242">**20 de setembro de 2018**: Enviado suporte para tempo de execução do JavaScript para funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8545b-242">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="8545b-243">Para obter mais informações, consulte [Tempo de execução para funções personalizadas do Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8545b-243">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="8545b-244">\* para o Canal Office Insiders</span><span class="sxs-lookup"><span data-stu-id="8545b-244">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="8545b-245">Confira também</span><span class="sxs-lookup"><span data-stu-id="8545b-245">See also</span></span>

* [<span data-ttu-id="8545b-246">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8545b-246">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8545b-247">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="8545b-247">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="8545b-248">Práticas recomendadas para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8545b-248">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="8545b-249">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="8545b-249">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)