---
ms.date: 10/09/2018
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (versão prévia)
ms.openlocfilehash: e52039f2618f793f688cd89c5d62bac0a8632667
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506116"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="23cc7-103">Criar funções personalizadas no Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="23cc7-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="23cc7-p101">As funções personalizadas permitem que desenvolvedores adicionem novas funções ao Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários podem acessar as funções personalizadas como fazem com qualquer função nativa no Excel, como `SUM()`. Este artigo descreve como criar funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="23cc7-p102">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel. A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par de números especificado pelo usuário como os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p102">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet. The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="23cc7-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="23cc7-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="23cc7-110">A seção de [Problemas conhecidos](#known-issues) mais adiante neste artigo especifica as limitações atuais das funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="23cc7-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="23cc7-111">Componentes de um projeto de suplemento de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="23cc7-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="23cc7-112">Se você usar o [gerador de Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento de funções personalizadas do Excel, verá os seguintes arquivos no projeto criado pelo gerador:</span><span class="sxs-lookup"><span data-stu-id="23cc7-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="23cc7-113">Arquivo</span><span class="sxs-lookup"><span data-stu-id="23cc7-113">File</span></span> | <span data-ttu-id="23cc7-114">Formato do arquivo</span><span class="sxs-lookup"><span data-stu-id="23cc7-114">File format</span></span> | <span data-ttu-id="23cc7-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="23cc7-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="23cc7-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="23cc7-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="23cc7-117">ou</span><span class="sxs-lookup"><span data-stu-id="23cc7-117">or</span></span><br/><span data-ttu-id="23cc7-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="23cc7-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="23cc7-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="23cc7-119">JavaScript</span></span><br/><span data-ttu-id="23cc7-120">ou</span><span class="sxs-lookup"><span data-stu-id="23cc7-120">or</span></span><br/><span data-ttu-id="23cc7-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="23cc7-121">TypeScript</span></span> | <span data-ttu-id="23cc7-122">Contém o código que define as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="23cc7-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="23cc7-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="23cc7-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="23cc7-124">JSON</span><span class="sxs-lookup"><span data-stu-id="23cc7-124">JSON</span></span> | <span data-ttu-id="23cc7-125">Contém metadados que descrevem as funções personalizadas e permitem que o Excel registre as funções personalizadas e as disponibilize para o usuário final.</span><span class="sxs-lookup"><span data-stu-id="23cc7-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="23cc7-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="23cc7-126">**./index.html**</span></span> | <span data-ttu-id="23cc7-127">HTML</span><span class="sxs-lookup"><span data-stu-id="23cc7-127">HTML</span></span> | <span data-ttu-id="23cc7-128">Fornece uma referência de &lt;script&gt; ao arquivo JavaScript que define as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="23cc7-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="23cc7-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="23cc7-129">**Manifest.XML**</span></span> | <span data-ttu-id="23cc7-130">XML</span><span class="sxs-lookup"><span data-stu-id="23cc7-130">XML</span></span> | <span data-ttu-id="23cc7-131">Especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="23cc7-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="23cc7-132">As seções a seguir fornecem mais informações sobre esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="23cc7-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="23cc7-133">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="23cc7-133">Script file</span></span> 

<span data-ttu-id="23cc7-134">O arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** no projeto criado pelo gerador Yo Office) contém o código que define as funções personalizadas e mapeia os nomes das funções personalizadas para os objetos no [arquivo de metadados JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="23cc7-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="23cc7-p103">Por exemplo, o código a seguir define as funções personalizadas `add` e `increment` e, em seguida, especifica informações de mapeamento para ambas as funções. A função `add` é mapeada para o objeto no arquivo de metadados JSON em que o valor da propriedade `id` é **ADD**, e a função `increment` é mapeada para o objeto no arquivo de metadados em que o valor da propriedade `id` é **INCREMENT**. Consulte as [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para obter mais informações sobre o mapeamento de nomes de funções no arquivo de script para objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p103">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions. The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="23cc7-138">Arquivo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="23cc7-138">JSON metadata file</span></span> 

<span data-ttu-id="23cc7-p104">O arquivo de metadados de funções personalizadas (**./config/customfunctions.json** no projeto criado pelo gerador Yo Office) fornece as informações que o Excel precisa para registrar as funções personalizadas e disponibilizá-las para os usuários finais. Funções personalizadas serão registradas quando um usuário executa um suplemento pela primeira vez. Depois disso, eles ficam disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento foi inicialmente executado.)</span><span class="sxs-lookup"><span data-stu-id="23cc7-p104">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users. Custom functions are registered when a user runs an add-in for the first time. After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="23cc7-142">As configurações do servidor que hospeda o arquivo JSON devem ter [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="23cc7-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="23cc7-p105">O código a seguir no **customfunctions.json** especifica os metadados para as funções `add` e  `increment`, descritas anteriormente. A tabela apresentada na sequência do exemplo código a seguir fornece informações detalhadas sobre as propriedades individuais dentro desse objeto JSON. Consulte [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para obter mais informações sobre como especificar o valor das propriedades `id` e `name` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p105">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously. The table that follows this code sample provides detailed information about the individual properties within this JSON object. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="23cc7-p106">A tabela a seguir lista as propriedades normalmente presentes no arquivo de metadados JSON. Para obter informações mais detalhadas sobre o arquivo de metadados JSON, consulte [Metadados de funções personalizadas](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="23cc7-p106">The following table lists the properties that are typically present in the JSON metadata file. For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="23cc7-148">Propriedade</span><span class="sxs-lookup"><span data-stu-id="23cc7-148">Property</span></span>  | <span data-ttu-id="23cc7-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="23cc7-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="23cc7-p107">Uma ID exclusiva para a função. Essa ID não deve ser alterada depois de definida.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p107">A unique ID for the function. This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="23cc7-p108">O nome da função exibido para os usuários finais no Excel. No Excel, esse nome de função será prefixado pelo namespace das funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="23cc7-p108">Name of the function that the end user sees in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="23cc7-154">URL da página que é exibida quando o usuário solicita ajuda.</span><span class="sxs-lookup"><span data-stu-id="23cc7-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="23cc7-p109">Descreve o que a função faz. Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu de preenchimento automático dentro do Excel.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p109">Describes what the function does. This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="23cc7-p110">Objeto que define o tipo de informação retornada pela função. O valor da propriedade filha `type` pode ser **string**, **number** ou **boolean**. O valor da propriedade filha `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado).</span><span class="sxs-lookup"><span data-stu-id="23cc7-p110">Object that defines the type of information that is returned by the function. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `parameters` | <span data-ttu-id="23cc7-p111">Uma matriz que define os parâmetros de entrada da função. As propriedades filhas `name` e `description` aparecem no IntelliSense do Excel. O valor da propriedade filha `type` pode ser **string**, **number** ou **boolean**. O valor da propriedade filha `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado).</span><span class="sxs-lookup"><span data-stu-id="23cc7-p111">Array that defines the input parameters for the function. The `name` and `description` child properties appear in the Excel intelliSense. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `options` | <span data-ttu-id="23cc7-164">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="23cc7-164">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="23cc7-165">Para obter mais informações sobre como essa propriedade pode ser usada, consulte [Funções de fluxo contínuo](#streaming-functions) e [Cancelamento de uma função](#canceling-a-function) mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="23cc7-165">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="23cc7-166">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="23cc7-166">Manifest file</span></span>

<span data-ttu-id="23cc7-p113">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto que criado pelo gerador Yo Office) especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML. A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p113">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files. The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
> <span data-ttu-id="23cc7-p114">Funções do Excel são pré-inseridas pelo namespace especificado em seu arquivo de manifesto XML. O namespace de uma função vem antes do nome dela e é separado por um ponto. Por exemplo, para chamar a função `ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque a CONTOSO é o namespace e `ADD42` é o nome da função especificado no arquivo JSON. O namespace funciona como um identificador para a sua empresa ou para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p114">Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="23cc7-173">Funções que retornam dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="23cc7-173">Functions that return data from external sources</span></span>

<span data-ttu-id="23cc7-174">Se uma função personalizada recupera dados de uma fonte externa, como web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="23cc7-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="23cc7-175">Retornar um Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="23cc7-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="23cc7-176">Resolver a Promessa com o valor final usando a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="23cc7-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="23cc7-p115">Funções personalizadas exibem um resultado temporário `#GETTING_DATA` na célula enquanto o Excel aguarda o resultado final. Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p115">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result. Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="23cc7-p116">No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro. Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa [XHR](custom-functions-runtime.md#xhr-example) para chamar um serviço Web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p116">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer. Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="23cc7-181">Funções de fluxo contínuo</span><span class="sxs-lookup"><span data-stu-id="23cc7-181">Streaming functions</span></span>

<span data-ttu-id="23cc7-182">Funções personalizadas de fluxo contínuo permitem que você atribua dados para as células repetidamente ao longo do tempo, sem que o usuário precise solicitar explicitamente uma atualização de dados.</span><span class="sxs-lookup"><span data-stu-id="23cc7-182">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="23cc7-183">O exemplo de código a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="23cc7-183">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="23cc7-184">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="23cc7-184">Note the following about this code:</span></span>

- <span data-ttu-id="23cc7-185">O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.</span><span class="sxs-lookup"><span data-stu-id="23cc7-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="23cc7-186">O segundo parâmetro de entrada, `handler`, não é exibido para o usuário final no Excel, quando ele seleciona a função do menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="23cc7-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="23cc7-187">O retorno de chamada `onCanceled` define a função que é executada quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="23cc7-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="23cc7-188">Você deve implementar um manipulador de cancelamento como este para qualquer função de fluxo contínuo.</span><span class="sxs-lookup"><span data-stu-id="23cc7-188">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="23cc7-189">Para obter mais informações, confira [Cancelamento de função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="23cc7-189">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="23cc7-190">Quando você especifica os metadados de uma função de fluxo contínuo no arquivo de metadados JSON, deve definir as propriedades `"cancelable": true` e `"stream": true` dentro do objeto `options`, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="23cc7-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="23cc7-191">Cancelamento de uma função</span><span class="sxs-lookup"><span data-stu-id="23cc7-191">Canceling a function</span></span>

<span data-ttu-id="23cc7-192">Em alguns casos, talvez seja necessário cancelar a execução de uma função personalizada de fluxo contínuo para reduzir seu consumo de largura de banda, a memória de trabalho e a carga da CPU.</span><span class="sxs-lookup"><span data-stu-id="23cc7-192">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="23cc7-193">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="23cc7-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="23cc7-194">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="23cc7-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="23cc7-p120">Quando um dos argumentos (entradas) da função é alterado. Nesse caso, uma nova chamada de função é acionada após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p120">When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="23cc7-p121">Quando o usuário aciona o recálculo manualmente. Nesse caso, uma nova chamada de função é acionada após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p121">When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="23cc7-p122">Para habilitar a capacidade de cancelar uma função, você deve implementar um manipulador de cancelamento dentro da função do JavaScript e especificar a propriedade `"cancelable": true` dentro do objeto `options` nos metadados JSON que descrevem a função. Os exemplos de código na seção anterior deste artigo fornecem um exemplo dessas técnicas.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p122">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function. The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="23cc7-201">Compartilhamento e salvamento de estado</span><span class="sxs-lookup"><span data-stu-id="23cc7-201">Saving and sharing state</span></span>

<span data-ttu-id="23cc7-p123">Funções personalizadas podem salvar dados em variáveis globais do JavaScript, que podem ser usadas em chamadas subsequentes. Um estado salvo é útil quando usuários chamam a mesma função personalizada a partir de mais de uma célula, porque todas as instâncias da função podem acessar o estado. Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da web para evitar fazer chamadas adicionais para o mesmo recurso da web.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p123">Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="23cc7-p124">O exemplo de código a seguir mostra a implementação de uma função de fluxo contínuo de temperatura que salva o estado globalmente. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="23cc7-p124">The following code sample shows an implementation of a temperature-streaming function that saves state globally. Note the following about this code:</span></span>

- <span data-ttu-id="23cc7-207">A função `streamTemperature` atualiza o valor de temperatura que é exibido na célula cada segundo e usa a variável `savedTemperatures` como sua fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="23cc7-207">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="23cc7-208">Como `streamTemperature` é uma função de fluxo contínuo, ela implementa um manipulador de cancelamento que será executado quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="23cc7-208">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="23cc7-209">Se um usuário chama a função `streamTemperature` a partir de várias células no Excel, a função `streamTemperature` lê os dados da mesma `savedTemperatures` variável cada vez que ela é executada.</span><span class="sxs-lookup"><span data-stu-id="23cc7-209">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="23cc7-210">A função `refreshTemperature` lê a temperatura de um determinado termômetro a cada segundo e armazena o resultado na variável `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="23cc7-210">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="23cc7-211">Como a função `refreshTemperature` não está exposta aos usuários finais no Excel, ela não precisa ser registrada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="23cc7-211">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="23cc7-212">Trabalhando com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="23cc7-212">Working with ranges of data</span></span>

<span data-ttu-id="23cc7-p126">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados. No JavaScript, um intervalo de dados é representado como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p126">Your custom function may accept a range of data as an input parameter, or it may return a range of data. In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="23cc7-p127">Por exemplo, suponha que a sua função retorne o segundo valor mais alto de um intervalo de números armazenados no Excel. A função a seguir aceita o parâmetro `values`, que é do tipo `Excel.CustomFunctionDimensionality.matrix`. Observe que nos metadados JSON dessa função, você faria definiria a propriedade `type` do parâmetro como `matrix`.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p127">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="23cc7-218">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="23cc7-218">Handling errors</span></span>

<span data-ttu-id="23cc7-p128">Ao construir um suplemento que define funções personalizadas, certifique-se de incluir lógica para tratamento de erros para lidar com erros em tempo de execução. O tratamento de erros para funções personalizadas funciona da mesma forma que [o tratamento de erros para a API JavaScript do Excel de maneira geral](excel-add-ins-error-handling.md). No exemplo de código a seguir, `.catch` tratará quaisquer erros que ocorram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p128">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="23cc7-222">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="23cc7-222">Known issues</span></span>

- <span data-ttu-id="23cc7-223">As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="23cc7-223">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="23cc7-224">Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.</span><span class="sxs-lookup"><span data-stu-id="23cc7-224">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="23cc7-225">Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="23cc7-225">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="23cc7-226">A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.</span><span class="sxs-lookup"><span data-stu-id="23cc7-226">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="23cc7-p129">Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade. Atualize a página do navegador (F5) e insira novamente a função personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p129">Custom functions in Excel Online may stop working during a session after a period of inactivity. Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="23cc7-p130">Se você tiver vários suplementos em execução no Excel para Windows, poderá ver o resultado temporário **#GETTING_DATA** nas células de uma planilha. Feche todas as janelas do Excel e reinicie o programa.</span><span class="sxs-lookup"><span data-stu-id="23cc7-p130">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows. Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="23cc7-p131">Ferramentas de depuração específicas para funções personalizadas podem ser disponibilizadas futuramente. Enquanto isso, você pode depurar no Excel Online usando as ferramentas de desenvolvedor F12. Veja mais detalhes em [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="23cc7-p131">Debugging tools specifically for custom functions may be available in the future. In the meantime, you can debug on Excel Online using F12 developer tools. See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="23cc7-234">Log de mudanças</span><span class="sxs-lookup"><span data-stu-id="23cc7-234">Changelog</span></span>

- <span data-ttu-id="23cc7-235">**7 de novembro de 2017**: Enviados\* exemplos e versão prévia de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="23cc7-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="23cc7-236">**20 de Novembro de 2017**: Correção de bug de compatibilidade para quem usa o build 8801 e posteriores</span><span class="sxs-lookup"><span data-stu-id="23cc7-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="23cc7-237">**28 de novembro de 2017**: Enviado\* suporte para cancelamento em funções assíncronas (requer alteração para funções de fluxo contínuo)</span><span class="sxs-lookup"><span data-stu-id="23cc7-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="23cc7-238">**7 de maio de 2018**: Enviado\* o suporte para Mac, Excel Online e funções síncronas executadas no processo</span><span class="sxs-lookup"><span data-stu-id="23cc7-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="23cc7-p132">**20 de setembro de 2018**: Enviado o suporte para funções personalizadas de tempo de execução do JavaScript. Para obter mais informações, consulte o [Funções personalizadas para tempo de execução do Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="23cc7-p132">**September 20, 2018**: Shipped support for custom functions JavaScript runtime. For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="23cc7-241">\* para o Canal Office Insiders</span><span class="sxs-lookup"><span data-stu-id="23cc7-241">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="23cc7-242">Confira também</span><span class="sxs-lookup"><span data-stu-id="23cc7-242">See also</span></span>

* [<span data-ttu-id="23cc7-243">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="23cc7-243">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="23cc7-244">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="23cc7-244">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="23cc7-245">Práticas recomendadas para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="23cc7-245">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="23cc7-246">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="23cc7-246">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)