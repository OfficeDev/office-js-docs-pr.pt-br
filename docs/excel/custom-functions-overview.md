---
ms.date: 11/29/2018
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (Versão Prévia)
ms.openlocfilehash: daa0cea24473290a2bc1b5c931f2f7a00ddc8276
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156611"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="f0875-103">Criar funções personalizadas no Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="f0875-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="f0875-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="f0875-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="f0875-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="f0875-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="f0875-106">Este artigo descreve como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f0875-107">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="f0875-108">A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par dos números que o usuário especifica como parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="f0875-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="f0875-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="f0875-110">A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f0875-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="f0875-111">Componentes de um projeto de suplemento de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0875-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="f0875-112">Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas do Excel em um projeto, você verá os seguintes arquivos no projeto que o gerador cria:</span><span class="sxs-lookup"><span data-stu-id="f0875-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="f0875-113">File</span><span class="sxs-lookup"><span data-stu-id="f0875-113">File</span></span> | <span data-ttu-id="f0875-114">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="f0875-114">File format</span></span> | <span data-ttu-id="f0875-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="f0875-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="f0875-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="f0875-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="f0875-117">ou</span><span class="sxs-lookup"><span data-stu-id="f0875-117">or</span></span><br/><span data-ttu-id="f0875-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="f0875-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="f0875-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f0875-119">JavaScript</span></span><br/><span data-ttu-id="f0875-120">ou</span><span class="sxs-lookup"><span data-stu-id="f0875-120">or</span></span><br/><span data-ttu-id="f0875-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f0875-121">TypeScript</span></span> | <span data-ttu-id="f0875-122">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f0875-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="f0875-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="f0875-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="f0875-124">JSON</span><span class="sxs-lookup"><span data-stu-id="f0875-124">JSON</span></span> | <span data-ttu-id="f0875-125">Contém metadados que descrevem funções personalizadas e permitem que o Excel registre funções personalizadas e as disponibilize para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="f0875-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="f0875-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="f0875-126">**./index.html**</span></span> | <span data-ttu-id="f0875-127">HTML</span><span class="sxs-lookup"><span data-stu-id="f0875-127">HTML</span></span> | <span data-ttu-id="f0875-128">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f0875-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="f0875-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="f0875-129">**./manifest.xml**</span></span> | <span data-ttu-id="f0875-130">XML</span><span class="sxs-lookup"><span data-stu-id="f0875-130">XML</span></span> | <span data-ttu-id="f0875-131">Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="f0875-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="f0875-132">As seções a seguir fornecem mais informações sobre esses arquivos.</span><span class="sxs-lookup"><span data-stu-id="f0875-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="f0875-133">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="f0875-133">Script file</span></span>

<span data-ttu-id="f0875-134">O arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** no projeto gerador que o Yo Office cria) contém o código que define funções personalizadas e mapeia os nomes da funções personalizadas aos objetos em [arquivos de metadados JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="f0875-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="f0875-135">Por exemplo, o código a seguir define funções personalizadas `add` e `increment` e especifica as informações de mapeamento para as duas funções.</span><span class="sxs-lookup"><span data-stu-id="f0875-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="f0875-136">A `add` função mapeada para o objeto no arquivo nos metadados JSON onde o valor da `id` propriedade **Adicionar**e a `increment` função é mapeada para o objeto no arquivo metadados onde o valor da propriedade`id`é **INCREMENTO**.</span><span class="sxs-lookup"><span data-stu-id="f0875-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="f0875-137">Ver [Práticas recomendadas de funções personalizados](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para saber mais sobre mapeamento de nomes de função no arquivo de script para objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="f0875-138">Arquivo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="f0875-138">JSON metadata file</span></span> 

<span data-ttu-id="f0875-139">O arquivo de metadados de funções personalizadas (**./config/customfunctions.json** no projeto gerador que o Yo Office cria) fornece informações exigidas pelo Excel para registrar funções personalizadas e disponibilizá-las aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="f0875-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="f0875-140">Funções personalizadas são registradas quando um usuário usar um suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="f0875-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="f0875-141">Depois disso, eles estão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento foi inicialmente executado.)</span><span class="sxs-lookup"><span data-stu-id="f0875-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="f0875-142">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="f0875-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="f0875-143">O seguinte código em **customfunctions.json** especifica os metadados para a função `add` e a função `increment` descritas anteriormente.</span><span class="sxs-lookup"><span data-stu-id="f0875-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="f0875-144">A tabela que segue o código fornece informações detalhadas sobre as propriedades individuais nesse objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="f0875-145">Ver [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) para saber mais sobre como especificar o valor das propriedades`id` e `name` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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
        "stream": true,
        "volatile": false
      }
    }
  ]
}
```

<span data-ttu-id="f0875-146">A tabela a seguir lista as propriedades normalmente presentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="f0875-147">Para saber mais sobre o arquivo de metadados JSON, confira [Metadados de funções personalizadas](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="f0875-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="f0875-148">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f0875-148">Property</span></span>  | <span data-ttu-id="f0875-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="f0875-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="f0875-150">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-150">A unique ID for the function.</span></span> <span data-ttu-id="f0875-151">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="f0875-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="f0875-152">Nome da função que o usuário final vê no Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="f0875-153">No Excel, o nome de função será prefixado pelo namespace de funções personalizadas especificado no [arquivo de manifesto XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="f0875-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="f0875-154">A URL da página é exibida quando um usuário solicitar ajuda.</span><span class="sxs-lookup"><span data-stu-id="f0875-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="f0875-155">Descreve o que faz a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-155">Describes what the function does.</span></span> <span data-ttu-id="f0875-156">Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu de preenchimento automático do Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="f0875-157">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="f0875-158">Para obter informações detalhadas sobre esse objeto, consulte [resultado](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="f0875-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="f0875-159">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="f0875-160">Para obter informações detalhadas sobre esse objeto, consulte [parâmetros](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="f0875-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="f0875-161">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="f0875-162">Confira mais informações sobre como essa propriedade pode ser usada em [Função de streaming](#streaming-functions), [Como cancelar uma função](#canceling-a-function) e [Como declarar uma função volátil](#declaring-a-volatile-function) mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="f0875-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="f0875-163">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="f0875-163">Manifest file</span></span>

<span data-ttu-id="f0875-164">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="f0875-165">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f0875-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="f0875-166">Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f0875-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="f0875-167">O namespace da função vem antes do nome da função e são separados por um ponto.</span><span class="sxs-lookup"><span data-stu-id="f0875-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="f0875-168">Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="f0875-169">O namespace deve ser usado como identificador para o as sua empresa ou suplemento.</span><span class="sxs-lookup"><span data-stu-id="f0875-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="f0875-170">Um namespace pode conter apenas caracteres alfanuméricos e períodos.</span><span class="sxs-lookup"><span data-stu-id="f0875-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="f0875-171">Funções que retornam os dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="f0875-171">Functions that return data from external sources</span></span>

<span data-ttu-id="f0875-172">Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="f0875-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="f0875-173">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="f0875-174">Resolva a promessa com o uso da função retorno de chamada de valor final.</span><span class="sxs-lookup"><span data-stu-id="f0875-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="f0875-175">Exibição de funções personalizados mostra um `#GETTING_DATA` resultado temporário na célula enquanto o Excel espera do resultado final.</span><span class="sxs-lookup"><span data-stu-id="f0875-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="f0875-176">Os usuários podem interagir normalmente com o restante da planilha aguardando o resultado.</span><span class="sxs-lookup"><span data-stu-id="f0875-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="f0875-177">No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro.</span><span class="sxs-lookup"><span data-stu-id="f0875-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="f0875-178">Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa [XHR](custom-functions-runtime.md#xhr-example) para chamar um serviço web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="f0875-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="f0875-179">Funções Streaming</span><span class="sxs-lookup"><span data-stu-id="f0875-179">Streaming functions</span></span>

<span data-ttu-id="f0875-180">Funções personalizadas de streaming permitem a saída de dados das células repetidamente ao longo do tempo, sem a necessidade de um usuário explicitamente solicitar a atualização de dados.</span><span class="sxs-lookup"><span data-stu-id="f0875-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="f0875-181">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="f0875-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="f0875-182">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="f0875-182">Note the following about this code:</span></span>

- <span data-ttu-id="f0875-183">Cada novo valor usando o Excel automaticamente exibirá o `setResult` retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f0875-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="f0875-184">O segundo parâmetro de entrada, `handler`, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="f0875-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="f0875-185">O `onCanceled` retorno de chamada define a função que é executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="f0875-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="f0875-186">Implemente um identificador de cancelamento assim para qualquer função de streaming.</span><span class="sxs-lookup"><span data-stu-id="f0875-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="f0875-187">Para saber mais, confira [Cancelar uma função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="f0875-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="f0875-188">Quando você especifica os metadados para uma função streaming no arquivo de metadados JSON, você deve definir as propriedades `"cancelable": true` e `"stream": true` no `options` objeto, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f0875-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="f0875-189">Cancelar uma função</span><span class="sxs-lookup"><span data-stu-id="f0875-189">Canceling a function</span></span>

<span data-ttu-id="f0875-190">Em algumas situações, talvez seja necessário cancelar a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU.</span><span class="sxs-lookup"><span data-stu-id="f0875-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="f0875-191">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="f0875-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="f0875-192">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="f0875-192">When the user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="f0875-193">Quando é alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-193">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="f0875-194">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="f0875-194">In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="f0875-195">Quando o usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="f0875-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="f0875-196">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="f0875-196">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="f0875-197">Para habilitar o recurso cancelar uma função, implemente um identificador de cancelamento dentro da função JavaScript e especifique a propriedade `"cancelable": true` dentro do `options` objeto nos metadados JSON que descreve a função.</span><span class="sxs-lookup"><span data-stu-id="f0875-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="f0875-198">Amostras de código na seção anterior neste artigo fornecem um exemplo dessas técnicas.</span><span class="sxs-lookup"><span data-stu-id="f0875-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="f0875-199">Como declarar uma função volátil</span><span class="sxs-lookup"><span data-stu-id="f0875-199">Declaring a volatile function</span></span>

<span data-ttu-id="f0875-200">As [funções voláteis](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) são funções nas quais o valor muda de momento a momento, mesmo que nenhum dos argumentos da função tenha mudado.</span><span class="sxs-lookup"><span data-stu-id="f0875-200">[Volatile functions](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="f0875-201">Essas funções são recalculadas sempre que o Excel recalcular.</span><span class="sxs-lookup"><span data-stu-id="f0875-201">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="f0875-202">Por exemplo, imagine uma célula que chame a função `NOW`.</span><span class="sxs-lookup"><span data-stu-id="f0875-202">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="f0875-203">Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.</span><span class="sxs-lookup"><span data-stu-id="f0875-203">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="f0875-204">O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="f0875-204">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="f0875-205">Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](https://docs.microsoft.com/pt-BR/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="f0875-205">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](https://docs.microsoft.com/pt-BR/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>  
  
<span data-ttu-id="f0875-206">As funções personalizadas permitem que você crie suas próprias funções voláteis, que podem ser úteis ao lidar com datas, horas, números aleatórios e modelagem.</span><span class="sxs-lookup"><span data-stu-id="f0875-206">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="f0875-207">Por exemplo, as simulações de Monte Carlo exigem a geração de entradas aleatórias para determinar uma solução ideal.</span><span class="sxs-lookup"><span data-stu-id="f0875-207">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>  
  
<span data-ttu-id="f0875-208">Para declarar uma função volátil, adicione `"volatile": true` no objeto `options` para a função no arquivo JSON de metadados, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f0875-208">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="f0875-209">Observe que uma função não pode ser marcada como `"streaming": true` e `"volatile": true`; em casos em que ambas estejam marcadas com `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="f0875-209">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>  

```json
{
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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="f0875-210">Salvar e compartilhar estado</span><span class="sxs-lookup"><span data-stu-id="f0875-210">Saving and sharing state</span></span>

<span data-ttu-id="f0875-211">Funções personalizadas podem salvar os dados em variáveis, que podem ser usadas em chamadas subsequentes.</span><span class="sxs-lookup"><span data-stu-id="f0875-211">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="f0875-212">O estado salvo é útil quando os usuários solicitam a mesma função personalizada usando mais de uma célula, porque todas as ocorrências da função podem acessar o estado.</span><span class="sxs-lookup"><span data-stu-id="f0875-212">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="f0875-213">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="f0875-213">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="f0875-214">O código a seguir mostra uma implementação da função de streaming de temperatura que salva o estado globalmente.</span><span class="sxs-lookup"><span data-stu-id="f0875-214">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="f0875-215">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="f0875-215">Note the following about this code:</span></span>

- <span data-ttu-id="f0875-216">A função `streamTemperature` atualiza o valor de temperatura exibido na célula a cada segundo e ele usa a variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="f0875-216">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="f0875-217">Como `streamTemperature` é uma função de streaming, ela implementa um identificador de cancelamento que será executado quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="f0875-217">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="f0875-218">Se um usuário ligar a função`streamTemperature` de várias células no Excel, a função `streamTemperature` lê os dados a partir da mesma`savedTemperatures` variável toda vez que ela for executada.</span><span class="sxs-lookup"><span data-stu-id="f0875-218">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="f0875-219">`refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo e armazena o resultado na variável`savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="f0875-219">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="f0875-220">Como a função`refreshTemperature` não é exibida para os usuários finais no Excel, não é necessário ser registrado no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="f0875-220">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="f0875-221">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="f0875-221">Working with ranges of data</span></span>

<span data-ttu-id="f0875-222">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="f0875-222">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="f0875-223">Em JavaScript, um intervalo de dados é representado como uma matriz 2 multidimensional.</span><span class="sxs-lookup"><span data-stu-id="f0875-223">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="f0875-224">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-224">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f0875-225">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="f0875-225">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="f0875-226">Observe que, nos metadados JSON dessa função, você deve definir o parâmetro `type` propriedade para `matrix`.</span><span class="sxs-lookup"><span data-stu-id="f0875-226">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="f0875-227">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="f0875-227">Handling errors</span></span>

<span data-ttu-id="f0875-228">Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="f0875-228">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="f0875-229">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="f0875-229">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="f0875-230">No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="f0875-230">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="f0875-231">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="f0875-231">Known issues</span></span>

- <span data-ttu-id="f0875-232">As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-232">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="f0875-233">Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.</span><span class="sxs-lookup"><span data-stu-id="f0875-233">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="f0875-234">Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não serão aceitas.</span><span class="sxs-lookup"><span data-stu-id="f0875-234">Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="f0875-235">Implantação por meio do Portal de administração do Office 365 e AppSource ainda não estão habilitados.</span><span class="sxs-lookup"><span data-stu-id="f0875-235">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="f0875-236">Funções personalizadas no Excel Online podem deixar de funcionar durante uma sessão após um período de inatividade.</span><span class="sxs-lookup"><span data-stu-id="f0875-236">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="f0875-237">Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="f0875-237">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="f0875-238">Você pode ver o resultado temporário **# OBTENDO_DADOS** nas células de uma planilha, se você tiver vários suplementos em execução no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="f0875-238">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="f0875-239">Feche todas as janelas do Excel e reinicie o Excel.</span><span class="sxs-lookup"><span data-stu-id="f0875-239">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="f0875-240">Ferramentas de depuração especificas para funções personalizadas podem estar disponíveis no futuro.</span><span class="sxs-lookup"><span data-stu-id="f0875-240">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="f0875-241">Enquanto isso, você pode depurar no Excel Online usando ferramentas de desenvolvedor F12.</span><span class="sxs-lookup"><span data-stu-id="f0875-241">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="f0875-242">Ver mais detalhes em [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="f0875-242">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="f0875-243">Log de mudanças</span><span class="sxs-lookup"><span data-stu-id="f0875-243">Changelog</span></span>

- <span data-ttu-id="f0875-244">**7 de novembro de 2017**: enviados exemplos e visualizações de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0875-244">**Nov 7, 2017**: Shipped\* the custom functions preview and samples</span></span>
- <span data-ttu-id="f0875-245">**20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores</span><span class="sxs-lookup"><span data-stu-id="f0875-245">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="f0875-246">**28 de novembro de 2017**: enviado o suporte para cancelamento em funções assíncronas (requer a alteração de funções de streaming)</span><span class="sxs-lookup"><span data-stu-id="f0875-246">**Nov 28, 2017**: Shipped\* support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="f0875-247">**7 de maio de 2018**: Suporte enviado para Mac, Excel Online e funções síncronas em execução no processo</span><span class="sxs-lookup"><span data-stu-id="f0875-247">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="f0875-248">**20 de setembro de 2018**: Suporte enviado para funções personalizadas de tempo de execução do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f0875-248">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="f0875-249">Para saber mais, confira [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="f0875-249">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="f0875-250">**20 de outubro de 2018**: Com o [build do Insider de outubro](https://support.office.com/pt-BR/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), funções personalizadas agora exigem o parâmetro "id" na suas [funções personalizadas metadados](custom-functions-json.md) para área de trabalho do Windows e Online.</span><span class="sxs-lookup"><span data-stu-id="f0875-250">**October 20, 2018**: With the [October Insiders build](https://support.office.com/pt-BR/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="f0875-251">No Mac, esse parâmetro deve ser ignorado.</span><span class="sxs-lookup"><span data-stu-id="f0875-251">On Mac, this parameter should be ignored.</span></span>


<span data-ttu-id="f0875-252">Em \* canal[Office Insider](https://products.office.com/office-insider), (anteriormente chamado de "Insider – modo rápido")</span><span class="sxs-lookup"><span data-stu-id="f0875-252">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

## <a name="see-also"></a><span data-ttu-id="f0875-253">Confira também</span><span class="sxs-lookup"><span data-stu-id="f0875-253">See also</span></span>

* [<span data-ttu-id="f0875-254">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0875-254">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f0875-255">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="f0875-255">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="f0875-256">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="f0875-256">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="f0875-257">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="f0875-257">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
