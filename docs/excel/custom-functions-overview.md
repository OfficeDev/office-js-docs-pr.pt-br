---
ms.date: 09/20/2018
description: Criar uma função personalizada no Excel usando o JavaScript.
title: Criar funções personalizadas no Excel (Visualização)
ms.openlocfilehash: b214329fe50955d0f39d50f674152f475ca24b4d
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005040"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="ebb14-103">Criar funções personalizadas no Excel (Visualização)</span><span class="sxs-lookup"><span data-stu-id="ebb14-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="ebb14-104">Funções personalizadas permitem que os desenvolvedores adicionem novas funções para o Excel, definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="ebb14-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="ebb14-105">Os usuários podem então acessar funções personalizadas como qualquer outra função nativa do Excel (como `SUM()`).</span><span class="sxs-lookup"><span data-stu-id="ebb14-105">Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`).</span></span> <span data-ttu-id="ebb14-106">Este artigo explica como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-106">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="ebb14-107">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="ebb14-108">A `CONTOSO.ADD42` função personalizada foi projetada para adicionar 42 ao par de números que o usuário especifica como parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="ebb14-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="ebb14-109">The following code defines the `ADD42` custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="ebb14-110">As funções personalizadas agora estão disponíveis no Developer Preview para Windows, Mac e Excel Online.</span><span class="sxs-lookup"><span data-stu-id="ebb14-110">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="ebb14-111">Para testá-las, conclua estas etapas:</span><span class="sxs-lookup"><span data-stu-id="ebb14-111">To try them, complete these steps:</span></span>

1. <span data-ttu-id="ebb14-112">Instale o Office (compilação 10827 no Windows ou 13.329 no Mac) e participe do programa [Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="ebb14-112">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="ebb14-113">Você deve se associar ao programa Office Insider para ter acesso às funções personalizadas; atualmente as funções personalizadas estão desabilitadas em todas as versões do Office, a menos que você seja um membro do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="ebb14-113">You must join the Office Insider program in order to have access to custom functions; currently, custom functions are disabled across all Office builds unless you are a member of the Office Insider program.</span></span>

2. <span data-ttu-id="ebb14-114">Use [Yo Office](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento de funções personalizadas do Excel e siga as instruções no [README OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) para usar o projeto.</span><span class="sxs-lookup"><span data-stu-id="ebb14-114">Use [Yo Office](https://github.com/OfficeDev/generator-office) to create an Excel Custom Functions add-in project, and then follow the instructions in the [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to use the project.</span></span>

3. <span data-ttu-id="ebb14-115">Digite `=CONTOSO.ADD42(1,2)` em qualquer célula de uma planilha do Excel e pressione **Enter** para executar a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ebb14-115">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

> [!NOTE]
> <span data-ttu-id="ebb14-116">A seção de [Problemas conhecidos](#known-issues) neste artigo especifica as limitações atuais de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ebb14-116">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="ebb14-117">Noções básicas</span><span class="sxs-lookup"><span data-stu-id="ebb14-117">Learn the basics</span></span>

<span data-ttu-id="ebb14-118">No projeto de funções personalizadas que você criou usando o [Office Yo](https://github.com/OfficeDev/generator-office), você verá os seguintes arquivos:</span><span class="sxs-lookup"><span data-stu-id="ebb14-118">In the custom functions project that you've created using [Yo Office](https://github.com/OfficeDev/generator-office), you’ll see the following files:</span></span>

| <span data-ttu-id="ebb14-119">Arquivo</span><span class="sxs-lookup"><span data-stu-id="ebb14-119">File</span></span> | <span data-ttu-id="ebb14-120">Formato do arquivo</span><span class="sxs-lookup"><span data-stu-id="ebb14-120">File format</span></span> | <span data-ttu-id="ebb14-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="ebb14-121">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="ebb14-122">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="ebb14-122">**./src/customfunctions.js**</span></span> | <span data-ttu-id="ebb14-123">JavaScript</span><span class="sxs-lookup"><span data-stu-id="ebb14-123">JavaScript</span></span> | <span data-ttu-id="ebb14-124">Contém o código que define as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ebb14-124">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="ebb14-125">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="ebb14-125">**./config/customfunctions.json**</span></span> | <span data-ttu-id="ebb14-126">JSON</span><span class="sxs-lookup"><span data-stu-id="ebb14-126">JSON</span></span> | <span data-ttu-id="ebb14-127">Contém metadados que descrevem as funções personalizadas e permitem ao Excel registrar as funções personalizadas para torná-las disponíveis aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="ebb14-127">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="ebb14-128">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="ebb14-128">**./index.html**</span></span> | <span data-ttu-id="ebb14-129">HTML</span><span class="sxs-lookup"><span data-stu-id="ebb14-129">HTML</span></span> | <span data-ttu-id="ebb14-130">Fornece uma referência de &lt;script&gt; ao arquivo JavaScript que define as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ebb14-130">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="ebb14-131">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="ebb14-131">**Manifest.XML**</span></span> | <span data-ttu-id="ebb14-132">XML</span><span class="sxs-lookup"><span data-stu-id="ebb14-132">XML</span></span> | <span data-ttu-id="ebb14-133">Especifica o namespace para todas as funções personalizadas dentro do suplemento e o local dos arquivos JavaScript, JSON e HTML que estão listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="ebb14-133">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

### <a name="manifest-file-manifestxml"></a><span data-ttu-id="ebb14-134">Arquivo de manifesto (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="ebb14-134">Manifest file (manifest.xml)</span></span>

<span data-ttu-id="ebb14-135">O arquivo de manifesto XML para um suplemento que define as funções personalizadas especifica o namespace para todas as funções personalizadas dentro do suplemento e do local dos arquivos JavaScript, JSON e HTML.</span><span class="sxs-lookup"><span data-stu-id="ebb14-135">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="ebb14-136">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir em manifesto de um suplemento para habilitar o Excel a executar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ebb14-136">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="ebb14-137">As funções do Excel estão anexadas pelo namespace especificado em seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ebb14-137">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="ebb14-138">Um namespace de uma função vem antes do nome da função e são separados por um período.</span><span class="sxs-lookup"><span data-stu-id="ebb14-138">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="ebb14-139">Por exemplo, para chamar a função `ADD42()` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque CONTOSO é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="ebb14-139">For example, to call the function `ADD42()` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="ebb14-140">O prefixo de namespace deve ser usado como identificador para a sua empresa ou para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ebb14-140">The prefix is intended to be used as an identifier for your add-in.</span></span> 

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="ebb14-141">Arquivo JSON (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="ebb14-141">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="ebb14-142">Um arquivo de metadados de funções personalizadas fornece as informações que o Excel precisa para registrar as funções personalizadas e torná-las disponíveis aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="ebb14-142">A custom functions metadata file provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="ebb14-143">As funções personalizadas são registradas quando um usuário executa o suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="ebb14-143">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="ebb14-144">Depois disso, elas estarão disponíveis para esse mesmo usuário em todas as pastas de trabalho (ou seja, não apenas na pasta de trabalho onde o suplemento inicialmente executou.)</span><span class="sxs-lookup"><span data-stu-id="ebb14-144">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="ebb14-145">As configurações do servidor que hospedam o arquivo JSON devem ter [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="ebb14-145">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="ebb14-146">O código a seguir em **customfunctions.json** especifica os metadados para a função `ADD42` descrita anteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-146">The following code in **customfunctions.json** specifies the metadata for the `ADD42` function that was described previously in this article.</span></span> <span data-ttu-id="ebb14-147">Esses metadados definem o nome da função, descrição, valor de retorno, parâmetros de entrada e demais dados.</span><span class="sxs-lookup"><span data-stu-id="ebb14-147">This metadata defines the function's name, description, return value, input parameters, and more.</span></span> <span data-ttu-id="ebb14-148">A tabela que segue o exemplo de código a seguir fornece informações detalhadas sobre as propriedades individuais dentro desse objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="ebb14-148">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span>

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
        }
    ]
}
```

<span data-ttu-id="ebb14-149">A tabela a seguir lista as propriedades que estão normalmente presentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="ebb14-149">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="ebb14-150">Para obter informações mais detalhadas sobre o arquivo de metadados JSON, incluindo opções não usadas no exemplo anterior, consulte [Metadados de funções personalizadas](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="ebb14-150">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="ebb14-151">Propriedade</span><span class="sxs-lookup"><span data-stu-id="ebb14-151">Property</span></span>  | <span data-ttu-id="ebb14-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="ebb14-152">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="ebb14-153">Uma ID exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-153">A unique ID for the group.</span></span> <span data-ttu-id="ebb14-154">Essa ID não deve ser alterada depois de ser definida.</span><span class="sxs-lookup"><span data-stu-id="ebb14-154">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="ebb14-155">O nome da função que é mostrado no menu de preenchimento automático à medida que um usuário digita uma fórmula dentro de uma célula.</span><span class="sxs-lookup"><span data-stu-id="ebb14-155">Name of the function that is shown in the autocomplete menu as a user types a formula within a cell.</span></span> <span data-ttu-id="ebb14-156">No menu de preenchimento automático, esse valor será prefixado pelo namespace das funções personalizadas especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ebb14-156">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="ebb14-157">URL de uma página que é exibida quando o usuário solicita ajuda.</span><span class="sxs-lookup"><span data-stu-id="ebb14-157">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="ebb14-158">Descreve o que significa a função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-158">Describes what the function does.</span></span> <span data-ttu-id="ebb14-159">Esse valor aparece como uma dica de ferramenta quando a função é o item selecionado no menu Preenchimento Automático dentro do Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-159">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="ebb14-160">Objeto que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="ebb14-161">O valor da propriedade filho `type` pode ser uma **sequência de caracteres**, **número** ou **booleano**.</span><span class="sxs-lookup"><span data-stu-id="ebb14-161">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="ebb14-162">O valor da propriedade filho `dimensionality` pode ser **scalar** ou **matrix** (uma matriz bidimensional de valores do `type` especificado).</span><span class="sxs-lookup"><span data-stu-id="ebb14-162">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="ebb14-163">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-163">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="ebb14-164">As propriedades filho `name` e `description` são usadas no Intellisense do Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-164">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="ebb14-165">As propriedades filho `type` e `dimensionality` são idênticas às propriedades filho do objeto `result` descrito anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="ebb14-165">The `type` and `dimensionality` child properties are identical to the child properties of the `result` object that is described previously in this table.</span></span> |
| `options` | <span data-ttu-id="ebb14-166">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-166">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="ebb14-167">Para obter mais informações sobre como essa propriedade pode ser usada, consulte [Funções em fluxo contínuo](#streamed-functions) e [Cancelamento](#canceling-a-function) mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-167">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="ebb14-168">Funções que retornam dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="ebb14-168">Functions that return data from external sources</span></span>

<span data-ttu-id="ebb14-169">Se uma função personalizada recupera dados de uma fonte externa, como web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="ebb14-169">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="ebb14-170">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-170">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="ebb14-171">Resolver a Promise com o valor final usando a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ebb14-171">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="ebb14-172">Funções personalizadas exibem um resultado temporário `#GETTING_DATA` na célula enquanto o Excel aguarda o resultado final.</span><span class="sxs-lookup"><span data-stu-id="ebb14-172">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="ebb14-173">Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.</span><span class="sxs-lookup"><span data-stu-id="ebb14-173">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="ebb14-174">No exemplo de código a seguir, a função personalizada `getTemperature()` recupera a temperatura atual de um termômetro.</span><span class="sxs-lookup"><span data-stu-id="ebb14-174">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="ebb14-175">Observe que `sendWebRequest` é uma função hipotética (não especificada aqui) que usa o XHR para chamar um serviço Web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="ebb14-175">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="ebb14-176">Funções de streaming</span><span class="sxs-lookup"><span data-stu-id="ebb14-176">Streamed functions</span></span>

<span data-ttu-id="ebb14-177">Funções personalizadas em fluxo contínuo permitem que você transmita para células repetidamente ao longo do tempo, sem exigir que um usuário solicite explicitamente o recálculo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-177">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="ebb14-178">O exemplo de código a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="ebb14-179">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="ebb14-179">Note the following about this code:</span></span>

- <span data-ttu-id="ebb14-180">O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.</span><span class="sxs-lookup"><span data-stu-id="ebb14-180">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="ebb14-181">O parâmetro final `handler` nunca é especificado no seu código de registro e nunca é exibido no menu de preenchimento automático para usuários do Excel ao inserir a função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-181">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="ebb14-182">É a função de retorno de chamada `setResult` usada para passar dados da função para o Excel para atualizar o valor de uma célula.</span><span class="sxs-lookup"><span data-stu-id="ebb14-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>

- <span data-ttu-id="ebb14-183">Para que o Excel passe a função `setResult` no objeto `handler`, você deve declarar suporte para fluxo contínuo durante o registro da função, definindo a opção `"stream": true` na propriedade `options` para a função personalizada no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="ebb14-183">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="ebb14-184">Cancelamento de uma função</span><span class="sxs-lookup"><span data-stu-id="ebb14-184">Canceling a function</span></span>

<span data-ttu-id="ebb14-185">Em alguns casos, talvez seja necessário cancelar a execução de uma função personalizada em fluxo contínuo para reduzir seu consumo de largura de banda, a memória de trabalho e a carga da CPU.</span><span class="sxs-lookup"><span data-stu-id="ebb14-185">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="ebb14-186">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="ebb14-186">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="ebb14-187">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="ebb14-187">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="ebb14-188">Quando um dos argumentos (entradas) para a função é alterado.</span><span class="sxs-lookup"><span data-stu-id="ebb14-188">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="ebb14-189">Nesse caso, uma nova chamada de função é disparada após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="ebb14-189">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="ebb14-190">O usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-190">The user triggers recalculation manually.</span></span> <span data-ttu-id="ebb14-191">Nesse caso, uma nova chamada de função é disparada após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="ebb14-191">In this case, a new function call is triggered in addition to the cancelation.</span></span>

> [!NOTE]
> <span data-ttu-id="ebb14-192">Você deve implementar um manipulador de cancelamento para todas as funções de fluxo contínuo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-192">You must implement a cancellation handler for every streaming function.</span></span>

<span data-ttu-id="ebb14-193">Para tornar uma função cancelável, defina a opção `"cancelable": true` na propriedade `options` para a função personalizada no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="ebb14-193">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="ebb14-194">O código a seguir mostra a mesma função `incrementValue` descrita anteriormente, mas desta vez com um manipulador de cancelamento implementado.</span><span class="sxs-lookup"><span data-stu-id="ebb14-194">The following code shows the same `incrementValue` function that was described previously, but this time with a cancellation handler implemented.</span></span> <span data-ttu-id="ebb14-195">Neste exemplo, `clearInterval()` será executado quando a função `incrementValue` for cancelada.</span><span class="sxs-lookup"><span data-stu-id="ebb14-195">In this example, `clearInterval()` will run when the `incrementValue` function is canceled.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="ebb14-196">Compartilhamento e salvamento de estado</span><span class="sxs-lookup"><span data-stu-id="ebb14-196">Saving and sharing state</span></span>

<span data-ttu-id="ebb14-197">Funções personalizadas podem salvar os dados em variáveis JavaScript globais.</span><span class="sxs-lookup"><span data-stu-id="ebb14-197">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="ebb14-198">Em chamadas subsequentes, sua função personalizada pode usar valores salvos nessas variáveis.</span><span class="sxs-lookup"><span data-stu-id="ebb14-198">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="ebb14-199">O estado salvo é útil quando os usuários adicionam a mesma função personalizada a mais de uma célula, porque todas as instâncias da função podem compartilhar o estado.</span><span class="sxs-lookup"><span data-stu-id="ebb14-199">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="ebb14-200">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="ebb14-200">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="ebb14-201">O exemplo de código a seguir mostra uma implementação da função anterior de fluxo contínuo de temperatura que salva o estado de forma global.</span><span class="sxs-lookup"><span data-stu-id="ebb14-201">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="ebb14-202">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="ebb14-202">Note the following about this code:</span></span>

- <span data-ttu-id="ebb14-203">`refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="ebb14-203">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="ebb14-204">Novas temperaturas são salvas na variável `savedTemperatures`, mas o valor da célula não é atualizado diretamente.</span><span class="sxs-lookup"><span data-stu-id="ebb14-204">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="ebb14-205">Não deve ser chamada diretamente de uma célula da planilha, *por isso não está registrada no arquivo JSON*.</span><span class="sxs-lookup"><span data-stu-id="ebb14-205">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="ebb14-206">`streamTemperature` atualiza os valores de temperatura exibidos na célula a cada segundo e usa variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="ebb14-206">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="ebb14-207">Deve ser registrada no arquivo JSON e nomeada com todas as letras maiúsculas, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="ebb14-207">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="ebb14-208">Os usuários podem chamar `streamTemperature` de várias células na interface de usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-208">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="ebb14-209">Cada chamada lê dados da mesma variável `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="ebb14-209">Each call reads data from the same `savedTemperatures` variable.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="ebb14-210">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="ebb14-210">Working with ranges of data</span></span>

<span data-ttu-id="ebb14-211">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou ela pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="ebb14-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="ebb14-212">No JavaScript, um intervalo de dados é representado como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="ebb14-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="ebb14-213">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="ebb14-214">A função a seguir aceita o parâmetro `values`, que é do tipo `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="ebb14-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="ebb14-215">Observe que nos metadados JSON para esta função, você definiria a propriedade `type` do parâmetro como `matrix`.</span><span class="sxs-lookup"><span data-stu-id="ebb14-215">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="ebb14-216">Lidar com erros</span><span class="sxs-lookup"><span data-stu-id="ebb14-216">Handling errors</span></span>

<span data-ttu-id="ebb14-217">Quando você criar um suplemento que defina funções personalizadas, certifique-se de incluir a lógica de manipulação de erro para considerar os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="ebb14-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="ebb14-218">O tratamento de erros de funções personalizadas é o mesmo que [o tratamento de erros para a API do JavaScript Excel em geral](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="ebb14-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="ebb14-219">No exemplo de código a seguir, `.catch` manipulará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="ebb14-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi/comments/" + x;

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

## <a name="known-issues"></a><span data-ttu-id="ebb14-220">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="ebb14-220">Known issues</span></span>

- <span data-ttu-id="ebb14-221">As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="ebb14-222">Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.</span><span class="sxs-lookup"><span data-stu-id="ebb14-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="ebb14-223">Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="ebb14-223">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="ebb14-224">A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.</span><span class="sxs-lookup"><span data-stu-id="ebb14-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="ebb14-225">Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade.</span><span class="sxs-lookup"><span data-stu-id="ebb14-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="ebb14-226">Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="ebb14-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="ebb14-227">Se você tiver vários suplementos em execução no Excel para Windows, você poderá ver o resultado temporário **#GETTING_DATA** dentro de células de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="ebb14-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="ebb14-228">Feche todas as janelas do Excel e reinicie o Excel.</span><span class="sxs-lookup"><span data-stu-id="ebb14-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="ebb14-229">Outras ferramentas de depuração para funções personalizadas podem estar disponíveis no futuro.</span><span class="sxs-lookup"><span data-stu-id="ebb14-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="ebb14-230">Enquanto isso, você pode depurar no Excel Online usando as ferramentas de desenvolvedor F12.</span><span class="sxs-lookup"><span data-stu-id="ebb14-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="ebb14-231">Consulte mais detalhes em [Práticas recomendadas para funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="ebb14-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="ebb14-232">Log de mudanças</span><span class="sxs-lookup"><span data-stu-id="ebb14-232">Changelog</span></span>

- <span data-ttu-id="ebb14-233">**7 de novembro de 2017**: Enviados\* exemplos e versão prévia de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ebb14-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="ebb14-234">**20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores</span><span class="sxs-lookup"><span data-stu-id="ebb14-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="ebb14-235">**28 de novembro de 2017**: Enviado\* suporte para cancelamento em funções assíncronas (requer alteração para funções de fluxo contínuo)</span><span class="sxs-lookup"><span data-stu-id="ebb14-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="ebb14-236">**7 de maio de 2018**: Enviado\* o suporte para Mac, Excel Online e funções síncronas executadas no processo</span><span class="sxs-lookup"><span data-stu-id="ebb14-236">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="ebb14-237">**20 de setembro de 2018**: Enviado suporte para tempo de execução do JavaScript para funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ebb14-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="ebb14-238">Para obter mais informações, consulte [Tempo de execução para funções personalizadas do Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="ebb14-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="ebb14-239">\* para o Canal Office Insiders</span><span class="sxs-lookup"><span data-stu-id="ebb14-239">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="ebb14-240">Confira também</span><span class="sxs-lookup"><span data-stu-id="ebb14-240">See also</span></span>

* [<span data-ttu-id="ebb14-241">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ebb14-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ebb14-242">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="ebb14-242">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="ebb14-243">Práticas recomendadas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ebb14-243">Custom functions best practices</span></span>](custom-functions-best-practices.md)
