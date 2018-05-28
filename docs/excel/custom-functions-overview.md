# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="00103-101">Criar fun??es personalizadas no Excel (Visualiza??o)</span><span class="sxs-lookup"><span data-stu-id="00103-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="00103-102">Fun??es personalizadas (semelhantes a fun??es definidas pelo usu?rio ou UDFs) permitem que os desenvolvedores adicionem qualquer fun??o JavaScript no Excel usando um suplemento.</span><span class="sxs-lookup"><span data-stu-id="00103-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="00103-103">Os usu?rios podem acessar fun??es personalizadas como qualquer outra fun??o nativa no Excel (como `=SUM()`).</span><span class="sxs-lookup"><span data-stu-id="00103-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="00103-104">Este artigo explica como criar fun??es personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="00103-105">A ilustra??o a seguir mostra como um usu?rio final pode inserir uma fun??o personalizada em uma c?lula.</span><span class="sxs-lookup"><span data-stu-id="00103-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="00103-106">A fun??o que adiciona 42 a um par de n?meros.</span><span class="sxs-lookup"><span data-stu-id="00103-106">Here?s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="00103-107">Aqui est? o c?digo para a mesma fun??o personalizada.</span><span class="sxs-lookup"><span data-stu-id="00103-107">Here?s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="00103-108">As fun??es personalizadas agora est?o dispon?veis no Developer Preview para Windows, Mac e Excel Online.</span><span class="sxs-lookup"><span data-stu-id="00103-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="00103-109">Siga estas etapas para experiment?-las:</span><span class="sxs-lookup"><span data-stu-id="00103-109">Follow these steps to try them:</span></span>

1.  <span data-ttu-id="00103-110">Instale o Office (compila??o 9325 no Windows ou 13.329 no Mac) e participe do programa [Office Insider](https://products.office.com/en-us/office-insider).</span><span class="sxs-lookup"><span data-stu-id="00103-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/en-us/office-insider) program.</span></span> <span data-ttu-id="00103-111">(Observe que n?o ? suficiente apenas obter a compila??o mais recente; o recurso ser? desabilitado em qualquer compila??o at? voc? ingressar no programa Insider)</span><span class="sxs-lookup"><span data-stu-id="00103-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2.  <span data-ttu-id="00103-112">Clone o reposit?rio [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) e siga as instru??es no README.md para iniciar o suplemento no Excel, fazer altera??es no c?digo e depurar.</span><span class="sxs-lookup"><span data-stu-id="00103-112">Clone the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repo and follow the instructions in the README.md to start the add-in in Excel, make changes in the code, and debug.</span></span>
3.  <span data-ttu-id="00103-113">Digite `=CONTOSO.ADD42(1,2)` em qualquer c?lula e pressione **Inserir** para executar a fun??o personalizada.</span><span class="sxs-lookup"><span data-stu-id="00103-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="00103-114">Confira a se??o **Problemas conhecidos** no final deste artigo, que inclui as limita??es atuais das fun??es personalizadas e que ser? atualizado com o tempo.</span><span class="sxs-lookup"><span data-stu-id="00103-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="00103-115">No??es b?sicas</span><span class="sxs-lookup"><span data-stu-id="00103-115">Learn the basics</span></span>

<span data-ttu-id="00103-116">No reposit?rio de exemplo clonado, voc? ver? os seguintes arquivos:</span><span class="sxs-lookup"><span data-stu-id="00103-116">In the cloned sample repo, you?ll see the following files:</span></span>

- <span data-ttu-id="00103-117">**customfunctions.js**, que cont?m o c?digo de fun??o personalizado (veja o exemplo de c?digo simples acima para a `ADD42` fun??o).</span><span class="sxs-lookup"><span data-stu-id="00103-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="00103-118">**customfunctions.json**, que cont?m o registro JSON que informa ao Excel sobre sua fun??o personalizada.</span><span class="sxs-lookup"><span data-stu-id="00103-118">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function.</span></span> <span data-ttu-id="00103-119">O registro faz com que suas fun??es personalizadas apare?am na lista de fun??es dispon?veis exibidas quando um usu?rio digita em uma c?lula.</span><span class="sxs-lookup"><span data-stu-id="00103-119">Registration makes your custom functions appear in the list of available functions displayed when users type in cells.</span></span>
- <span data-ttu-id="00103-120">**customfunctions.html**, que fornece um &lt;Script&gt; de refer?ncia para o arquivo JS.</span><span class="sxs-lookup"><span data-stu-id="00103-120">customfunctions.html, which provides a Script reference to customfunctions.js.</span></span> <span data-ttu-id="00103-121">Este arquivo n?o ? exibido na interface do usu?rio do Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-121">This file does not display UI in Excel.</span></span>
- <span data-ttu-id="00103-122">**customfunctions.xml**, que informa ao Excel a localiza??o dos arquivos HTML, JavaScript e JSON; e tamb?m especifica um namespace para todas as fun??es personalizadas instaladas com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="00103-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-customfunctionsjson"></a><span data-ttu-id="00103-123">Arquivo JSON (customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="00103-123">JSON file (customfunctions.json)</span></span>

<span data-ttu-id="00103-124">O c?digo a seguir em customfunctions.json especifica os metadados para a mesma fun??o `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="00103-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="00103-125">Informa??es de refer?ncia detalhadas para o arquivo JSON, incluindo op??es n?o usadas neste exemplo, est?o em [Registro de Fun??es Personalizadas JSON](https://dev.office.com/reference/add-ins/custom-functions-json).</span><span class="sxs-lookup"><span data-stu-id="00103-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](https://dev.office.com/reference/add-ins/custom-functions-json).</span></span>

<span data-ttu-id="00103-126">Observe que, para este exemplo:</span><span class="sxs-lookup"><span data-stu-id="00103-126">Note that for this example:</span></span>

- <span data-ttu-id="00103-127">H? apenas uma fun??o personalizada, portanto, h? apenas um membro da `functions` matriz.</span><span class="sxs-lookup"><span data-stu-id="00103-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="00103-128">A propriedade `name` define o nome da fun??o.</span><span class="sxs-lookup"><span data-stu-id="00103-128">The `name` property defines the function name.</span></span> <span data-ttu-id="00103-129">Como voc? viu no gif animado mostrado anteriormente, um namespace (`CONTOSO`) ? anexado ao nome da fun??o no menu de preenchimento autom?tico do Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="00103-130">Esse prefixo ? definido no manifesto do suplemento, descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="00103-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="00103-131">O prefixo e o nome da fun??o s?o separados por um ponto e, por conven??o, nomes de fun??o e prefixos s?o mai?sculos.</span><span class="sxs-lookup"><span data-stu-id="00103-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="00103-132">Para usar a fun??o personalizada, o usu?rio digita o namespace seguido pelo nome da fun??o (`ADD42`) em uma c?lula, neste caso `=CONTOSO.ADD42`.</span><span class="sxs-lookup"><span data-stu-id="00103-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="00103-133">O prefixo deve ser usado como identificador para a sua empresa ou para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="00103-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="00103-134">O `description` aparece no menu de preenchimento autom?tico do Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="00103-135">Quando o usu?rio solicitar ajuda para uma fun??o, o Excel abre um painel de tarefas e exibe a p?gina da Web encontrada no URL especificado em `helpUrl` .</span><span class="sxs-lookup"><span data-stu-id="00103-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="00103-136">A propriedade `result` especifica o tipo de informa??o retornada pela fun??o para o Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="00103-137">A propriedade filho `type` pode `"string"`, `"number"`ou `"boolean"`.</span><span class="sxs-lookup"><span data-stu-id="00103-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="00103-138">A propriedade `dimensionality` pode ser `scalar` ou `matrix` (uma matriz bidimensional de valores do `type` especificado.)</span><span class="sxs-lookup"><span data-stu-id="00103-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="00103-139">A matriz `parameters` especifica, *em ordenar*, o tipo de dado em cada par?metro que ? passado para a fun??o.</span><span class="sxs-lookup"><span data-stu-id="00103-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="00103-140">As propriedades filho `name` e `description` s?o usadas no intellisense do Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="00103-141">As propriedades filho `type` e `dimensionality` s?o id?nticas ?s propriedades filho da propriedade `result` descrita acima.</span><span class="sxs-lookup"><span data-stu-id="00103-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="00103-142">A propriedade `options` permite que voc? personalize alguns aspectos de como e quando o Excel executa a fun??o.</span><span class="sxs-lookup"><span data-stu-id="00103-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="00103-143">H? mais informa??es sobre essas op??es posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="00103-143">There is more information about these options later in this article.</span></span>

 ```js
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "name": "ADD42", 
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
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
            "options": {
                "sync": true
            }
        }
    ]
}
```

> [!NOTE]
> <span data-ttu-id="00103-144">As fun??es personalizadas s?o registradas quando um usu?rio executa o suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="00103-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="00103-145">Depois disso, eles estar?o dispon?veis, para o mesmo usu?rio, em todas as pastas de trabalho (n?o apenas naquela em que o suplemento foi executado inicialmente).</span><span class="sxs-lookup"><span data-stu-id="00103-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="00103-146">As configura??es do servidor para o arquivo JSON devem ter [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) habilitado para que fun??es personalizadas funcionem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="00103-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-customfunctionsxml"></a><span data-ttu-id="00103-147">Arquivo de manifesto (customfunctions.xml)</span><span class="sxs-lookup"><span data-stu-id="00103-147">Manifest file (customfunctions.xml)</span></span>


<span data-ttu-id="00103-148">O seguinte ? um exemplo da marca??o `<ExtensionPoint>` e `<Resources>` que voc? inclui no manifesto do suplemento para permitir que o Excel execute suas fun??es.</span><span class="sxs-lookup"><span data-stu-id="00103-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="00103-149">Observe o seguinte sobre essa marca??o:</span><span class="sxs-lookup"><span data-stu-id="00103-149">Note the following about this code:</span></span>

- <span data-ttu-id="00103-150">O elemento `<Script>` e a identifica??o do recurso correspondente especificam a localiza??o do arquivo JavaScript com suas fun??es.</span><span class="sxs-lookup"><span data-stu-id="00103-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="00103-151">O elemento `<Page>` e a identifica??o do recurso correspondente especificam a localiza??o da p?gina HTML do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="00103-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="00103-152">A p?gina HTML inclui uma marca `<Script>` que carrega o arquivo JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="00103-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="00103-153">A p?gina HTML ? uma p?gina oculta e nunca ? exibida na interface de usu?rio.</span><span class="sxs-lookup"><span data-stu-id="00103-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="00103-154">O elemento `<Metadata>` e a identifica??o do recurso correspondente especificam a localiza??o do arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="00103-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="00103-155">Um elemento `<Namespace>` e a identifica??o do recurso correspondente especificam o prefixo para todas as fun??es personalizadas no suplemento.</span><span class="sxs-lookup"><span data-stu-id="00103-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a><span data-ttu-id="00103-156">Inicializa??o de fun??es personalizadas</span><span class="sxs-lookup"><span data-stu-id="00103-156">Initializing custom functions</span></span>

<span data-ttu-id="00103-157">Seu c?digo deve inicializar o recurso de fun??es personalizadas antes de us?-lo.</span><span class="sxs-lookup"><span data-stu-id="00103-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="00103-158">Voc? pode fazer isso em uma marca de &lt;Script&gt; no arquivo HTML (customfunctions.html) ou na parte superior do arquivo JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="00103-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="00103-159">Na visualiza??o de fun??es personalizadas, voc? pode escolher entre duas sintaxes para a inicializa??o.</span><span class="sxs-lookup"><span data-stu-id="00103-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="00103-160">O arquivo HTML no reposit?rio usa a seguinte sintaxe:</span><span class="sxs-lookup"><span data-stu-id="00103-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="00103-161">Voc? tamb?m pode usar a seguinte sintaxe:</span><span class="sxs-lookup"><span data-stu-id="00103-161">You can also nest JOIN statements using the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="00103-162">Fun??es s?ncronas e ass?ncronas</span><span class="sxs-lookup"><span data-stu-id="00103-162">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="00103-163">A fun??o `ADD42` acima ? s?ncrona em rela??o ao Excel (designada pela configura??o da op??o `"sync": true` no arquivo JSON).</span><span class="sxs-lookup"><span data-stu-id="00103-163">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="00103-164">As fun??es s?ncronas oferecem desempenho r?pido porque s?o executadas no mesmo processo que o Excel e em paralelo durante o c?lculo multithreaded.</span><span class="sxs-lookup"><span data-stu-id="00103-164">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="00103-165">Por outro lado, se sua fun??o personalizada recupera dados da Web, ela dever? ser ass?ncrona em rela??o ao Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-165">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="00103-166">Fun??es ass?ncronas devem:</span><span class="sxs-lookup"><span data-stu-id="00103-166">Asynchronous functions must:</span></span>

1. <span data-ttu-id="00103-167">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-167">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="00103-168">Resolver a Promise com o valor final usando a fun??o de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00103-168">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="00103-169">O c?digo a seguir mostra um exemplo de uma fun??o ass?ncrona que recupera a temperatura de um term?metro.</span><span class="sxs-lookup"><span data-stu-id="00103-169">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="00103-170">Observe que `sendWebRequest` ? uma fun??o hipot?tica, n?o especificada aqui, que usa o XHR para chamar um servi?o da Web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="00103-170">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="00103-171">Fun??es ass?ncronas exibem um erro tempor?rio `GETTING_DATA` na c?lula enquanto o Excel aguarda o resultado final.</span><span class="sxs-lookup"><span data-stu-id="00103-171">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="00103-172">Os usu?rios podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.</span><span class="sxs-lookup"><span data-stu-id="00103-172">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="00103-173">Fun??es personalizadas s?o ass?ncronas por padr?o.</span><span class="sxs-lookup"><span data-stu-id="00103-173">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="00103-174">Para designar fun??es como s?ncronas, defina a op??o `"sync": true` na propriedade `options` para a fun??o personalizada no arquivo JSON de registro.</span><span class="sxs-lookup"><span data-stu-id="00103-174">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="00103-175">Fun??es de fluxo cont?nuo</span><span class="sxs-lookup"><span data-stu-id="00103-175">Streamed functions</span></span>

<span data-ttu-id="00103-176">Uma fun??o ass?ncrona pode ser de fluxo cont?nuo.</span><span class="sxs-lookup"><span data-stu-id="00103-176">An asynchronous function can be streamed.</span></span> <span data-ttu-id="00103-177">Fun??es personalizadas de fluxo cont?nuo permitem que voc? insira dados em c?lulas repetidamente ao longo do tempo, sem precisar esperar que o Excel ou os usu?rios solicitem rec?lculos.</span><span class="sxs-lookup"><span data-stu-id="00103-177">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="00103-178">O exemplo a seguir ? uma fun??o personalizada que adiciona um n?mero ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="00103-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="00103-179">Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="00103-179">Note the following about this code:</span></span>

- <span data-ttu-id="00103-180">O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.</span><span class="sxs-lookup"><span data-stu-id="00103-180">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="00103-181">O par?metro final, `caller`, nunca ? especificado no c?digo de registro e nunca ? exibido no menu de preenchimento autom?tico para usu?rios do Excel ao inserir a fun??o.</span><span class="sxs-lookup"><span data-stu-id="00103-181">For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="00103-182">? a fun??o de retorno de chamada `setResult` usada para passar dados da fun??o para o Excel para atualizar o valor de uma c?lula.</span><span class="sxs-lookup"><span data-stu-id="00103-182">It?s an object that contains a `setResult` callback function that?s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="00103-183">Para que o Excel passe a fun??o `setResult` no objeto `caller`, voc? deve declarar suporte para fluxo cont?nuo durante o registro da fun??o, definindo a op??o `"stream": true` na propriedade `options` para a fun??o personalizada no arquivo JSON de registro.</span><span class="sxs-lookup"><span data-stu-id="00103-183">In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="00103-184">Cancelamento</span><span class="sxs-lookup"><span data-stu-id="00103-184">Cancellation</span></span>

<span data-ttu-id="00103-185">Voc? pode cancelar fun??es e fun??es ass?ncronas de streaming.</span><span class="sxs-lookup"><span data-stu-id="00103-185">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="00103-186">? importante cancelar as chamadas de fun??o para reduzir o consumo de largura de banda, a mem?ria de trabalho e a carga da CPU.</span><span class="sxs-lookup"><span data-stu-id="00103-186">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="00103-187">O Excel cancela chamadas de fun??es nas seguintes situa??es:</span><span class="sxs-lookup"><span data-stu-id="00103-187">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="00103-188">O usu?rio edita ou exclui uma c?lula que faz refer?ncia ? fun??o.</span><span class="sxs-lookup"><span data-stu-id="00103-188">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="00103-189">? alterado um dos argumentos (entradas) para a fun??o.</span><span class="sxs-lookup"><span data-stu-id="00103-189">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="00103-190">Nesse caso, uma nova chamada de fun??o ? disparada, al?m do cancelamento.</span><span class="sxs-lookup"><span data-stu-id="00103-190">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="00103-p124">O usu?rio aciona um rec?lculo manualmente. Como no caso acima, uma nova chamada de fun??o ? disparada, al?m do cancelamento.</span><span class="sxs-lookup"><span data-stu-id="00103-p124">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="00103-193">Voc? *deve* implementar um manipulador de cancelamento para todas as fun??es de fluxo cont?nuo.</span><span class="sxs-lookup"><span data-stu-id="00103-193">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="00103-194">Fun??es ass?ncronas e que n?o sejam de fluxo cont?nuo podem ou n?o ser cancel?veis; a decis?o ? sua.</span><span class="sxs-lookup"><span data-stu-id="00103-194">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="00103-195">Fun??es s?ncronas n?o podem ser canceladas.</span><span class="sxs-lookup"><span data-stu-id="00103-195">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="00103-196">Para tornar uma fun??o cancel?vel, defina a op??o `"cancelable": true` na propriedade `options` para a fun??o personalizada no arquivo JSON de registro.</span><span class="sxs-lookup"><span data-stu-id="00103-196">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="00103-197">O c?digo a seguir mostra o exemplo anterior com o cancelamento implementado.</span><span class="sxs-lookup"><span data-stu-id="00103-197">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="00103-198">No c?digo, o objeto `caller` cont?m uma fun??o `onCanceled` que deve ser definida para cada fun??o personalizada cancel?vel.</span><span class="sxs-lookup"><span data-stu-id="00103-198">In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="00103-199">Compartilhamento e salvamento de estado</span><span class="sxs-lookup"><span data-stu-id="00103-199">Saving and sharing state</span></span>

<span data-ttu-id="00103-200">Fun??es personalizadas ass?ncronas podem salvar dados em vari?veis JavaScript globais.</span><span class="sxs-lookup"><span data-stu-id="00103-200">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="00103-201">Em chamadas subsequentes, sua fun??o personalizada pode usar os valores salvos nessas vari?veis.</span><span class="sxs-lookup"><span data-stu-id="00103-201">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="00103-202">O estado salvo ? ?til quando os usu?rios adicionam a mesma fun??o personalizada a mais de uma c?lula, porque todas as inst?ncias da fun??o podem compartilhar o estado.</span><span class="sxs-lookup"><span data-stu-id="00103-202">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="00103-203">Por exemplo, voc? pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="00103-203">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="00103-204">O c?digo a seguir mostra uma implementa??o da fun??o de fluxo cont?nuo anterior de temperatura que salva o estado de forma global.</span><span class="sxs-lookup"><span data-stu-id="00103-204">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="00103-205">Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="00103-205">Note the following about this code:</span></span>

- <span data-ttu-id="00103-206">`refreshTemperature` ? uma fun??o de fluxo cont?nuo que l? a temperatura de um determinado term?metro a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="00103-206">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="00103-207">Novas temperaturas s?o salvas na vari?vel `savedTemperatures`, mas o valor da c?lula n?o ? atualizado diretamente.</span><span class="sxs-lookup"><span data-stu-id="00103-207">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="00103-208">N?o deve ser chamada diretamente de uma c?lula da planilha, *por isso n?o est? registrada no arquivo JSON*.</span><span class="sxs-lookup"><span data-stu-id="00103-208">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="00103-209">`streamTemperature` atualiza os valores de temperatura exibidos na c?lula a cada segundo e usa vari?vel `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="00103-209">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="00103-210">Deve ser registrada no arquivo JSON e nomeada com todas as letras mai?sculas, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="00103-210">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="00103-211">Os usu?rios podem chamar `streamTemperature` de v?rias c?lulas na interface de usu?rio do Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-211">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="00103-212">Cada chamada l? dados da mesma vari?vel `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="00103-212">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
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

> [!NOTE]
> <span data-ttu-id="00103-213">Fun??es s?ncronas (designadas pela configura??o da op??o `"sync": true` no arquivo JSON) n?o podem compartilhar estado porque o Excel faz o paralelismo delas durante o c?lculo multithreaded.</span><span class="sxs-lookup"><span data-stu-id="00103-213">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="00103-214">Somente fun??es ass?ncronas podem compartilhar estado porque as fun??es s?ncronas de um suplemento compartilham o mesmo contexto JavaScript em cada sess?o.</span><span class="sxs-lookup"><span data-stu-id="00103-214">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="00103-215">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="00103-215">Working with ranges of data</span></span>

<span data-ttu-id="00103-216">Sua fun??o personalizada pode levar a um intervalo de dados como um par?metro ou voc? pode retornar um intervalo de dados de uma fun??o personalizada.</span><span class="sxs-lookup"><span data-stu-id="00103-216">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="00103-217">Por exemplo, suponha que sua fun??o retorne o segundo maior valor de um intervalo de n?meros armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-217">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="00103-218">A fun??o a seguir usa o par?metro `values`, que ? um tipo de par?metro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="00103-218">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="00103-219">Note que no registro JSON para esta fun??o, voc? definiria a propriedade `type` do par?metro para `matrix`.</span><span class="sxs-lookup"><span data-stu-id="00103-219">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
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

<span data-ttu-id="00103-220">Como pode ver, os intervalos s?o tratados em JavaScript como matrizes de matrizes de linhas (como uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="00103-220">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="00103-221">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="00103-221">Known issues</span></span>

- <span data-ttu-id="00103-222">Descri??es de par?metro e URLs de Ajuda ainda n?o s?o usados pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="00103-222">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="00103-223">Fun??es personalizadas n?o est?o atualmente dispon?veis no Excel para clientes m?veis.</span><span class="sxs-lookup"><span data-stu-id="00103-223">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="00103-224">Atualmente, os suplementos dependem de um processo de navegador oculto para executar fun??es personalizadas ass?ncronas.</span><span class="sxs-lookup"><span data-stu-id="00103-224">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="00103-225">No futuro, o JavaScript ser? executado diretamente em algumas plataformas para garantir que as fun??es personalizadas sejam mais r?pidas e usem menos mem?ria.</span><span class="sxs-lookup"><span data-stu-id="00103-225">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="00103-226">Al?m disso, a p?gina HTML referenciada pelo elemento `<Page>` no manifesto n?o ser? necess?ria para a maioria das plataformas, j? que o Excel executa o JavaScript diretamente.</span><span class="sxs-lookup"><span data-stu-id="00103-226">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won?t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="00103-227">Para se preparar para essa altera??o, certifique-se de que suas fun??es personalizadas n?o usem o DOM da p?gina da Web.</span><span class="sxs-lookup"><span data-stu-id="00103-227">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="00103-228">As APIs de hospedagem suportadas para acessar a Web ser?o [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) e [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) usando GET ou POST.</span><span class="sxs-lookup"><span data-stu-id="00103-228">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="00103-229">Fun??es vol?teis (aquelas que recalculam automaticamente sempre que dados n?o relacionados s?o alterados na planilha) ainda n?o s?o suportadas.</span><span class="sxs-lookup"><span data-stu-id="00103-229">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="00103-230">A depura??o s? est? habilitada para fun??es ass?ncronas no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="00103-230">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="00103-231">A implanta??o por meio do Portal de Administra??o do Office 365 e do AppSource ainda n?o est? habilitada.</span><span class="sxs-lookup"><span data-stu-id="00103-231">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="00103-232">Fun??es personalizadas no Excel Online podem parar de funcionar durante uma sess?o ap?s um per?odo de inatividade.</span><span class="sxs-lookup"><span data-stu-id="00103-232">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="00103-233">Atualize a p?gina do navegador (F5) e insira novamente uma fun??o personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="00103-233">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="00103-234">Log de mudan?as</span><span class="sxs-lookup"><span data-stu-id="00103-234">Changelog</span></span>

- <span data-ttu-id="00103-235">**7 de novembro de 2017**: enviados exemplos e visualiza??es de fun??es personalizadas</span><span class="sxs-lookup"><span data-stu-id="00103-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="00103-236">**20 de Nov de 2017**: corre??o de bug de compatibilidade para quem usa as vers?es 8801 e posteriores</span><span class="sxs-lookup"><span data-stu-id="00103-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="00103-237">**28 de novembro de 2017**: Enviado o suporte para cancelamento em fun??es ass?ncronas (requer altera??o para fun??es de fluxo cont?nuo)</span><span class="sxs-lookup"><span data-stu-id="00103-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="00103-238">**7 de maio de 2018**: enviado o suporte para Mac, Excel Online e fun??es s?ncronas em execu??o no processo</span><span class="sxs-lookup"><span data-stu-id="00103-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
