# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="ca64e-101">Criar funções personalizadas no Excel (Visualização)</span><span class="sxs-lookup"><span data-stu-id="ca64e-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="ca64e-102">Funções personalizadas (semelhantes a funções definidas pelo usuário ou UDFs) permitem que os desenvolvedores adicionem qualquer função JavaScript no Excel usando um suplemento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="ca64e-103">Os usuários então podem acessar funções personalizadas como qualquer outra função nativa do Excel (como `=SUM()`).</span><span class="sxs-lookup"><span data-stu-id="ca64e-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="ca64e-104">Este artigo explica como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="ca64e-105">A ilustração a seguir mostra como um usuário final pode inserir uma função personalizada em uma célula.</span><span class="sxs-lookup"><span data-stu-id="ca64e-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="ca64e-106">A função que adiciona 42 a um par de números.</span><span class="sxs-lookup"><span data-stu-id="ca64e-106">Here’s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="ca64e-107">Aqui está o código para a mesma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ca64e-107">Here’s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="ca64e-108">As funções personalizadas agora estão disponíveis no Developer Preview para Windows, Mac e Excel Online.</span><span class="sxs-lookup"><span data-stu-id="ca64e-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="ca64e-109">Siga estas etapas para experimentá-las:</span><span class="sxs-lookup"><span data-stu-id="ca64e-109">Follow these steps to try them:</span></span>

1. <span data-ttu-id="ca64e-110">Instale o Office (compilação 9325 no Windows ou 13.329 no Mac) e participe do programa [Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="ca64e-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="ca64e-111">(Observe que não é suficiente apenas obter a compilação mais recente; o recurso será desabilitado em qualquer compilação até você ingressar no programa Insider)</span><span class="sxs-lookup"><span data-stu-id="ca64e-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2. <span data-ttu-id="ca64e-112">Criar um projeto de suplemento de funções personalizadas do Excel usando o [Yo Office](https://github.com/OfficeDev/generator-office)e siga as instruções no [README.md do projeto](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) para iniciar o suplemento no Excel, fazer as alterações no código e depurar.</span><span class="sxs-lookup"><span data-stu-id="ca64e-112">Create an Excel Custom Functions Add-in project using [Yo Office](https://github.com/OfficeDev/generator-office), and follow the instructions in the [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to start the add-in in Excel, make changes in the code, and debug.</span></span>
3. <span data-ttu-id="ca64e-113">Digite `=CONTOSO.ADD42(1,2)` em qualquer célula e pressione **Enter** para executar a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ca64e-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="ca64e-114">Confira a seção **Problemas conhecidos** no final deste artigo, que inclui as limitações atuais das funções personalizadas e que será atualizado com o tempo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="ca64e-115">Noções básicas</span><span class="sxs-lookup"><span data-stu-id="ca64e-115">Learn the basics</span></span>

<span data-ttu-id="ca64e-116">No repositório de exemplo clonado, você verá os seguintes arquivos:</span><span class="sxs-lookup"><span data-stu-id="ca64e-116">In the cloned sample repo, you’ll see the following files:</span></span>

- <span data-ttu-id="ca64e-117">**./src/customfunctions.js**, que contém o código de função personalizada (veja o exemplo de código simples acima para a função `ADD42`).</span><span class="sxs-lookup"><span data-stu-id="ca64e-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="ca64e-p105">**./config/customfunctions.json**, que contém o JSON de registro que informa o Excel a respeito da sua função personalizada. O registro faz com que as suas funções personalizadas apareçam na lista de funções disponíveis exibida quando um usuário digita em uma célula.</span><span class="sxs-lookup"><span data-stu-id="ca64e-p105">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function. Registration makes your custom functions appear in the list of available functions displayed when a user types in a cell.</span></span>
- <span data-ttu-id="ca64e-p106">**./index.html**, que fornece uma referência de &lt;Script&gt; para o arquivo JS. Esse arquivo não exibe a interface do usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-p106">**customfunctions.html**, which provides a &lt;Script&gt; reference to the JS file. This file does not display UI in Excel.</span></span>
- <span data-ttu-id="ca64e-122">**./manifest.xml**, que informa ao Excel a localização dos arquivos HTML, JavaScript e JSON; e também especifica um namespace para todas as funções personalizadas instaladas com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="ca64e-123"> Arquivo JSON (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="ca64e-123">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="ca64e-124">O código a seguir em customfunctions.json especifica os metadados para a mesma função `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="ca64e-125">Informações de referência detalhadas para o arquivo JSON, incluindo opções não usadas neste exemplo, estão em [Registro de Funções Personalizadas JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="ca64e-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](custom-functions-json.md).</span></span>

<span data-ttu-id="ca64e-126">Observe que, para este exemplo:</span><span class="sxs-lookup"><span data-stu-id="ca64e-126">Note that for this example:</span></span>

- <span data-ttu-id="ca64e-127">Há apenas uma função personalizada, portanto, há apenas um membro da `functions` matriz.</span><span class="sxs-lookup"><span data-stu-id="ca64e-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="ca64e-128">A propriedade `name` define o nome da função.</span><span class="sxs-lookup"><span data-stu-id="ca64e-128">The `name` property defines the function name.</span></span> <span data-ttu-id="ca64e-129">Como você viu no gif animado mostrado anteriormente, um namespace (`CONTOSO`) é anexado ao nome da função no menu de preenchimento automático do Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="ca64e-130">Esse prefixo é definido no manifesto do suplemento, descrito abaixo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="ca64e-131">O prefixo e o nome da função são separados por um ponto e, por convenção, nomes de função e prefixos são maiúsculos.</span><span class="sxs-lookup"><span data-stu-id="ca64e-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="ca64e-132">Para usar a função personalizada, o usuário digita o namespace seguido pelo nome da função (`ADD42`) em uma célula, neste caso `=CONTOSO.ADD42`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="ca64e-133">O prefixo deve ser usado como identificador para a sua empresa ou para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="ca64e-134">O `description`  aparece no menu de preenchimento automático do Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="ca64e-135">Quando o usuário solicitar ajuda para uma função, o Excel abre um painel de tarefas e exibe a página da Web encontrada no URL especificado em `helpUrl` .</span><span class="sxs-lookup"><span data-stu-id="ca64e-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="ca64e-136">A propriedade `result` especifica o tipo de informação retornada pela função para o Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="ca64e-137">A propriedade filho `type` pode `"string"`, `"number"`ou `"boolean"`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="ca64e-138">A propriedade `dimensionality` pode ser `scalar` ou `matrix` (uma matriz bidimensional de valores do `type` especificado.)</span><span class="sxs-lookup"><span data-stu-id="ca64e-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="ca64e-139">A matriz `parameters` especifica, *em ordenar*, o tipo de dado em cada parâmetro que é passado para a função.</span><span class="sxs-lookup"><span data-stu-id="ca64e-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="ca64e-140">As propriedades filho `name` e `description` são usadas no intellisense do Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="ca64e-141">As propriedades filho `type` e `dimensionality` são idênticas às propriedades filho da propriedade `result` descrita acima.</span><span class="sxs-lookup"><span data-stu-id="ca64e-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="ca64e-142">A propriedade `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="ca64e-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="ca64e-143">Há mais informações sobre essas opções posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-143">There is more information about these options later in this article.</span></span>

 ```js
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
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
> <span data-ttu-id="ca64e-144">As funções personalizadas são registradas quando um usuário executa o suplemento pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="ca64e-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="ca64e-145">Depois disso, eles estarão disponíveis, para o mesmo usuário, em todas as pastas de trabalho (não apenas naquela em que o suplemento foi executado inicialmente).</span><span class="sxs-lookup"><span data-stu-id="ca64e-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="ca64e-146">As configurações do servidor para o arquivo JSON devem ter [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que funções personalizadas funcionem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="ca64e-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-manifestxml"></a><span data-ttu-id="ca64e-147">Arquivo de manifesto (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="ca64e-147">Manifest file (manifest.xml)</span></span>


<span data-ttu-id="ca64e-148">O seguinte é um exemplo da marcação `<ExtensionPoint>` e `<Resources>` que você inclui no manifesto do suplemento para permitir que o Excel execute suas funções.</span><span class="sxs-lookup"><span data-stu-id="ca64e-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="ca64e-149">Observe o seguinte sobre essa marcação:</span><span class="sxs-lookup"><span data-stu-id="ca64e-149">Note the following about this code:</span></span>

- <span data-ttu-id="ca64e-150">O elemento `<Script>` e a identificação do recurso correspondente especificam a localização do arquivo JavaScript com suas funções.</span><span class="sxs-lookup"><span data-stu-id="ca64e-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="ca64e-151">O elemento `<Page>` e a identificação do recurso correspondente especificam a localização da página HTML do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="ca64e-152">A página HTML inclui uma marca `<Script>` que carrega o arquivo JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="ca64e-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="ca64e-153">A página HTML é uma página oculta e nunca é exibida na interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="ca64e-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="ca64e-154">O elemento `<Metadata>` e a identificação do recurso correspondente especificam a localização do arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="ca64e-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="ca64e-155">Um elemento `<Namespace>` e a identificação do recurso correspondente especificam o prefixo para todas as funções personalizadas no suplemento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


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

## <a name="initializing-custom-functions"></a><span data-ttu-id="ca64e-156">Inicialização de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ca64e-156">Initializing custom functions</span></span>

<span data-ttu-id="ca64e-157">Seu código deve inicializar o recurso de funções personalizadas antes de usá-lo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="ca64e-158">Você pode fazer isso em uma marca de &lt;Script&gt; no arquivo HTML (customfunctions.html) ou na parte superior do arquivo JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="ca64e-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="ca64e-159">Na visualização de funções personalizadas, você pode escolher entre duas sintaxes para a inicialização.</span><span class="sxs-lookup"><span data-stu-id="ca64e-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="ca64e-160">O arquivo HTML no repositório usa a seguinte sintaxe:</span><span class="sxs-lookup"><span data-stu-id="ca64e-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="ca64e-161">Você também pode usar a seguinte sintaxe:</span><span class="sxs-lookup"><span data-stu-id="ca64e-161">You can also use the following conditions:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a><span data-ttu-id="ca64e-162">Lidar com erros</span><span class="sxs-lookup"><span data-stu-id="ca64e-162">Handling errors</span></span>
<span data-ttu-id="ca64e-163">O tratamento de erros de funções personalizadas é o mesmo que [o tratamento de erros para a API do JavaScript Excel em geral](./excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="ca64e-163">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](./excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="ca64e-164">Normalmente, você usará `.catch` para lidar com erros.</span><span class="sxs-lookup"><span data-stu-id="ca64e-164">Generally, you will use `.catch` to handle errors.</span></span> <span data-ttu-id="ca64e-165">O código a seguir fornece um exemplo de `.catch`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-165">The code below gives an example of `.catch`.</span></span> 

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
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

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="ca64e-166">Funções síncronas e assíncronas</span><span class="sxs-lookup"><span data-stu-id="ca64e-166">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="ca64e-167">A função `ADD42` acima é síncrona em relação ao Excel (designada pela configuração da opção `"sync": true` no arquivo JSON).</span><span class="sxs-lookup"><span data-stu-id="ca64e-167">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="ca64e-168">As funções síncronas oferecem desempenho rápido porque são executadas no mesmo processo que o Excel e em paralelo durante o cálculo multithreaded.</span><span class="sxs-lookup"><span data-stu-id="ca64e-168">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="ca64e-169">Por outro lado, se sua função personalizada recupera dados da Web, ela deverá ser assíncrona em relação ao Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-169">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="ca64e-170">Funções assíncronas devem:</span><span class="sxs-lookup"><span data-stu-id="ca64e-170">Asynchronous functions must:</span></span>

1. <span data-ttu-id="ca64e-171">Retornar uma Promessa do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-171">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="ca64e-172">Resolver a Promise com o valor final usando a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ca64e-172">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="ca64e-173">O código a seguir mostra um exemplo de função assíncrona personalizada que recupera a temperatura de um termômetro.</span><span class="sxs-lookup"><span data-stu-id="ca64e-173">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="ca64e-174">Observe que `sendWebRequest` é uma função hipotética, não especificada aqui, que usa o XHR para chamar um serviço da Web de temperatura.</span><span class="sxs-lookup"><span data-stu-id="ca64e-174">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="ca64e-175">Funções assíncronas exibem um erro temporário `GETTING_DATA` na célula enquanto o Excel aguarda o resultado final.</span><span class="sxs-lookup"><span data-stu-id="ca64e-175">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="ca64e-176">Os usuários podem interagir normalmente com o restante da planilha enquanto aguardam o resultado.</span><span class="sxs-lookup"><span data-stu-id="ca64e-176">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="ca64e-177">Funções personalizadas são assíncronas por padrão.</span><span class="sxs-lookup"><span data-stu-id="ca64e-177">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="ca64e-178">Para designar funções como síncronas, defina a opção `"sync": true` na propriedade `options` para a função personalizada no arquivo JSON de registro.</span><span class="sxs-lookup"><span data-stu-id="ca64e-178">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="ca64e-179">Funções de fluxo</span><span class="sxs-lookup"><span data-stu-id="ca64e-179">Streamed functions</span></span>

<span data-ttu-id="ca64e-180">Uma função assíncrona pode ser de fluxo contínuo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-180">An asynchronous function can be streamed.</span></span> <span data-ttu-id="ca64e-181">Funções personalizadas de fluxo permitem que você insira dados em células repetidamente ao longo do tempo, sem precisar esperar que o Excel ou os usuários solicitem recálculos.</span><span class="sxs-lookup"><span data-stu-id="ca64e-181">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="ca64e-182">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-182">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="ca64e-183">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="ca64e-183">Note the following about this code:</span></span>

- <span data-ttu-id="ca64e-184">O Excel exibe cada novo valor automaticamente usando o retorno de chamada `setResult`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-184">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="ca64e-185">O parâmetro final, `caller`, nunca é especificado no código de registro e nunca é exibido no menu de preenchimento automático para usuários do Excel ao inserir a função.</span><span class="sxs-lookup"><span data-stu-id="ca64e-185">For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="ca64e-186">É a função de retorno de chamada `setResult` usada para passar dados da função para o Excel para atualizar o valor de uma célula.</span><span class="sxs-lookup"><span data-stu-id="ca64e-186">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="ca64e-187">Para que o Excel passe a função `setResult` no objeto `caller`, você deve declarar suporte para fluxo contínuo durante o registro da função, definindo a opção `"stream": true` na propriedade `options` para a função personalizada no arquivo JSON de registro.</span><span class="sxs-lookup"><span data-stu-id="ca64e-187">In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="ca64e-188">Cancelamento</span><span class="sxs-lookup"><span data-stu-id="ca64e-188">Cancellation</span></span>

<span data-ttu-id="ca64e-189">Você pode cancelar funções e funções assíncronas de streaming.</span><span class="sxs-lookup"><span data-stu-id="ca64e-189">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="ca64e-190">É importante cancelar as chamadas de função para reduzir o consumo de largura de banda, a memória de trabalho e a carga da CPU.</span><span class="sxs-lookup"><span data-stu-id="ca64e-190">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="ca64e-191">O Excel cancela chamadas de funções nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="ca64e-191">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="ca64e-192">O usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="ca64e-192">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="ca64e-193">É alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="ca64e-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="ca64e-194">Nesse caso, uma nova chamada de função é disparada, além do cancelamento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="ca64e-p125">O usuário aciona um recálculo manualmente. Como no caso acima, uma nova chamada de função é disparada, além do cancelamento.</span><span class="sxs-lookup"><span data-stu-id="ca64e-p125">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="ca64e-197">Você *deve* implementar um manipulador de cancelamento para todas as funções de fluxo contínuo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-197">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="ca64e-198">Funções assíncronas e que não sejam de fluxo contínuo podem ou não ser canceláveis; a decisão é sua.</span><span class="sxs-lookup"><span data-stu-id="ca64e-198">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="ca64e-199">Funções síncronas não podem ser canceladas.</span><span class="sxs-lookup"><span data-stu-id="ca64e-199">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="ca64e-200">Para tornar uma função cancelável, defina a opção `"cancelable": true` na propriedade `options` para a função personalizada no arquivo JSON de registro.</span><span class="sxs-lookup"><span data-stu-id="ca64e-200">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="ca64e-201">O código a seguir mostra o exemplo anterior com o cancelamento implementado.</span><span class="sxs-lookup"><span data-stu-id="ca64e-201">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="ca64e-202">No código, o objeto `caller` contém uma função `onCanceled` que deve ser definida para cada função personalizada cancelável.</span><span class="sxs-lookup"><span data-stu-id="ca64e-202">In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="ca64e-203">Compartilhamento e salvamento de estado</span><span class="sxs-lookup"><span data-stu-id="ca64e-203">Saving and sharing state</span></span>

<span data-ttu-id="ca64e-204">Funções personalizadas assíncronas podem salvar dados em variáveis JavaScript globais.</span><span class="sxs-lookup"><span data-stu-id="ca64e-204">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="ca64e-205">Em chamadas subsequentes, sua função personalizada pode usar valores salvos nessas variáveis.</span><span class="sxs-lookup"><span data-stu-id="ca64e-205">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="ca64e-206">O estado salvo é útil quando os usuários adicionam a mesma função personalizada a mais de uma célula, porque todas as instâncias da função podem compartilhar o estado.</span><span class="sxs-lookup"><span data-stu-id="ca64e-206">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="ca64e-207">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="ca64e-207">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="ca64e-208">O código a seguir mostra uma implementação da função anterior de fluxo contínuo de temperatura que salva o estado de forma global.</span><span class="sxs-lookup"><span data-stu-id="ca64e-208">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="ca64e-209">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="ca64e-209">Note the following about this code:</span></span>

- <span data-ttu-id="ca64e-210">`refreshTemperature` é uma função de fluxo que lê a temperatura de um determinado termômetro a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="ca64e-210">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="ca64e-211">Novas temperaturas são salvas na variável `savedTemperatures`, mas o valor da célula não é atualizado diretamente.</span><span class="sxs-lookup"><span data-stu-id="ca64e-211">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="ca64e-212">Não deve ser chamada diretamente de uma célula da planilha, *por isso não está registrada no arquivo JSON*.</span><span class="sxs-lookup"><span data-stu-id="ca64e-212">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="ca64e-213">`streamTemperature` atualiza os valores de temperatura exibidos na célula a cada segundo e usa variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="ca64e-213">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="ca64e-214">Deve ser registrada no arquivo JSON e nomeada com todas as letras maiúsculas, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-214">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="ca64e-215">Os usuários podem chamar `streamTemperature` de várias células na interface de usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-215">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="ca64e-216">Cada chamada lê dados da mesma variável `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-216">Each call reads data from the same `savedTemperatures` variable.</span></span>

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
> <span data-ttu-id="ca64e-217">Funções síncronas (designadas pela configuração da opção `"sync": true` no arquivo JSON) não podem compartilhar estado porque o Excel faz o paralelismo delas durante o cálculo multithreaded.</span><span class="sxs-lookup"><span data-stu-id="ca64e-217">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="ca64e-218">Somente funções assíncronas podem compartilhar estado porque as funções síncronas de um suplemento compartilham o mesmo contexto JavaScript em cada sessão.</span><span class="sxs-lookup"><span data-stu-id="ca64e-218">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="ca64e-219">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="ca64e-219">Working with ranges of data</span></span>

<span data-ttu-id="ca64e-220">Sua função personalizada pode levar a um intervalo de dados como um parâmetro ou você pode retornar um intervalo de dados de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ca64e-220">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="ca64e-221">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-221">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="ca64e-222">A função a seguir usa o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-222">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="ca64e-223">Note que no registro JSON para esta função, você definiria a propriedade `type` do parâmetro para `matrix`.</span><span class="sxs-lookup"><span data-stu-id="ca64e-223">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

<span data-ttu-id="ca64e-224">Como pode ver, os intervalos são tratados em JavaScript como matrizes de matrizes de linhas (como uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="ca64e-224">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="ca64e-225">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="ca64e-225">Known issues</span></span>

- <span data-ttu-id="ca64e-226">As descrições de URLs e parâmetros de Ajuda ainda não são usadas pelo Excel.</span><span class="sxs-lookup"><span data-stu-id="ca64e-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="ca64e-227">Funções personalizadas não estão atualmente disponíveis no Excel para clientes móveis.</span><span class="sxs-lookup"><span data-stu-id="ca64e-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="ca64e-228">Atualmente, os suplementos dependem de um processo de navegador oculto para executar funções personalizadas assíncronas.</span><span class="sxs-lookup"><span data-stu-id="ca64e-228">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="ca64e-229">No futuro, o JavaScript será executado diretamente em algumas plataformas para garantir que as funções personalizadas sejam mais rápidas e usem menos memória.</span><span class="sxs-lookup"><span data-stu-id="ca64e-229">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="ca64e-230">Além disso, a página HTML referenciada pelo elemento `<Page>` no manifesto não será necessária para a maioria das plataformas, já que o Excel executa o JavaScript diretamente.</span><span class="sxs-lookup"><span data-stu-id="ca64e-230">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="ca64e-231">Para se preparar para essa alteração, certifique-se de que suas funções personalizadas não usem o DOM da página da Web.</span><span class="sxs-lookup"><span data-stu-id="ca64e-231">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="ca64e-232">As APIs de hospedagem suportadas para acessar a Web serão [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) e [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) usando GET ou POST.</span><span class="sxs-lookup"><span data-stu-id="ca64e-232">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="ca64e-233">Funções voláteis (aquelas que recalculam automaticamente sempre que dados não relacionados são alterados na planilha) ainda não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="ca64e-233">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="ca64e-234">A depuração só está habilitada para funções assíncronas no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="ca64e-234">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="ca64e-235">A implantação por meio do Portal de Administração do Office 365 e do AppSource ainda não está habilitada.</span><span class="sxs-lookup"><span data-stu-id="ca64e-235">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="ca64e-236">Funções personalizadas no Excel Online podem parar de funcionar durante uma sessão após um período de inatividade.</span><span class="sxs-lookup"><span data-stu-id="ca64e-236">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="ca64e-237">Atualize a página do navegador (F5) e insira novamente uma função personalizada para restaurar o recurso.</span><span class="sxs-lookup"><span data-stu-id="ca64e-237">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="ca64e-238">Log de mudanças</span><span class="sxs-lookup"><span data-stu-id="ca64e-238">Changelog</span></span>

- <span data-ttu-id="ca64e-239">**7 de novembro de 2017**: enviados exemplos e visualizações de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ca64e-239">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="ca64e-240">**20 de Nov de 2017**: correção de bug de compatibilidade para quem usa as versões 8801 e posteriores</span><span class="sxs-lookup"><span data-stu-id="ca64e-240">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="ca64e-241">**28 de novembro de 2017**: enviado o suporte para cancelamento em funções assíncronas (requer a alteração de funções de streaming)</span><span class="sxs-lookup"><span data-stu-id="ca64e-241">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="ca64e-242">**7 de maio de 2018**: enviado o suporte para Mac, Excel Online e funções síncronas em execução no processo</span><span class="sxs-lookup"><span data-stu-id="ca64e-242">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
