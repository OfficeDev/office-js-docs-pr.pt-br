# <a name="custom-function-metadata"></a><span data-ttu-id="c61ab-101">Metadados da função personalizada</span><span class="sxs-lookup"><span data-stu-id="c61ab-101">Custom function metadata</span></span>

<span data-ttu-id="c61ab-102">Ao incluir [funções personalizadas](custom-functions-overview.md) em um suplemento do Excel, você deve hospedar um arquivo JSON que contenha metadados sobre as funções (além de hospedar um arquivo JavaScript com as funções e um arquivo HTML sem interface do usuário para servir como pai do arquivo JavaScript).</span><span class="sxs-lookup"><span data-stu-id="c61ab-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="c61ab-103">Este artigo descreve o formato do arquivo JSON com exemplos.</span><span class="sxs-lookup"><span data-stu-id="c61ab-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="c61ab-104">Há um arquivo JSON de amostra completo disponível [aqui](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="c61ab-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="c61ab-105">Matriz de funções</span><span class="sxs-lookup"><span data-stu-id="c61ab-105">Functions array</span></span>

<span data-ttu-id="c61ab-106">Os metadados são um objeto JSON que contém uma única propriedade `functions` cujo valor é uma matriz de objetos.</span><span class="sxs-lookup"><span data-stu-id="c61ab-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="c61ab-107">Cada um desses objetos representa uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="c61ab-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="c61ab-108">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="c61ab-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="c61ab-109">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c61ab-109">Property</span></span>  |  <span data-ttu-id="c61ab-110">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="c61ab-110">Data Type</span></span>  |  <span data-ttu-id="c61ab-111">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="c61ab-111">Required?</span></span>  |  <span data-ttu-id="c61ab-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="c61ab-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="c61ab-113">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-113">string</span></span>  |  <span data-ttu-id="c61ab-114">Não</span><span class="sxs-lookup"><span data-stu-id="c61ab-114">No</span></span>  |  <span data-ttu-id="c61ab-115">Uma descrição da função que aparece na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="c61ab-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="c61ab-116">Por exemplo, “Converte um valor Celsius em Fahrenheit”.</span><span class="sxs-lookup"><span data-stu-id="c61ab-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="c61ab-117">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-117">string</span></span>  |   <span data-ttu-id="c61ab-118">Não</span><span class="sxs-lookup"><span data-stu-id="c61ab-118">No</span></span>  |  <span data-ttu-id="c61ab-119">A URL na qual seus usuários podem obter ajuda sobre a função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="c61ab-120">(Ela é exibida em um painel de tarefas.) Por exemplo, “http://contoso.com/help/convertcelsiustofahrenheit.html”</span><span class="sxs-lookup"><span data-stu-id="c61ab-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="c61ab-121">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-121">string</span></span>  |  <span data-ttu-id="c61ab-122">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-122">Yes</span></span>  |  <span data-ttu-id="c61ab-123">O nome da função como será exibido (precedido de um namespace) na interface do usuário do Excel quando um usuário estiver selecionando uma função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="c61ab-124">Deve ser o mesmo que o nome da função nos locais em que estiver definido no JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c61ab-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="c61ab-125">objeto</span><span class="sxs-lookup"><span data-stu-id="c61ab-125">object</span></span>  |  <span data-ttu-id="c61ab-126">Não</span><span class="sxs-lookup"><span data-stu-id="c61ab-126">No</span></span>  |  <span data-ttu-id="c61ab-127">Configurar como o Excel processa a função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="c61ab-128">Consulte [objeto de opções](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="c61ab-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="c61ab-129">matriz</span><span class="sxs-lookup"><span data-stu-id="c61ab-129">array</span></span>  |  <span data-ttu-id="c61ab-130">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-130">Yes</span></span>  |  <span data-ttu-id="c61ab-131">Metadados sobre os parâmetros para a função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="c61ab-132">Consulte [matriz de parâmetros](#parameters-array) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="c61ab-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="c61ab-133">objeto</span><span class="sxs-lookup"><span data-stu-id="c61ab-133">object</span></span>  |  <span data-ttu-id="c61ab-134">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-134">Yes</span></span>  |  <span data-ttu-id="c61ab-135">Metadados sobre o valor retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="c61ab-136">Consulte [objeto de resultado](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="c61ab-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="c61ab-137">Objeto de opções</span><span class="sxs-lookup"><span data-stu-id="c61ab-137">Options object</span></span>

<span data-ttu-id="c61ab-138">O objeto `options` configura como o Excel processa a função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="c61ab-139">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="c61ab-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="c61ab-140">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c61ab-140">Property</span></span>  |  <span data-ttu-id="c61ab-141">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="c61ab-141">Data Type</span></span>  |  <span data-ttu-id="c61ab-142">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="c61ab-142">Required?</span></span>  |  <span data-ttu-id="c61ab-143">Descrição</span><span class="sxs-lookup"><span data-stu-id="c61ab-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="c61ab-144">booleano</span><span class="sxs-lookup"><span data-stu-id="c61ab-144">boolean</span></span>  |  <span data-ttu-id="c61ab-145">Não, o padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-145">No, default is `false`.</span></span>  |  <span data-ttu-id="c61ab-146">Se for `true`, o Excel chama o manipulador `onCanceled` sempre que o usuário executar uma ação que tenha o efeito de cancelar a função; por exemplo, ao disparar manualmente o recálculo ou ao editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-146">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="c61ab-147">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="c61ab-147">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="c61ab-148">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="c61ab-148">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="c61ab-149">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-149">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="c61ab-150">Observe que `cancelable` e `sync` não podem ambos ser `true`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-150">Note, `cancelable` and `sync` cannot both be `true`.</span></span>  |
|  `stream`  |  <span data-ttu-id="c61ab-151">booleano</span><span class="sxs-lookup"><span data-stu-id="c61ab-151">boolean</span></span>  |  <span data-ttu-id="c61ab-152">Não, o padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-152">No, default is `false`.</span></span>  |  <span data-ttu-id="c61ab-153">Se for `true`, a função pode ser repetidamente a saída da célula, mesmo quando invocada apenas uma vez.</span><span class="sxs-lookup"><span data-stu-id="c61ab-153">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="c61ab-154">Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação.</span><span class="sxs-lookup"><span data-stu-id="c61ab-154">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="c61ab-155">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="c61ab-155">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="c61ab-156">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="c61ab-156">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="c61ab-157">A função não deve ter a instrução `return`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-157">The function should have no `return` statement.</span></span> <span data-ttu-id="c61ab-158">Em vez disso, o valor do resultado é passado como o argumento do método de retorno de chamada `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-158">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="c61ab-159">Observe que `stream` e `sync` não podem ambos ser `true`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-159">Note, `stream` and `sync` may not both be `true`.</span></span>|
|  `sync`  |  <span data-ttu-id="c61ab-160">booleano</span><span class="sxs-lookup"><span data-stu-id="c61ab-160">boolean</span></span>  |  <span data-ttu-id="c61ab-161">Não, o padrão é `false`</span><span class="sxs-lookup"><span data-stu-id="c61ab-161">No, default is `false`</span></span>  |  <span data-ttu-id="c61ab-162">Se for `true`, a função é executada de forma síncrona e deve retornar um valor.</span><span class="sxs-lookup"><span data-stu-id="c61ab-162">If `true`, the function runs synchronously and it must return a value.</span></span> <span data-ttu-id="c61ab-163">Se for `false`, a função é executada de forma assíncrona e deve retornar o objeto `OfficeExtension.Promise`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-163">If `false`, the function runs asynchronously and it must return a `OfficeExtension.Promise` object.</span></span> <span data-ttu-id="c61ab-164">Observe que talvez `sync` não seja `true` se `cancelable` ou `stream` for `true`.</span><span class="sxs-lookup"><span data-stu-id="c61ab-164">Note, `sync`  may not be `true` if either `cancelable` or `stream` are `true`.</span></span>  |

## <a name="parameters-array"></a><span data-ttu-id="c61ab-165">Matriz de parâmetros</span><span class="sxs-lookup"><span data-stu-id="c61ab-165">Parameters array</span></span>

<span data-ttu-id="c61ab-166">A propriedade `parameters` é uma matriz de objetos.</span><span class="sxs-lookup"><span data-stu-id="c61ab-166">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="c61ab-167">Cada um desses objetos representa um parâmetro.</span><span class="sxs-lookup"><span data-stu-id="c61ab-167">Each of these objects represents a parameter.</span></span> <span data-ttu-id="c61ab-168">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="c61ab-168">The following table contains its properties:</span></span>

|  <span data-ttu-id="c61ab-169">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c61ab-169">Property</span></span>  |  <span data-ttu-id="c61ab-170">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="c61ab-170">Data Type</span></span>  |  <span data-ttu-id="c61ab-171">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="c61ab-171">Required?</span></span>  |  <span data-ttu-id="c61ab-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="c61ab-172">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="c61ab-173">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-173">string</span></span>  |  <span data-ttu-id="c61ab-174">Não</span><span class="sxs-lookup"><span data-stu-id="c61ab-174">No</span></span> |  <span data-ttu-id="c61ab-175">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="c61ab-175">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="c61ab-176">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-176">string</span></span>  |  <span data-ttu-id="c61ab-177">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-177">Yes</span></span>  |  <span data-ttu-id="c61ab-178">Deve ser “escalar”, ou seja, um valor que não é matriz; ou “matriz”, ou seja, uma matriz de matrizes de linhas.</span><span class="sxs-lookup"><span data-stu-id="c61ab-178">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="c61ab-179">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-179">string</span></span>  |  <span data-ttu-id="c61ab-180">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-180">Yes</span></span>  |  <span data-ttu-id="c61ab-181">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="c61ab-181">The name of the parameter.</span></span> <span data-ttu-id="c61ab-182">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="c61ab-182">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="c61ab-183">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-183">string</span></span>  |  <span data-ttu-id="c61ab-184">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-184">Yes</span></span>  |  <span data-ttu-id="c61ab-185">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="c61ab-185">The data type of the parameter.</span></span> <span data-ttu-id="c61ab-186">Deve ser “booleano”, “número” ou “sequência de caracteres”.</span><span class="sxs-lookup"><span data-stu-id="c61ab-186">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="c61ab-187">Objeto de resultado</span><span class="sxs-lookup"><span data-stu-id="c61ab-187">Result object</span></span>

<span data-ttu-id="c61ab-188">A propriedade `results` fornece metadados sobre o valor retornado da função.</span><span class="sxs-lookup"><span data-stu-id="c61ab-188">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="c61ab-189">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="c61ab-189">The following table contains its properties:</span></span>

|  <span data-ttu-id="c61ab-190">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c61ab-190">Property</span></span>  |  <span data-ttu-id="c61ab-191">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="c61ab-191">Data Type</span></span>  |  <span data-ttu-id="c61ab-192">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="c61ab-192">Required?</span></span>  |  <span data-ttu-id="c61ab-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="c61ab-193">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="c61ab-194">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-194">string</span></span>  |  <span data-ttu-id="c61ab-195">Não</span><span class="sxs-lookup"><span data-stu-id="c61ab-195">No</span></span>  |  <span data-ttu-id="c61ab-196">Deve ser “escalar”, ou seja, um valor que não é matriz; ou “matriz”, ou seja, uma matriz de matrizes de linhas.</span><span class="sxs-lookup"><span data-stu-id="c61ab-196">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="c61ab-197">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="c61ab-197">string</span></span>  |  <span data-ttu-id="c61ab-198">Sim</span><span class="sxs-lookup"><span data-stu-id="c61ab-198">Yes</span></span>  |  <span data-ttu-id="c61ab-199">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="c61ab-199">The data type of the parameter.</span></span> <span data-ttu-id="c61ab-200">Deve ser “booleano”, “número” ou “sequência de caracteres”.</span><span class="sxs-lookup"><span data-stu-id="c61ab-200">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="c61ab-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c61ab-201">Example</span></span>

<span data-ttu-id="c61ab-202">O código JSON a seguir é um exemplo de um arquivo de metadados para funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="c61ab-202">The following JSON code is an example of a metadata file for custom functions.</span></span>

```json
{
    "functions": [
        {
            "name": "ADD42", 
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "ADD42ASYNC", 
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": false
            }
        },
        {
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": [],
            "options": {
                "sync": true
            }
        },
        {
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "increment",
                    "description": "the number to be added each time",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": false,
                "stream": true,
                "cancelable": true
            }
        },
        {
            "name": "SECONDHIGHEST", 
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "range",
                    "description": "the input range",
                    "type": "number",
                    "dimensionality": "matrix"
                }
            ],
            "options": {
                "sync": true
            }
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="c61ab-203">Confira também</span><span class="sxs-lookup"><span data-stu-id="c61ab-203">See also</span></span>
[<span data-ttu-id="c61ab-204">Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c61ab-204">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="c61ab-205">Diretrizes e exemplos de fórmulas de matriz</span><span class="sxs-lookup"><span data-stu-id="c61ab-205">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
