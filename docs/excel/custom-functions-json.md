# <a name="custom-function-metadata"></a><span data-ttu-id="2d18c-101">Metadados da função personalizada</span><span class="sxs-lookup"><span data-stu-id="2d18c-101">Custom function metadata</span></span>

<span data-ttu-id="2d18c-102">Ao incluir [funções personalizadas](custom-functions-overview.md) em um suplemento do Excel, você deve hospedar um arquivo JSON que contenha metadados sobre as funções (além de hospedar um arquivo JavaScript com as funções e um arquivo HTML sem interface do usuário para servir como pai do arquivo JavaScript).</span><span class="sxs-lookup"><span data-stu-id="2d18c-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="2d18c-103">Este artigo descreve o formato do arquivo JSON com exemplos.</span><span class="sxs-lookup"><span data-stu-id="2d18c-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="2d18c-104">Há um arquivo JSON de amostra completo disponível [aqui](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="2d18c-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="2d18c-105">Matriz de funções</span><span class="sxs-lookup"><span data-stu-id="2d18c-105">Functions array</span></span>

<span data-ttu-id="2d18c-106">Os metadados são um objeto JSON que contém uma única propriedade `functions` cujo valor é uma matriz de objetos.</span><span class="sxs-lookup"><span data-stu-id="2d18c-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="2d18c-107">Cada um desses objetos representa uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="2d18c-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="2d18c-108">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="2d18c-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="2d18c-109">Propriedade</span><span class="sxs-lookup"><span data-stu-id="2d18c-109">Property</span></span>  |  <span data-ttu-id="2d18c-110">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="2d18c-110">Data Type</span></span>  |  <span data-ttu-id="2d18c-111">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="2d18c-111">Required?</span></span>  |  <span data-ttu-id="2d18c-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="2d18c-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="2d18c-113">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-113">string</span></span>  |  <span data-ttu-id="2d18c-114">Não</span><span class="sxs-lookup"><span data-stu-id="2d18c-114">No</span></span>  |  <span data-ttu-id="2d18c-115">Uma descrição da função que aparece na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="2d18c-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="2d18c-116">Por exemplo, “Converte um valor Celsius em Fahrenheit”.</span><span class="sxs-lookup"><span data-stu-id="2d18c-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="2d18c-117">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-117">string</span></span>  |   <span data-ttu-id="2d18c-118">Não</span><span class="sxs-lookup"><span data-stu-id="2d18c-118">No</span></span>  |  <span data-ttu-id="2d18c-119">A URL na qual seus usuários podem obter ajuda sobre a função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="2d18c-120">(Ela é exibida em um painel de tarefas.) Por exemplo, “http://contoso.com/help/convertcelsiustofahrenheit.html”</span><span class="sxs-lookup"><span data-stu-id="2d18c-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="2d18c-121">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-121">string</span></span>  |  <span data-ttu-id="2d18c-122">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-122">Yes</span></span>  |  <span data-ttu-id="2d18c-123">O nome da função como será exibido (precedido de um namespace) na interface do usuário do Excel quando um usuário estiver selecionando uma função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="2d18c-124">Deve ser o mesmo que o nome da função nos locais em que estiver definido no JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2d18c-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="2d18c-125">objeto</span><span class="sxs-lookup"><span data-stu-id="2d18c-125">object</span></span>  |  <span data-ttu-id="2d18c-126">Não</span><span class="sxs-lookup"><span data-stu-id="2d18c-126">No</span></span>  |  <span data-ttu-id="2d18c-127">Configurar como o Excel processa a função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="2d18c-128">Consulte [objeto de opções](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="2d18c-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="2d18c-129">matriz</span><span class="sxs-lookup"><span data-stu-id="2d18c-129">array</span></span>  |  <span data-ttu-id="2d18c-130">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-130">Yes</span></span>  |  <span data-ttu-id="2d18c-131">Metadados sobre os parâmetros para a função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="2d18c-132">Consulte [matriz de parâmetros](#parameters-array) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="2d18c-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="2d18c-133">objeto</span><span class="sxs-lookup"><span data-stu-id="2d18c-133">object</span></span>  |  <span data-ttu-id="2d18c-134">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-134">Yes</span></span>  |  <span data-ttu-id="2d18c-135">Metadados sobre o valor retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="2d18c-136">Consulte [objeto de resultado](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="2d18c-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="2d18c-137">Objeto Options</span><span class="sxs-lookup"><span data-stu-id="2d18c-137">Options object</span></span>

<span data-ttu-id="2d18c-138">O objeto `options` configura como o Excel processa a função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="2d18c-139">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="2d18c-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="2d18c-140">Propriedade</span><span class="sxs-lookup"><span data-stu-id="2d18c-140">Property</span></span>  |  <span data-ttu-id="2d18c-141">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="2d18c-141">Data Type</span></span>  |  <span data-ttu-id="2d18c-142">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="2d18c-142">Required?</span></span>  |  <span data-ttu-id="2d18c-143">Descrição</span><span class="sxs-lookup"><span data-stu-id="2d18c-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="2d18c-144">booleano</span><span class="sxs-lookup"><span data-stu-id="2d18c-144">boolean</span></span>  |  <span data-ttu-id="2d18c-145">Não, o padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="2d18c-145">No, default is `false`.</span></span>  |  <span data-ttu-id="2d18c-p110">Se `true`, o Excel chama o manipulador de `onCanceled` sempre que o usuário realizar uma ação que tem o efeito de cancelar a função; por exemplo, disparando manualmente o recálculo ou editando uma célula referenciada pela função. Se você usar essa opção, o Excel chamará a função JavaScript com o parâmetro adicional `caller`. (***Não*** registre esse parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="2d18c-p110">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. Note,  and  cannot both be .</span></span>|
|  `stream`  |  <span data-ttu-id="2d18c-150">booleano</span><span class="sxs-lookup"><span data-stu-id="2d18c-150">boolean</span></span>  |  <span data-ttu-id="2d18c-151">Não, o padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="2d18c-151">No, default is `false`.</span></span>  |  <span data-ttu-id="2d18c-152">Se for `true`, a função pode ser repetidamente a saída da célula, mesmo quando invocada apenas uma vez.</span><span class="sxs-lookup"><span data-stu-id="2d18c-152">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="2d18c-153">Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação.</span><span class="sxs-lookup"><span data-stu-id="2d18c-153">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="2d18c-154">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="2d18c-154">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="2d18c-155">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="2d18c-155">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="2d18c-156">A função não deve ter a instrução `return`.</span><span class="sxs-lookup"><span data-stu-id="2d18c-156">The function should have no `return` statement.</span></span> <span data-ttu-id="2d18c-157">Em vez disso, o valor do resultado é passado como o argumento do método de retorno de chamada `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="2d18c-157">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span>|

## <a name="parameters-array"></a><span data-ttu-id="2d18c-158">Matriz de parâmetros</span><span class="sxs-lookup"><span data-stu-id="2d18c-158">Parameters array</span></span>

<span data-ttu-id="2d18c-159">A propriedade `parameters` é uma matriz de objetos.</span><span class="sxs-lookup"><span data-stu-id="2d18c-159">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="2d18c-160">Cada um desses objetos representa um parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2d18c-160">Each of these objects represents a parameter.</span></span> <span data-ttu-id="2d18c-161">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="2d18c-161">The following table contains its properties:</span></span>

|  <span data-ttu-id="2d18c-162">Propriedade</span><span class="sxs-lookup"><span data-stu-id="2d18c-162">Property</span></span>  |  <span data-ttu-id="2d18c-163">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="2d18c-163">Data Type</span></span>  |  <span data-ttu-id="2d18c-164">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="2d18c-164">Required?</span></span>  |  <span data-ttu-id="2d18c-165">Descrição</span><span class="sxs-lookup"><span data-stu-id="2d18c-165">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="2d18c-166">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-166">string</span></span>  |  <span data-ttu-id="2d18c-167">Não</span><span class="sxs-lookup"><span data-stu-id="2d18c-167">No</span></span> |  <span data-ttu-id="2d18c-168">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2d18c-168">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="2d18c-169">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-169">string</span></span>  |  <span data-ttu-id="2d18c-170">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-170">Yes</span></span>  |  <span data-ttu-id="2d18c-171">Deve ser “escalar”, ou seja, um valor que não é matriz; ou “matriz”, ou seja, uma matriz de matrizes de linhas.</span><span class="sxs-lookup"><span data-stu-id="2d18c-171">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="2d18c-172">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-172">string</span></span>  |  <span data-ttu-id="2d18c-173">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-173">Yes</span></span>  |  <span data-ttu-id="2d18c-174">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2d18c-174">The name of the parameter.</span></span> <span data-ttu-id="2d18c-175">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="2d18c-175">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="2d18c-176">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-176">string</span></span>  |  <span data-ttu-id="2d18c-177">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-177">Yes</span></span>  |  <span data-ttu-id="2d18c-178">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2d18c-178">The data type of the parameter.</span></span> <span data-ttu-id="2d18c-179">Deve ser “booleano”, “número” ou “sequência de caracteres”.</span><span class="sxs-lookup"><span data-stu-id="2d18c-179">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="2d18c-180">Objeto de resultado</span><span class="sxs-lookup"><span data-stu-id="2d18c-180">Result object</span></span>

<span data-ttu-id="2d18c-181">A propriedade `results` fornece metadados sobre o valor retornado da função.</span><span class="sxs-lookup"><span data-stu-id="2d18c-181">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="2d18c-182">A tabela a seguir contém suas propriedades:</span><span class="sxs-lookup"><span data-stu-id="2d18c-182">The following table contains its properties:</span></span>

|  <span data-ttu-id="2d18c-183">Propriedade</span><span class="sxs-lookup"><span data-stu-id="2d18c-183">Property</span></span>  |  <span data-ttu-id="2d18c-184">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="2d18c-184">Data Type</span></span>  |  <span data-ttu-id="2d18c-185">Obrigatório?</span><span class="sxs-lookup"><span data-stu-id="2d18c-185">Required?</span></span>  |  <span data-ttu-id="2d18c-186">Descrição</span><span class="sxs-lookup"><span data-stu-id="2d18c-186">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="2d18c-187">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-187">string</span></span>  |  <span data-ttu-id="2d18c-188">Não</span><span class="sxs-lookup"><span data-stu-id="2d18c-188">No</span></span>  |  <span data-ttu-id="2d18c-189">Deve ser “escalar”, ou seja, um valor que não é matriz; ou “matriz”, ou seja, uma matriz de matrizes de linhas.</span><span class="sxs-lookup"><span data-stu-id="2d18c-189">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="2d18c-190">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="2d18c-190">string</span></span>  |  <span data-ttu-id="2d18c-191">Sim</span><span class="sxs-lookup"><span data-stu-id="2d18c-191">Yes</span></span>  |  <span data-ttu-id="2d18c-192">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="2d18c-192">The data type of the parameter.</span></span> <span data-ttu-id="2d18c-193">Deve ser “booleano”, “número” ou “sequência de caracteres”.</span><span class="sxs-lookup"><span data-stu-id="2d18c-193">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="2d18c-194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2d18c-194">Example</span></span>

<span data-ttu-id="2d18c-195">O código JSON a seguir é um exemplo de um arquivo de metadados para funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2d18c-195">The following JSON code is an example of a metadata file for custom functions.</span></span>

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
            ]
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
            ]
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
            ]
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
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
            ]
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="2d18c-196">Confira também</span><span class="sxs-lookup"><span data-stu-id="2d18c-196">See also</span></span>
[<span data-ttu-id="2d18c-197">Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="2d18c-197">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="2d18c-198">Diretrizes e exemplos de fórmulas de matriz</span><span class="sxs-lookup"><span data-stu-id="2d18c-198">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
