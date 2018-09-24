---
ms.date: 09/20/2018
description: Defina metadados para funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062141"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="1071b-103">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1071b-103">Custom functions metadata</span></span>

<span data-ttu-id="1071b-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o seu projeto de suplemento deve incluir um arquivo de metadados JSON que fornece as informações que o Excel precisa para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="1071b-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="1071b-105">Este artigo descreve o formato do arquivo JSON de metadados.</span><span class="sxs-lookup"><span data-stu-id="1071b-105">This article describes the format of the JSON file with examples.</span></span>

> [!NOTE]
> <span data-ttu-id="1071b-106">Para obter informações sobre os outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criação de funções personalizadas no Excel](custom-functions-overview.md#learn-the-basics).</span><span class="sxs-lookup"><span data-stu-id="1071b-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md#learn-the-basics).</span></span>

## <a name="example-metadata"></a><span data-ttu-id="1071b-107">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="1071b-107">Example metadata</span></span>

<span data-ttu-id="1071b-108">O exemplo a seguir mostra o conteúdo de um arquivo JSON de metadados para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1071b-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="1071b-109">As seções seguintes a esse exemplo fornecem informações detalhadas sobre as propriedades individuais deste exemplo JSON.</span><span class="sxs-lookup"><span data-stu-id="1071b-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
    "functions": [
        {
            "id": "ADD42",
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
            "id": "ADD42ASYNC",
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
            "id": "ISEVEN",
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
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
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
            "id": "SECONDHIGHEST",
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

> [!NOTE]
> <span data-ttu-id="1071b-110">Um exemplo completo de arquivo JSON está disponível no [repositório GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="1071b-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="1071b-111">functions</span><span class="sxs-lookup"><span data-stu-id="1071b-111">functions</span></span> 

<span data-ttu-id="1071b-112">A propriedade `functions` é uma matriz de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1071b-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="1071b-113">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="1071b-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="1071b-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1071b-114">Property</span></span>  |  <span data-ttu-id="1071b-115">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1071b-115">Data type</span></span>  |  <span data-ttu-id="1071b-116">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1071b-116">Required</span></span>  |  <span data-ttu-id="1071b-117">Descrição</span><span class="sxs-lookup"><span data-stu-id="1071b-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="1071b-118">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-118">string</span></span>  |  <span data-ttu-id="1071b-119">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-119">No</span></span>  |  <span data-ttu-id="1071b-120">Uma descrição da função que aparece na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="1071b-120">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="1071b-121">Por exemplo, **Converte um valor Celsius em Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="1071b-121">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="1071b-122">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-122">string</span></span>  |   <span data-ttu-id="1071b-123">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-123">No</span></span>  |  <span data-ttu-id="1071b-124">A URL onde os usuários podem obter informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="1071b-124">URL where your users can get help about the function.</span></span> <span data-ttu-id="1071b-125">(É exibida em um painel de tarefas.) Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="1071b-125">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span> |
| `id`     | <span data-ttu-id="1071b-126">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-126">string</span></span> | <span data-ttu-id="1071b-127">Sim</span><span class="sxs-lookup"><span data-stu-id="1071b-127">Yes</span></span> | <span data-ttu-id="1071b-128">Um ID exclusivo para a função.</span><span class="sxs-lookup"><span data-stu-id="1071b-128">A unique ID for the group.</span></span> <span data-ttu-id="1071b-129">Esse ID não deve ser alterado depois de ser definido.</span><span class="sxs-lookup"><span data-stu-id="1071b-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="1071b-130">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-130">string</span></span>  |  <span data-ttu-id="1071b-131">Sim</span><span class="sxs-lookup"><span data-stu-id="1071b-131">Yes</span></span>  |  <span data-ttu-id="1071b-132">O nome da função como será exibido (precedido de um namespace) na interface do usuário do Excel quando um usuário estiver selecionando uma função.</span><span class="sxs-lookup"><span data-stu-id="1071b-132">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="1071b-133">Não precisa ser igual ao nome da função nos locais em que estiver definido no JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1071b-133">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="1071b-134">object</span><span class="sxs-lookup"><span data-stu-id="1071b-134">object</span></span>  |  <span data-ttu-id="1071b-135">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-135">No</span></span>  |  <span data-ttu-id="1071b-136">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="1071b-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1071b-137">Confira [objeto options](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="1071b-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="1071b-138">matriz</span><span class="sxs-lookup"><span data-stu-id="1071b-138">array</span></span>  |  <span data-ttu-id="1071b-139">Sim</span><span class="sxs-lookup"><span data-stu-id="1071b-139">Yes</span></span>  |  <span data-ttu-id="1071b-140">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="1071b-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="1071b-141">Confira [matriz de parâmetros](#parameters-array) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="1071b-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="1071b-142">object</span><span class="sxs-lookup"><span data-stu-id="1071b-142">object</span></span>  |  <span data-ttu-id="1071b-143">Sim</span><span class="sxs-lookup"><span data-stu-id="1071b-143">Yes</span></span>  |  <span data-ttu-id="1071b-144">Objeto que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="1071b-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1071b-145">Confira [objeto result](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="1071b-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="1071b-146">options</span><span class="sxs-lookup"><span data-stu-id="1071b-146">options</span></span>

<span data-ttu-id="1071b-147">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="1071b-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1071b-148">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="1071b-148">The following table describes the properties of the project.</span></span>

|  <span data-ttu-id="1071b-149">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1071b-149">Property</span></span>  |  <span data-ttu-id="1071b-150">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1071b-150">Data type</span></span>  |  <span data-ttu-id="1071b-151">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1071b-151">Required</span></span>  |  <span data-ttu-id="1071b-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="1071b-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="1071b-153">booleano</span><span class="sxs-lookup"><span data-stu-id="1071b-153">boolean</span></span>  |  <span data-ttu-id="1071b-154">Não, o padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1071b-154">No, default is `false`.</span></span>  |  <span data-ttu-id="1071b-155">Se for `true`, o Excel chama o manipulador `onCanceled` sempre que o usuário executar uma ação que tenha o efeito de cancelar a função; por exemplo, ao disparar manualmente o recálculo ou ao editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="1071b-155">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="1071b-156">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="1071b-156">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="1071b-157">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="1071b-157">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="1071b-158">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="1071b-158">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="1071b-159">Para obter mais informações, consulte [Cancelamento de uma função](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="1071b-159">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="1071b-160">booleano</span><span class="sxs-lookup"><span data-stu-id="1071b-160">boolean</span></span>  |  <span data-ttu-id="1071b-161">Não, o padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1071b-161">No, default is `false`.</span></span>  |  <span data-ttu-id="1071b-162">Se for `true`, a função pode ser repetidamente a saída da célula, mesmo quando invocada apenas uma vez.</span><span class="sxs-lookup"><span data-stu-id="1071b-162">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="1071b-163">Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação.</span><span class="sxs-lookup"><span data-stu-id="1071b-163">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="1071b-164">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="1071b-164">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="1071b-165">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="1071b-165">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="1071b-166">A função não deve ter a instrução `return`.</span><span class="sxs-lookup"><span data-stu-id="1071b-166">The function should have no `return` statement.</span></span> <span data-ttu-id="1071b-167">Em vez disso, o valor do resultado é passado como argumento do método de retorno de chamada `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="1071b-167">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="1071b-168">Para obter mais informações, consulte [Funções de fluxo contínuo](custom-functions-overview.md#streamed-functions).</span><span class="sxs-lookup"><span data-stu-id="1071b-168">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="1071b-169">parameters</span><span class="sxs-lookup"><span data-stu-id="1071b-169">parameters</span></span>

<span data-ttu-id="1071b-170">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1071b-170">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="1071b-171">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="1071b-171">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="1071b-172">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1071b-172">Property</span></span>  |  <span data-ttu-id="1071b-173">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1071b-173">Data type</span></span>  |  <span data-ttu-id="1071b-174">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1071b-174">Required</span></span>  |  <span data-ttu-id="1071b-175">Descrição</span><span class="sxs-lookup"><span data-stu-id="1071b-175">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="1071b-176">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-176">string</span></span>  |  <span data-ttu-id="1071b-177">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-177">No</span></span> |  <span data-ttu-id="1071b-178">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1071b-178">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="1071b-179">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-179">string</span></span>  |  <span data-ttu-id="1071b-180">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-180">No</span></span>  |  <span data-ttu-id="1071b-181">Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="1071b-181">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="1071b-182">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-182">string</span></span>  |  <span data-ttu-id="1071b-183">Sim</span><span class="sxs-lookup"><span data-stu-id="1071b-183">Yes</span></span>  |  <span data-ttu-id="1071b-184">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1071b-184">The name of the parameter.</span></span> <span data-ttu-id="1071b-185">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="1071b-185">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="1071b-186">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-186">string</span></span>  |  <span data-ttu-id="1071b-187">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-187">No</span></span>  |  <span data-ttu-id="1071b-188">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1071b-188">The data type of the parameter.</span></span> <span data-ttu-id="1071b-189">Deve ser **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="1071b-189">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="1071b-190">result</span><span class="sxs-lookup"><span data-stu-id="1071b-190">result</span></span>

<span data-ttu-id="1071b-191">O objeto `results` define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="1071b-191">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1071b-192">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="1071b-192">The following table describes the properties of the project.</span></span>

|  <span data-ttu-id="1071b-193">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1071b-193">Property</span></span>  |  <span data-ttu-id="1071b-194">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1071b-194">Data type</span></span>  |  <span data-ttu-id="1071b-195">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1071b-195">Required</span></span>  |  <span data-ttu-id="1071b-196">Descrição</span><span class="sxs-lookup"><span data-stu-id="1071b-196">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="1071b-197">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-197">string</span></span>  |  <span data-ttu-id="1071b-198">Não</span><span class="sxs-lookup"><span data-stu-id="1071b-198">No</span></span>  |  <span data-ttu-id="1071b-199">Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="1071b-199">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="1071b-200">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1071b-200">string</span></span>  |  <span data-ttu-id="1071b-201">Sim</span><span class="sxs-lookup"><span data-stu-id="1071b-201">Yes</span></span>  |  <span data-ttu-id="1071b-202">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1071b-202">The data type of the parameter.</span></span> <span data-ttu-id="1071b-203">Deve ser **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="1071b-203">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="1071b-204">Confira também</span><span class="sxs-lookup"><span data-stu-id="1071b-204">See also</span></span>

* [<span data-ttu-id="1071b-205">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="1071b-205">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="1071b-206">Runtime para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="1071b-206">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="1071b-207">Práticas recomendadas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1071b-207">Custom functions best practices</span></span>](custom-functions-best-practices.md)