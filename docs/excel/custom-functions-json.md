---
ms.date: 10/17/2018
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados de funções personalizadas no Excel
ms.openlocfilehash: 0c77474188a2deefd23a73bb64e87569bb1fa52a
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298541"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="93720-103">Metadados de funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="93720-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="93720-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro do suplemento do Excel, seu projeto de suplemento deve incluir um arquivo de metadados JSON que fornece as informações exigidas pelo Excel para registrar funções personalizadas e disponibilizá-las para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="93720-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="93720-105">Este artigo descreve o formato do arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="93720-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="93720-106">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="93720-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="93720-107">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="93720-107">Example metadata</span></span>

<span data-ttu-id="93720-108">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="93720-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="93720-109">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="93720-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
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
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
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
      "description":  "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
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
> <span data-ttu-id="93720-110">Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="93720-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="93720-111">functions</span><span class="sxs-lookup"><span data-stu-id="93720-111">functions</span></span> 

<span data-ttu-id="93720-112">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="93720-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="93720-113">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="93720-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="93720-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="93720-114">Property</span></span>  |  <span data-ttu-id="93720-115">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="93720-115">Data type</span></span>  |  <span data-ttu-id="93720-116">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="93720-116">Required</span></span>  |  <span data-ttu-id="93720-117">Descrição</span><span class="sxs-lookup"><span data-stu-id="93720-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="93720-118">string</span><span class="sxs-lookup"><span data-stu-id="93720-118">string</span></span>  |  <span data-ttu-id="93720-119">Não</span><span class="sxs-lookup"><span data-stu-id="93720-119">No</span></span>  |  <span data-ttu-id="93720-120">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="93720-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="93720-121">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="93720-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="93720-122">string</span><span class="sxs-lookup"><span data-stu-id="93720-122">string</span></span>  |   <span data-ttu-id="93720-123">Não</span><span class="sxs-lookup"><span data-stu-id="93720-123">No</span></span>  |  <span data-ttu-id="93720-124">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="93720-124">URL that provides information about the function.</span></span> <span data-ttu-id="93720-125">(Ela é exibida em um painel de tarefas). Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="93720-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="93720-126">string</span><span class="sxs-lookup"><span data-stu-id="93720-126">string</span></span> | <span data-ttu-id="93720-127">Sim</span><span class="sxs-lookup"><span data-stu-id="93720-127">Yes</span></span> | <span data-ttu-id="93720-128">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="93720-128">A unique ID for the group.</span></span> <span data-ttu-id="93720-129">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="93720-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="93720-130">string</span><span class="sxs-lookup"><span data-stu-id="93720-130">string</span></span>  |  <span data-ttu-id="93720-131">Sim</span><span class="sxs-lookup"><span data-stu-id="93720-131">Yes</span></span>  |  <span data-ttu-id="93720-132">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="93720-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="93720-133">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="93720-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="93720-134">object</span><span class="sxs-lookup"><span data-stu-id="93720-134">object</span></span>  |  <span data-ttu-id="93720-135">Não</span><span class="sxs-lookup"><span data-stu-id="93720-135">No</span></span>  |  <span data-ttu-id="93720-136">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="93720-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="93720-137">Confira o [objeto options](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="93720-137">See object load [options](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="93720-138">array</span><span class="sxs-lookup"><span data-stu-id="93720-138">array</span></span>  |  <span data-ttu-id="93720-139">Sim</span><span class="sxs-lookup"><span data-stu-id="93720-139">Yes</span></span>  |  <span data-ttu-id="93720-140">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="93720-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="93720-141">Confira a [matriz de parâmetros](#parameters-array) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="93720-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="93720-142">object</span><span class="sxs-lookup"><span data-stu-id="93720-142">object</span></span>  |  <span data-ttu-id="93720-143">Sim</span><span class="sxs-lookup"><span data-stu-id="93720-143">Yes</span></span>  |  <span data-ttu-id="93720-144">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="93720-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="93720-145">Confira o [objeto result](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="93720-145">See object load [options](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="93720-146">options</span><span class="sxs-lookup"><span data-stu-id="93720-146">options</span></span>

<span data-ttu-id="93720-147">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="93720-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="93720-148">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="93720-148">The following table lists the parts of the `options` claim.</span></span>

|  <span data-ttu-id="93720-149">Propriedade</span><span class="sxs-lookup"><span data-stu-id="93720-149">Property</span></span>  |  <span data-ttu-id="93720-150">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="93720-150">Data type</span></span>  |  <span data-ttu-id="93720-151">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="93720-151">Required</span></span>  |  <span data-ttu-id="93720-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="93720-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="93720-153">booliano</span><span class="sxs-lookup"><span data-stu-id="93720-153">boolean</span></span>  |  <span data-ttu-id="93720-154">Não</span><span class="sxs-lookup"><span data-stu-id="93720-154">No</span></span><br/><br/><span data-ttu-id="93720-155">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="93720-155">Default value is `false`.</span></span>  |  <span data-ttu-id="93720-156">Se o valor for `true`, o Excel chamará o manipulador `onCanceled` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="93720-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="93720-157">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="93720-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="93720-158">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="93720-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="93720-159">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="93720-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="93720-160">Para saber mais, confira [Cancelar uma função](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="93720-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="93720-161">booliano</span><span class="sxs-lookup"><span data-stu-id="93720-161">boolean</span></span>  |  <span data-ttu-id="93720-162">Não</span><span class="sxs-lookup"><span data-stu-id="93720-162">No</span></span><br/><br/><span data-ttu-id="93720-163">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="93720-163">Default value is `false`.</span></span>  |  <span data-ttu-id="93720-164">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="93720-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="93720-165">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="93720-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="93720-166">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="93720-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="93720-167">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="93720-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="93720-168">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="93720-168">The function should have no `return` statement.</span></span> <span data-ttu-id="93720-169">Em vez disso, o valor resultante é passado como o argumento do método de retorno `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="93720-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="93720-170">Para saber mais informações, confira [Funções de streaming](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="93720-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="93720-171">parameters</span><span class="sxs-lookup"><span data-stu-id="93720-171">parameters</span></span>

<span data-ttu-id="93720-172">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="93720-172">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="93720-173">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="93720-173">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="93720-174">Propriedade</span><span class="sxs-lookup"><span data-stu-id="93720-174">Property</span></span>  |  <span data-ttu-id="93720-175">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="93720-175">Data type</span></span>  |  <span data-ttu-id="93720-176">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="93720-176">Required</span></span>  |  <span data-ttu-id="93720-177">Descrição</span><span class="sxs-lookup"><span data-stu-id="93720-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="93720-178">string</span><span class="sxs-lookup"><span data-stu-id="93720-178">string</span></span>  |  <span data-ttu-id="93720-179">Não</span><span class="sxs-lookup"><span data-stu-id="93720-179">No</span></span> |  <span data-ttu-id="93720-180">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="93720-180">A description of the error.</span></span> <span data-ttu-id="93720-181">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="93720-181">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="93720-182">string</span><span class="sxs-lookup"><span data-stu-id="93720-182">string</span></span>  |  <span data-ttu-id="93720-183">Não</span><span class="sxs-lookup"><span data-stu-id="93720-183">No</span></span>  |  <span data-ttu-id="93720-184">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="93720-184">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="93720-185">string</span><span class="sxs-lookup"><span data-stu-id="93720-185">string</span></span>  |  <span data-ttu-id="93720-186">Sim</span><span class="sxs-lookup"><span data-stu-id="93720-186">Yes</span></span>  |  <span data-ttu-id="93720-187">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="93720-187">The name of the parameter.</span></span> <span data-ttu-id="93720-188">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="93720-188">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="93720-189">string</span><span class="sxs-lookup"><span data-stu-id="93720-189">string</span></span>  |  <span data-ttu-id="93720-190">Não</span><span class="sxs-lookup"><span data-stu-id="93720-190">No</span></span>  |  <span data-ttu-id="93720-191">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="93720-191">The type of the parameter.</span></span> <span data-ttu-id="93720-192">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="93720-192">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="93720-193">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="93720-193">If this property is not specified, the data type defaults to **any**.</span></span> |

## <a name="result"></a><span data-ttu-id="93720-194">result</span><span class="sxs-lookup"><span data-stu-id="93720-194">result</span></span>

<span data-ttu-id="93720-195">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="93720-195">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="93720-196">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="93720-196">The following table lists the parts of the `result` claim.</span></span>

|  <span data-ttu-id="93720-197">Propriedade</span><span class="sxs-lookup"><span data-stu-id="93720-197">Property</span></span>  |  <span data-ttu-id="93720-198">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="93720-198">Data type</span></span>  |  <span data-ttu-id="93720-199">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="93720-199">Required</span></span>  |  <span data-ttu-id="93720-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="93720-200">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="93720-201">string</span><span class="sxs-lookup"><span data-stu-id="93720-201">string</span></span>  |  <span data-ttu-id="93720-202">Não</span><span class="sxs-lookup"><span data-stu-id="93720-202">No</span></span>  |  <span data-ttu-id="93720-203">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="93720-203">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="93720-204">string</span><span class="sxs-lookup"><span data-stu-id="93720-204">string</span></span>  |  <span data-ttu-id="93720-205">Sim</span><span class="sxs-lookup"><span data-stu-id="93720-205">Yes</span></span>  |  <span data-ttu-id="93720-206">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="93720-206">The type of the parameter.</span></span> <span data-ttu-id="93720-207">Deve ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="93720-207">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="93720-208">Confira também</span><span class="sxs-lookup"><span data-stu-id="93720-208">See also</span></span>

* [<span data-ttu-id="93720-209">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="93720-209">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="93720-210">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="93720-210">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="93720-211">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="93720-211">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="93720-212">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="93720-212">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
