---
ms.date: 11/26/2018
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados de funções personalizadas no Excel
ms.openlocfilehash: a3d4427af2c6ab46133cb4e2fd9ce384a6a8336c
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156590"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="9a15d-103">Metadados de funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="9a15d-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="9a15d-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro do suplemento do Excel, seu projeto de suplemento deve incluir um arquivo de metadados JSON que fornece as informações exigidas pelo Excel para registrar funções personalizadas e disponibilizá-las para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="9a15d-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="9a15d-105">Este artigo descreve o formato do arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="9a15d-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="9a15d-106">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="9a15d-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="9a15d-107">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="9a15d-107">Example metadata</span></span>

<span data-ttu-id="9a15d-108">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="9a15d-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="9a15d-109">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="9a15d-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="9a15d-110">Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="9a15d-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="9a15d-111">functions</span><span class="sxs-lookup"><span data-stu-id="9a15d-111">functions</span></span> 

<span data-ttu-id="9a15d-112">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="9a15d-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="9a15d-113">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="9a15d-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="9a15d-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9a15d-114">Property</span></span>  |  <span data-ttu-id="9a15d-115">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9a15d-115">Data type</span></span>  |  <span data-ttu-id="9a15d-116">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9a15d-116">Required</span></span>  |  <span data-ttu-id="9a15d-117">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a15d-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="9a15d-118">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-118">string</span></span>  |  <span data-ttu-id="9a15d-119">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-119">No</span></span>  |  <span data-ttu-id="9a15d-120">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="9a15d-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="9a15d-121">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="9a15d-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="9a15d-122">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-122">string</span></span>  |   <span data-ttu-id="9a15d-123">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-123">No</span></span>  |  <span data-ttu-id="9a15d-124">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-124">URL that provides information about the function.</span></span> <span data-ttu-id="9a15d-125">(Ela é exibida em um painel de tarefas). Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="9a15d-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="9a15d-126">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-126">string</span></span> | <span data-ttu-id="9a15d-127">Sim</span><span class="sxs-lookup"><span data-stu-id="9a15d-127">Yes</span></span> | <span data-ttu-id="9a15d-128">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-128">A unique ID for the function.</span></span> <span data-ttu-id="9a15d-129">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="9a15d-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="9a15d-130">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-130">string</span></span>  |  <span data-ttu-id="9a15d-131">Sim</span><span class="sxs-lookup"><span data-stu-id="9a15d-131">Yes</span></span>  |  <span data-ttu-id="9a15d-132">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="9a15d-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="9a15d-133">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9a15d-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="9a15d-134">object</span><span class="sxs-lookup"><span data-stu-id="9a15d-134">object</span></span>  |  <span data-ttu-id="9a15d-135">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-135">No</span></span>  |  <span data-ttu-id="9a15d-136">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="9a15d-137">Confira o [objeto options](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="9a15d-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="9a15d-138">array</span><span class="sxs-lookup"><span data-stu-id="9a15d-138">array</span></span>  |  <span data-ttu-id="9a15d-139">Sim</span><span class="sxs-lookup"><span data-stu-id="9a15d-139">Yes</span></span>  |  <span data-ttu-id="9a15d-140">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="9a15d-141">Confira a [matriz de parâmetros](#parameters-array) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="9a15d-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="9a15d-142">object</span><span class="sxs-lookup"><span data-stu-id="9a15d-142">object</span></span>  |  <span data-ttu-id="9a15d-143">Sim</span><span class="sxs-lookup"><span data-stu-id="9a15d-143">Yes</span></span>  |  <span data-ttu-id="9a15d-144">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="9a15d-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="9a15d-145">Confira o [objeto result](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="9a15d-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="9a15d-146">options</span><span class="sxs-lookup"><span data-stu-id="9a15d-146">options</span></span>

<span data-ttu-id="9a15d-147">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="9a15d-148">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-148">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="9a15d-149">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9a15d-149">Property</span></span>  |  <span data-ttu-id="9a15d-150">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9a15d-150">Data type</span></span>  |  <span data-ttu-id="9a15d-151">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9a15d-151">Required</span></span>  |  <span data-ttu-id="9a15d-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a15d-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="9a15d-153">booliano</span><span class="sxs-lookup"><span data-stu-id="9a15d-153">boolean</span></span>  |  <span data-ttu-id="9a15d-154">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-154">No</span></span><br/><br/><span data-ttu-id="9a15d-155">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-155">Default value is `false`.</span></span>  |  <span data-ttu-id="9a15d-156">Se o valor for `true`, o Excel chamará o manipulador `onCanceled` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="9a15d-157">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="9a15d-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="9a15d-158">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="9a15d-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="9a15d-159">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="9a15d-160">Para saber mais, confira [Cancelar uma função](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="9a15d-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="9a15d-161">booliano</span><span class="sxs-lookup"><span data-stu-id="9a15d-161">boolean</span></span>  |  <span data-ttu-id="9a15d-162">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-162">No</span></span><br/><br/><span data-ttu-id="9a15d-163">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-163">Default value is `false`.</span></span>  |  <span data-ttu-id="9a15d-164">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="9a15d-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="9a15d-165">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="9a15d-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="9a15d-166">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="9a15d-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="9a15d-167">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="9a15d-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="9a15d-168">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-168">The function should have no `return` statement.</span></span> <span data-ttu-id="9a15d-169">Em vez disso, o valor resultante é passado como o argumento do método de retorno `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="9a15d-170">Para saber mais informações, confira [Funções de streaming](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="9a15d-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="9a15d-171">booliano</span><span class="sxs-lookup"><span data-stu-id="9a15d-171">boolean</span></span> | <span data-ttu-id="9a15d-172">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-172">No</span></span> <br/><br/><span data-ttu-id="9a15d-173">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-173">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="9a15d-174">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="9a15d-174">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="9a15d-175">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="9a15d-175">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="9a15d-176">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="9a15d-176">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="9a15d-177">parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a15d-177">parameters</span></span>

<span data-ttu-id="9a15d-178">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9a15d-178">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="9a15d-179">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="9a15d-179">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="9a15d-180">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9a15d-180">Property</span></span>  |  <span data-ttu-id="9a15d-181">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9a15d-181">Data type</span></span>  |  <span data-ttu-id="9a15d-182">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9a15d-182">Required</span></span>  |  <span data-ttu-id="9a15d-183">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a15d-183">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="9a15d-184">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-184">string</span></span>  |  <span data-ttu-id="9a15d-185">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-185">No</span></span> |  <span data-ttu-id="9a15d-186">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9a15d-186">A description of the parameter.</span></span> <span data-ttu-id="9a15d-187">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="9a15d-187">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="9a15d-188">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-188">string</span></span>  |  <span data-ttu-id="9a15d-189">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-189">No</span></span>  |  <span data-ttu-id="9a15d-190">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="9a15d-190">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="9a15d-191">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-191">string</span></span>  |  <span data-ttu-id="9a15d-192">Sim</span><span class="sxs-lookup"><span data-stu-id="9a15d-192">Yes</span></span>  |  <span data-ttu-id="9a15d-193">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9a15d-193">The name of the parameter.</span></span> <span data-ttu-id="9a15d-194">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="9a15d-194">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="9a15d-195">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-195">string</span></span>  |  <span data-ttu-id="9a15d-196">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-196">No</span></span>  |  <span data-ttu-id="9a15d-197">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9a15d-197">The data type of the parameter.</span></span> <span data-ttu-id="9a15d-198">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="9a15d-198">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="9a15d-199">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="9a15d-199">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="9a15d-200">booliano</span><span class="sxs-lookup"><span data-stu-id="9a15d-200">boolean</span></span> | <span data-ttu-id="9a15d-201">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-201">No</span></span> | <span data-ttu-id="9a15d-202">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="9a15d-202">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="9a15d-203">Se a propriedade `type` de um parâmetro opcional não for especificada ou definida como `any`, é provável que você tenha problemas, como erros de lint em seu IDE e parâmetros opcionais que não serão exibidos quando a função estiver sendo inserida em uma célula no Excel.</span><span class="sxs-lookup"><span data-stu-id="9a15d-203">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="9a15d-204">A previsão é para ser alterado em dezembro de 2018.</span><span class="sxs-lookup"><span data-stu-id="9a15d-204">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="9a15d-205">result</span><span class="sxs-lookup"><span data-stu-id="9a15d-205">result</span></span>

<span data-ttu-id="9a15d-206">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="9a15d-206">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="9a15d-207">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="9a15d-207">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="9a15d-208">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9a15d-208">Property</span></span>  |  <span data-ttu-id="9a15d-209">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9a15d-209">Data type</span></span>  |  <span data-ttu-id="9a15d-210">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9a15d-210">Required</span></span>  |  <span data-ttu-id="9a15d-211">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a15d-211">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="9a15d-212">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-212">string</span></span>  |  <span data-ttu-id="9a15d-213">Não</span><span class="sxs-lookup"><span data-stu-id="9a15d-213">No</span></span>  |  <span data-ttu-id="9a15d-214">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="9a15d-214">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="9a15d-215">string</span><span class="sxs-lookup"><span data-stu-id="9a15d-215">string</span></span>  |  <span data-ttu-id="9a15d-216">Sim</span><span class="sxs-lookup"><span data-stu-id="9a15d-216">Yes</span></span>  |  <span data-ttu-id="9a15d-217">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9a15d-217">The data type of the parameter.</span></span> <span data-ttu-id="9a15d-218">Deve ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="9a15d-218">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="9a15d-219">Confira também</span><span class="sxs-lookup"><span data-stu-id="9a15d-219">See also</span></span>

* [<span data-ttu-id="9a15d-220">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="9a15d-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="9a15d-221">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="9a15d-221">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="9a15d-222">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="9a15d-222">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="9a15d-223">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="9a15d-223">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
