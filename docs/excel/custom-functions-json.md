---
ms.date: 03/29/2019
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados de funções personalizadas no Excel (visualização)
localization_priority: Normal
ms.openlocfilehash: 3703699348e99fd076fe0e3affac88038e3aaf59
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914253"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="52c68-103">Metadados de funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="52c68-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="52c68-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o projeto do suplemento inclui um arquivo de metadados JSON que fornece as informações que o Excel requer para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="52c68-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="52c68-105">Este arquivo é gerado:</span><span class="sxs-lookup"><span data-stu-id="52c68-105">This file is generated either:</span></span>

- <span data-ttu-id="52c68-106">por você, em um arquivo JSON manuscrito</span><span class="sxs-lookup"><span data-stu-id="52c68-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="52c68-107">nos comentários do JSDoc inseridos no início da função</span><span class="sxs-lookup"><span data-stu-id="52c68-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="52c68-108">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52c68-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="52c68-109">Este artigo descreve o formato do arquivo de metadados JSON, supondo que você o esteja escrevendo à mão.</span><span class="sxs-lookup"><span data-stu-id="52c68-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="52c68-110">Para obter informações sobre a geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="52c68-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="52c68-111">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="52c68-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="52c68-112">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="52c68-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="52c68-113">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="52c68-113">Example metadata</span></span>

<span data-ttu-id="52c68-114">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="52c68-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="52c68-115">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="52c68-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="52c68-116">Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="52c68-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="52c68-117">functions</span><span class="sxs-lookup"><span data-stu-id="52c68-117">functions</span></span> 

<span data-ttu-id="52c68-118">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="52c68-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="52c68-119">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="52c68-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="52c68-120">Propriedade</span><span class="sxs-lookup"><span data-stu-id="52c68-120">Property</span></span>  |  <span data-ttu-id="52c68-121">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="52c68-121">Data type</span></span>  |  <span data-ttu-id="52c68-122">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="52c68-122">Required</span></span>  |  <span data-ttu-id="52c68-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="52c68-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="52c68-124">string</span><span class="sxs-lookup"><span data-stu-id="52c68-124">string</span></span>  |  <span data-ttu-id="52c68-125">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-125">No</span></span>  |  <span data-ttu-id="52c68-126">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="52c68-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="52c68-127">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="52c68-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="52c68-128">string</span><span class="sxs-lookup"><span data-stu-id="52c68-128">string</span></span>  |   <span data-ttu-id="52c68-129">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-129">No</span></span>  |  <span data-ttu-id="52c68-130">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="52c68-130">URL that provides information about the function.</span></span> <span data-ttu-id="52c68-131">(Ela é exibida em um painel de tarefas). Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="52c68-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="52c68-132">string</span><span class="sxs-lookup"><span data-stu-id="52c68-132">string</span></span> | <span data-ttu-id="52c68-133">Sim</span><span class="sxs-lookup"><span data-stu-id="52c68-133">Yes</span></span> | <span data-ttu-id="52c68-134">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="52c68-134">A unique ID for the function.</span></span> <span data-ttu-id="52c68-135">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="52c68-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="52c68-136">string</span><span class="sxs-lookup"><span data-stu-id="52c68-136">string</span></span>  |  <span data-ttu-id="52c68-137">Sim</span><span class="sxs-lookup"><span data-stu-id="52c68-137">Yes</span></span>  |  <span data-ttu-id="52c68-138">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="52c68-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="52c68-139">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="52c68-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="52c68-140">objeto</span><span class="sxs-lookup"><span data-stu-id="52c68-140">object</span></span>  |  <span data-ttu-id="52c68-141">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-141">No</span></span>  |  <span data-ttu-id="52c68-142">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="52c68-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="52c68-143">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="52c68-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="52c68-144">array</span><span class="sxs-lookup"><span data-stu-id="52c68-144">array</span></span>  |  <span data-ttu-id="52c68-145">Sim</span><span class="sxs-lookup"><span data-stu-id="52c68-145">Yes</span></span>  |  <span data-ttu-id="52c68-146">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="52c68-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="52c68-147">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="52c68-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="52c68-148">object</span><span class="sxs-lookup"><span data-stu-id="52c68-148">object</span></span>  |  <span data-ttu-id="52c68-149">Sim</span><span class="sxs-lookup"><span data-stu-id="52c68-149">Yes</span></span>  |  <span data-ttu-id="52c68-150">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="52c68-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="52c68-151">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="52c68-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="52c68-152">options</span><span class="sxs-lookup"><span data-stu-id="52c68-152">options</span></span>

<span data-ttu-id="52c68-153">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="52c68-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="52c68-154">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="52c68-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="52c68-155">Propriedade</span><span class="sxs-lookup"><span data-stu-id="52c68-155">Property</span></span>  |  <span data-ttu-id="52c68-156">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="52c68-156">Data type</span></span>  |  <span data-ttu-id="52c68-157">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="52c68-157">Required</span></span>  |  <span data-ttu-id="52c68-158">Descrição</span><span class="sxs-lookup"><span data-stu-id="52c68-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="52c68-159">booliano</span><span class="sxs-lookup"><span data-stu-id="52c68-159">boolean</span></span>  |  <span data-ttu-id="52c68-160">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-160">No</span></span><br/><br/><span data-ttu-id="52c68-161">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="52c68-161">Default value is `false`.</span></span>  |  <span data-ttu-id="52c68-162">Se o valor for `true`, o Excel chamará o manipulador `onCanceled` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="52c68-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="52c68-163">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="52c68-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="52c68-164">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="52c68-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="52c68-165">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="52c68-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="52c68-166">Para saber mais, confira [Cancelar uma função](custom-functions-web-reqs.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="52c68-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="52c68-167">booliano</span><span class="sxs-lookup"><span data-stu-id="52c68-167">boolean</span></span> | <span data-ttu-id="52c68-168">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-168">No</span></span> <br/><br/><span data-ttu-id="52c68-169">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="52c68-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="52c68-170">Se true, sua função personalizada pode acessar o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="52c68-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="52c68-171">Para obter o endereço da célula que chamou sua função personalizada, use Context. Address em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="52c68-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="52c68-172">Para saber mais, confira [determinar quais célula chamada sua função personalizada](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span><span class="sxs-lookup"><span data-stu-id="52c68-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="52c68-173">As funções personalizadas não podem ser definidas como streaming e requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="52c68-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="52c68-174">Ao usar essa opção, o parâmetro "invocationContext" deve ser o último parâmetro passado em opções.</span><span class="sxs-lookup"><span data-stu-id="52c68-174">When using this option, the 'invocationContext' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="52c68-175">booliano</span><span class="sxs-lookup"><span data-stu-id="52c68-175">boolean</span></span>  |  <span data-ttu-id="52c68-176">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-176">No</span></span><br/><br/><span data-ttu-id="52c68-177">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="52c68-177">Default value is `false`.</span></span>  |  <span data-ttu-id="52c68-178">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="52c68-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="52c68-179">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="52c68-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="52c68-180">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="52c68-180">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="52c68-181">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="52c68-181">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="52c68-182">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="52c68-182">The function should have no `return` statement.</span></span> <span data-ttu-id="52c68-183">Em vez disso, o valor resultante é passado como o argumento do método de retorno `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="52c68-183">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="52c68-184">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="52c68-184">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="52c68-185">booliano</span><span class="sxs-lookup"><span data-stu-id="52c68-185">boolean</span></span> | <span data-ttu-id="52c68-186">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-186">No</span></span> <br/><br/><span data-ttu-id="52c68-187">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="52c68-187">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="52c68-188">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="52c68-188">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="52c68-189">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="52c68-189">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="52c68-190">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="52c68-190">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="52c68-191">parâmetros</span><span class="sxs-lookup"><span data-stu-id="52c68-191">parameters</span></span>

<span data-ttu-id="52c68-192">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="52c68-192">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="52c68-193">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="52c68-193">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="52c68-194">Propriedade</span><span class="sxs-lookup"><span data-stu-id="52c68-194">Property</span></span>  |  <span data-ttu-id="52c68-195">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="52c68-195">Data type</span></span>  |  <span data-ttu-id="52c68-196">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="52c68-196">Required</span></span>  |  <span data-ttu-id="52c68-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="52c68-197">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="52c68-198">string</span><span class="sxs-lookup"><span data-stu-id="52c68-198">string</span></span>  |  <span data-ttu-id="52c68-199">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-199">No</span></span> |  <span data-ttu-id="52c68-200">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="52c68-200">A description of the parameter.</span></span> <span data-ttu-id="52c68-201">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="52c68-201">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="52c68-202">string</span><span class="sxs-lookup"><span data-stu-id="52c68-202">string</span></span>  |  <span data-ttu-id="52c68-203">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-203">No</span></span>  |  <span data-ttu-id="52c68-204">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="52c68-204">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="52c68-205">string</span><span class="sxs-lookup"><span data-stu-id="52c68-205">string</span></span>  |  <span data-ttu-id="52c68-206">Sim</span><span class="sxs-lookup"><span data-stu-id="52c68-206">Yes</span></span>  |  <span data-ttu-id="52c68-207">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="52c68-207">The name of the parameter.</span></span> <span data-ttu-id="52c68-208">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="52c68-208">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="52c68-209">string</span><span class="sxs-lookup"><span data-stu-id="52c68-209">string</span></span>  |  <span data-ttu-id="52c68-210">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-210">No</span></span>  |  <span data-ttu-id="52c68-211">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="52c68-211">The data type of the parameter.</span></span> <span data-ttu-id="52c68-212">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="52c68-212">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="52c68-213">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="52c68-213">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="52c68-214">booliano</span><span class="sxs-lookup"><span data-stu-id="52c68-214">boolean</span></span> | <span data-ttu-id="52c68-215">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-215">No</span></span> | <span data-ttu-id="52c68-216">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="52c68-216">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="52c68-217">Se a propriedade `type` de um parâmetro opcional não for especificada ou definida como `any`, é provável que você tenha problemas, como erros de lint em seu IDE e parâmetros opcionais que não serão exibidos quando a função estiver sendo inserida em uma célula no Excel.</span><span class="sxs-lookup"><span data-stu-id="52c68-217">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="52c68-218">A previsão é para ser alterado em dezembro de 2018.</span><span class="sxs-lookup"><span data-stu-id="52c68-218">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="52c68-219">result</span><span class="sxs-lookup"><span data-stu-id="52c68-219">result</span></span>

<span data-ttu-id="52c68-220">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="52c68-220">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="52c68-221">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="52c68-221">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="52c68-222">Propriedade</span><span class="sxs-lookup"><span data-stu-id="52c68-222">Property</span></span>  |  <span data-ttu-id="52c68-223">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="52c68-223">Data type</span></span>  |  <span data-ttu-id="52c68-224">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="52c68-224">Required</span></span>  |  <span data-ttu-id="52c68-225">Descrição</span><span class="sxs-lookup"><span data-stu-id="52c68-225">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="52c68-226">string</span><span class="sxs-lookup"><span data-stu-id="52c68-226">string</span></span>  |  <span data-ttu-id="52c68-227">Não</span><span class="sxs-lookup"><span data-stu-id="52c68-227">No</span></span>  |  <span data-ttu-id="52c68-228">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="52c68-228">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="52c68-229">string</span><span class="sxs-lookup"><span data-stu-id="52c68-229">string</span></span>  |  <span data-ttu-id="52c68-230">Sim</span><span class="sxs-lookup"><span data-stu-id="52c68-230">Yes</span></span>  |  <span data-ttu-id="52c68-231">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="52c68-231">The data type of the parameter.</span></span> <span data-ttu-id="52c68-232">Deve ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="52c68-232">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="52c68-233">Confira também</span><span class="sxs-lookup"><span data-stu-id="52c68-233">See also</span></span>

* [<span data-ttu-id="52c68-234">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="52c68-234">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="52c68-235">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="52c68-235">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="52c68-236">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="52c68-236">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="52c68-237">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="52c68-237">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="52c68-238">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="52c68-238">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
