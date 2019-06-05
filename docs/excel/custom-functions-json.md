---
ms.date: 05/30/2019
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: e51e4e8ee89eb1f345ee0c564e9b2ff8119806b2
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706120"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="9252c-103">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="9252c-103">Custom functions metadata</span></span>

<span data-ttu-id="9252c-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o projeto do suplemento inclui um arquivo de metadados JSON que fornece as informações que o Excel requer para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="9252c-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="9252c-105">Este arquivo é gerado:</span><span class="sxs-lookup"><span data-stu-id="9252c-105">This file is generated either:</span></span>

- <span data-ttu-id="9252c-106">Por você, em um arquivo JSON manuscrito</span><span class="sxs-lookup"><span data-stu-id="9252c-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="9252c-107">Nos comentários do JSDoc inseridos no início da função</span><span class="sxs-lookup"><span data-stu-id="9252c-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="9252c-108">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9252c-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="9252c-109">Este artigo descreve o formato do arquivo de metadados JSON, supondo que você o esteja escrevendo à mão.</span><span class="sxs-lookup"><span data-stu-id="9252c-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="9252c-110">Para obter informações sobre a geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="9252c-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="9252c-111">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="9252c-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="9252c-112">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="9252c-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="9252c-113">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="9252c-113">Example metadata</span></span>

<span data-ttu-id="9252c-114">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="9252c-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="9252c-115">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="9252c-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
> <span data-ttu-id="9252c-116">Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="9252c-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="9252c-117">functions</span><span class="sxs-lookup"><span data-stu-id="9252c-117">functions</span></span> 

<span data-ttu-id="9252c-118">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="9252c-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="9252c-119">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="9252c-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="9252c-120">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9252c-120">Property</span></span>  |  <span data-ttu-id="9252c-121">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9252c-121">Data type</span></span>  |  <span data-ttu-id="9252c-122">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9252c-122">Required</span></span>  |  <span data-ttu-id="9252c-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="9252c-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="9252c-124">string</span><span class="sxs-lookup"><span data-stu-id="9252c-124">string</span></span>  |  <span data-ttu-id="9252c-125">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-125">No</span></span>  |  <span data-ttu-id="9252c-126">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="9252c-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="9252c-127">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="9252c-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="9252c-128">string</span><span class="sxs-lookup"><span data-stu-id="9252c-128">string</span></span>  |   <span data-ttu-id="9252c-129">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-129">No</span></span>  |  <span data-ttu-id="9252c-130">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="9252c-130">URL that provides information about the function.</span></span> <span data-ttu-id="9252c-131">(Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="9252c-131">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="9252c-132">string</span><span class="sxs-lookup"><span data-stu-id="9252c-132">string</span></span> | <span data-ttu-id="9252c-133">Sim</span><span class="sxs-lookup"><span data-stu-id="9252c-133">Yes</span></span> | <span data-ttu-id="9252c-134">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="9252c-134">A unique ID for the function.</span></span> <span data-ttu-id="9252c-135">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="9252c-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="9252c-136">string</span><span class="sxs-lookup"><span data-stu-id="9252c-136">string</span></span>  |  <span data-ttu-id="9252c-137">Sim</span><span class="sxs-lookup"><span data-stu-id="9252c-137">Yes</span></span>  |  <span data-ttu-id="9252c-138">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="9252c-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="9252c-139">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9252c-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="9252c-140">objeto</span><span class="sxs-lookup"><span data-stu-id="9252c-140">object</span></span>  |  <span data-ttu-id="9252c-141">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-141">No</span></span>  |  <span data-ttu-id="9252c-142">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="9252c-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="9252c-143">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="9252c-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="9252c-144">array</span><span class="sxs-lookup"><span data-stu-id="9252c-144">array</span></span>  |  <span data-ttu-id="9252c-145">Sim</span><span class="sxs-lookup"><span data-stu-id="9252c-145">Yes</span></span>  |  <span data-ttu-id="9252c-146">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="9252c-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="9252c-147">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="9252c-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="9252c-148">object</span><span class="sxs-lookup"><span data-stu-id="9252c-148">object</span></span>  |  <span data-ttu-id="9252c-149">Sim</span><span class="sxs-lookup"><span data-stu-id="9252c-149">Yes</span></span>  |  <span data-ttu-id="9252c-150">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="9252c-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="9252c-151">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="9252c-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="9252c-152">options</span><span class="sxs-lookup"><span data-stu-id="9252c-152">options</span></span>

<span data-ttu-id="9252c-153">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="9252c-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="9252c-154">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="9252c-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="9252c-155">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9252c-155">Property</span></span>  |  <span data-ttu-id="9252c-156">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9252c-156">Data type</span></span>  |  <span data-ttu-id="9252c-157">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9252c-157">Required</span></span>  |  <span data-ttu-id="9252c-158">Descrição</span><span class="sxs-lookup"><span data-stu-id="9252c-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="9252c-159">booliano</span><span class="sxs-lookup"><span data-stu-id="9252c-159">boolean</span></span>  |  <span data-ttu-id="9252c-160">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-160">No</span></span><br/><br/><span data-ttu-id="9252c-161">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9252c-161">Default value is `false`.</span></span>  |  <span data-ttu-id="9252c-162">Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="9252c-162">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="9252c-163">As funções de cancelamento normalmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados.</span><span class="sxs-lookup"><span data-stu-id="9252c-163">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="9252c-164">Uma função não pode ser streaming e cancelamento.</span><span class="sxs-lookup"><span data-stu-id="9252c-164">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="9252c-165">Para obter mais informações, consulte a observação próxima ao final de [fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="9252c-165">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="9252c-166">booliano</span><span class="sxs-lookup"><span data-stu-id="9252c-166">boolean</span></span> | <span data-ttu-id="9252c-167">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-167">No</span></span> <br/><br/><span data-ttu-id="9252c-168">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9252c-168">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="9252c-169">Se true, sua função personalizada pode acessar o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="9252c-169">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="9252c-170">Para obter o endereço da célula que chamou sua função personalizada, use Context. Address em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="9252c-170">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="9252c-171">Para saber mais, confira [determinar quais célula chamada sua função personalizada](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span><span class="sxs-lookup"><span data-stu-id="9252c-171">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="9252c-172">As funções personalizadas não podem ser definidas como streaming e requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="9252c-172">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="9252c-173">Ao usar essa opção, o parâmetro "invocar" deve ser o último parâmetro passado em opções.</span><span class="sxs-lookup"><span data-stu-id="9252c-173">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="9252c-174">booliano</span><span class="sxs-lookup"><span data-stu-id="9252c-174">boolean</span></span>  |  <span data-ttu-id="9252c-175">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-175">No</span></span><br/><br/><span data-ttu-id="9252c-176">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9252c-176">Default value is `false`.</span></span>  |  <span data-ttu-id="9252c-177">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="9252c-177">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="9252c-178">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="9252c-178">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="9252c-179">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="9252c-179">The function should have no `return` statement.</span></span> <span data-ttu-id="9252c-180">Em vez disso, o valor resultante é passado como o argumento do método de retorno `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="9252c-180">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="9252c-181">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="9252c-181">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="9252c-182">booliano</span><span class="sxs-lookup"><span data-stu-id="9252c-182">boolean</span></span> | <span data-ttu-id="9252c-183">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-183">No</span></span> <br/><br/><span data-ttu-id="9252c-184">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="9252c-184">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="9252c-185">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="9252c-185">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="9252c-186">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="9252c-186">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="9252c-187">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="9252c-187">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="9252c-188">parâmetros</span><span class="sxs-lookup"><span data-stu-id="9252c-188">parameters</span></span>

<span data-ttu-id="9252c-189">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9252c-189">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="9252c-190">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="9252c-190">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="9252c-191">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9252c-191">Property</span></span>  |  <span data-ttu-id="9252c-192">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9252c-192">Data type</span></span>  |  <span data-ttu-id="9252c-193">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9252c-193">Required</span></span>  |  <span data-ttu-id="9252c-194">Descrição</span><span class="sxs-lookup"><span data-stu-id="9252c-194">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="9252c-195">string</span><span class="sxs-lookup"><span data-stu-id="9252c-195">string</span></span>  |  <span data-ttu-id="9252c-196">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-196">No</span></span> |  <span data-ttu-id="9252c-197">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9252c-197">A description of the parameter.</span></span> <span data-ttu-id="9252c-198">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="9252c-198">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="9252c-199">string</span><span class="sxs-lookup"><span data-stu-id="9252c-199">string</span></span>  |  <span data-ttu-id="9252c-200">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-200">No</span></span>  |  <span data-ttu-id="9252c-201">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="9252c-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="9252c-202">string</span><span class="sxs-lookup"><span data-stu-id="9252c-202">string</span></span>  |  <span data-ttu-id="9252c-203">Sim</span><span class="sxs-lookup"><span data-stu-id="9252c-203">Yes</span></span>  |  <span data-ttu-id="9252c-204">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9252c-204">The name of the parameter.</span></span> <span data-ttu-id="9252c-205">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="9252c-205">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="9252c-206">string</span><span class="sxs-lookup"><span data-stu-id="9252c-206">string</span></span>  |  <span data-ttu-id="9252c-207">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-207">No</span></span>  |  <span data-ttu-id="9252c-208">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9252c-208">The data type of the parameter.</span></span> <span data-ttu-id="9252c-209">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="9252c-209">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="9252c-210">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="9252c-210">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="9252c-211">booliano</span><span class="sxs-lookup"><span data-stu-id="9252c-211">boolean</span></span> | <span data-ttu-id="9252c-212">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-212">No</span></span> | <span data-ttu-id="9252c-213">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="9252c-213">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="9252c-214">result</span><span class="sxs-lookup"><span data-stu-id="9252c-214">result</span></span>

<span data-ttu-id="9252c-215">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="9252c-215">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="9252c-216">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="9252c-216">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="9252c-217">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9252c-217">Property</span></span>  |  <span data-ttu-id="9252c-218">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="9252c-218">Data type</span></span>  |  <span data-ttu-id="9252c-219">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9252c-219">Required</span></span>  |  <span data-ttu-id="9252c-220">Descrição</span><span class="sxs-lookup"><span data-stu-id="9252c-220">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="9252c-221">string</span><span class="sxs-lookup"><span data-stu-id="9252c-221">string</span></span>  |  <span data-ttu-id="9252c-222">Não</span><span class="sxs-lookup"><span data-stu-id="9252c-222">No</span></span>  |  <span data-ttu-id="9252c-223">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="9252c-223">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="9252c-224">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="9252c-224">Next steps</span></span>
<span data-ttu-id="9252c-225">Conheça as [práticas recomendadas para nomear sua função](custom-functions-naming.md) ou descubra como [localizar sua função](custom-functions-localize.md) usando o método JSON manuscrito descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="9252c-225">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="9252c-226">Confira também</span><span class="sxs-lookup"><span data-stu-id="9252c-226">See also</span></span>

* [<span data-ttu-id="9252c-227">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="9252c-227">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="9252c-228">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="9252c-228">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* <span data-ttu-id="9252c-229">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="9252c-229">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="9252c-230">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="9252c-230">Create custom functions in Excel</span></span>](custom-functions-overview.md)