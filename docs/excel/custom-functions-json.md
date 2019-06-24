---
ms.date: 06/20/2019
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: f97a339972a8ac134bd30c87b86c4701cb4b5fc4
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127867"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="1a9a9-103">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1a9a9-103">Custom functions metadata</span></span>

<span data-ttu-id="1a9a9-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o projeto do suplemento inclui um arquivo de metadados JSON que fornece as informações que o Excel requer para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="1a9a9-105">Este arquivo é gerado:</span><span class="sxs-lookup"><span data-stu-id="1a9a9-105">This file is generated either:</span></span>

- <span data-ttu-id="1a9a9-106">Por você, em um arquivo JSON manuscrito</span><span class="sxs-lookup"><span data-stu-id="1a9a9-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="1a9a9-107">Nos comentários do JSDoc inseridos no início da função</span><span class="sxs-lookup"><span data-stu-id="1a9a9-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="1a9a9-108">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="1a9a9-109">Este artigo descreve o formato do arquivo de metadados JSON, supondo que você o esteja escrevendo à mão.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="1a9a9-110">Para obter informações sobre a geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="1a9a9-111">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="1a9a9-112">As configurações de servidor no servidor que hospeda o arquivo JSON devem ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que as funções personalizadas funcionem corretamente no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="1a9a9-113">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="1a9a9-113">Example metadata</span></span>

<span data-ttu-id="1a9a9-114">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="1a9a9-115">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="1a9a9-116">Um arquivo JSON de exemplo completo está disponível no histórico de confirmação do repositório do GitHub do [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) .</span><span class="sxs-lookup"><span data-stu-id="1a9a9-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="1a9a9-117">À medida que o projeto é ajustado para gerar JSON automaticamente, um exemplo completo de JSON manuscrito só está disponível em versões anteriores do projeto.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-117">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="functions"></a><span data-ttu-id="1a9a9-118">functions</span><span class="sxs-lookup"><span data-stu-id="1a9a9-118">functions</span></span> 

<span data-ttu-id="1a9a9-119">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-119">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="1a9a9-120">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-120">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="1a9a9-121">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1a9a9-121">Property</span></span>  |  <span data-ttu-id="1a9a9-122">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1a9a9-122">Data type</span></span>  |  <span data-ttu-id="1a9a9-123">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1a9a9-123">Required</span></span>  |  <span data-ttu-id="1a9a9-124">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a9a9-124">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="1a9a9-125">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-125">string</span></span>  |  <span data-ttu-id="1a9a9-126">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-126">No</span></span>  |  <span data-ttu-id="1a9a9-127">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-127">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="1a9a9-128">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-128">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="1a9a9-129">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-129">string</span></span>  |   <span data-ttu-id="1a9a9-130">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-130">No</span></span>  |  <span data-ttu-id="1a9a9-131">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-131">URL that provides information about the function.</span></span> <span data-ttu-id="1a9a9-132">(Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-132">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="1a9a9-133">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-133">string</span></span> | <span data-ttu-id="1a9a9-134">Sim</span><span class="sxs-lookup"><span data-stu-id="1a9a9-134">Yes</span></span> | <span data-ttu-id="1a9a9-135">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-135">A unique ID for the function.</span></span> <span data-ttu-id="1a9a9-136">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-136">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="1a9a9-137">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-137">string</span></span>  |  <span data-ttu-id="1a9a9-138">Sim</span><span class="sxs-lookup"><span data-stu-id="1a9a9-138">Yes</span></span>  |  <span data-ttu-id="1a9a9-139">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-139">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="1a9a9-140">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-140">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="1a9a9-141">objeto</span><span class="sxs-lookup"><span data-stu-id="1a9a9-141">object</span></span>  |  <span data-ttu-id="1a9a9-142">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-142">No</span></span>  |  <span data-ttu-id="1a9a9-143">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-143">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1a9a9-144">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-144">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="1a9a9-145">array</span><span class="sxs-lookup"><span data-stu-id="1a9a9-145">array</span></span>  |  <span data-ttu-id="1a9a9-146">Sim</span><span class="sxs-lookup"><span data-stu-id="1a9a9-146">Yes</span></span>  |  <span data-ttu-id="1a9a9-147">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-147">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="1a9a9-148">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-148">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="1a9a9-149">object</span><span class="sxs-lookup"><span data-stu-id="1a9a9-149">object</span></span>  |  <span data-ttu-id="1a9a9-150">Sim</span><span class="sxs-lookup"><span data-stu-id="1a9a9-150">Yes</span></span>  |  <span data-ttu-id="1a9a9-151">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-151">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1a9a9-152">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-152">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="1a9a9-153">options</span><span class="sxs-lookup"><span data-stu-id="1a9a9-153">options</span></span>

<span data-ttu-id="1a9a9-154">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-154">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1a9a9-155">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-155">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="1a9a9-156">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1a9a9-156">Property</span></span>  |  <span data-ttu-id="1a9a9-157">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1a9a9-157">Data type</span></span>  |  <span data-ttu-id="1a9a9-158">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1a9a9-158">Required</span></span>  |  <span data-ttu-id="1a9a9-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a9a9-159">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="1a9a9-160">booliano</span><span class="sxs-lookup"><span data-stu-id="1a9a9-160">boolean</span></span>  |  <span data-ttu-id="1a9a9-161">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-161">No</span></span><br/><br/><span data-ttu-id="1a9a9-162">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-162">Default value is `false`.</span></span>  |  <span data-ttu-id="1a9a9-163">Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-163">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="1a9a9-164">As funções de cancelamento normalmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-164">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="1a9a9-165">Uma função não pode ser streaming e cancelamento.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-165">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="1a9a9-166">Para obter mais informações, consulte a observação próxima ao final de [fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-166">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="1a9a9-167">booliano</span><span class="sxs-lookup"><span data-stu-id="1a9a9-167">boolean</span></span> | <span data-ttu-id="1a9a9-168">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-168">No</span></span> <br/><br/><span data-ttu-id="1a9a9-169">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="1a9a9-170">Se true, sua função personalizada pode acessar o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="1a9a9-171">Para obter o endereço da célula que chamou sua função personalizada, use Context. Address em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="1a9a9-172">Para saber mais, confira [determinar quais célula chamada sua função personalizada](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="1a9a9-173">As funções personalizadas não podem ser definidas como streaming e requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="1a9a9-174">Ao usar essa opção, o parâmetro "invocar" deve ser o último parâmetro passado em opções.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-174">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="1a9a9-175">booliano</span><span class="sxs-lookup"><span data-stu-id="1a9a9-175">boolean</span></span>  |  <span data-ttu-id="1a9a9-176">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-176">No</span></span><br/><br/><span data-ttu-id="1a9a9-177">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-177">Default value is `false`.</span></span>  |  <span data-ttu-id="1a9a9-178">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="1a9a9-179">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="1a9a9-180">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-180">The function should have no `return` statement.</span></span> <span data-ttu-id="1a9a9-181">Em vez disso, o valor resultante é passado como o argumento do método de retorno `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-181">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="1a9a9-182">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-182">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="1a9a9-183">booliano</span><span class="sxs-lookup"><span data-stu-id="1a9a9-183">boolean</span></span> | <span data-ttu-id="1a9a9-184">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-184">No</span></span> <br/><br/><span data-ttu-id="1a9a9-185">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-185">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="1a9a9-186">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-186">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="1a9a9-187">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-187">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="1a9a9-188">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-188">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="1a9a9-189">parâmetros</span><span class="sxs-lookup"><span data-stu-id="1a9a9-189">parameters</span></span>

<span data-ttu-id="1a9a9-190">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-190">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="1a9a9-191">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-191">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="1a9a9-192">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1a9a9-192">Property</span></span>  |  <span data-ttu-id="1a9a9-193">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1a9a9-193">Data type</span></span>  |  <span data-ttu-id="1a9a9-194">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1a9a9-194">Required</span></span>  |  <span data-ttu-id="1a9a9-195">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a9a9-195">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="1a9a9-196">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-196">string</span></span>  |  <span data-ttu-id="1a9a9-197">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-197">No</span></span> |  <span data-ttu-id="1a9a9-198">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-198">A description of the parameter.</span></span> <span data-ttu-id="1a9a9-199">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-199">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="1a9a9-200">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-200">string</span></span>  |  <span data-ttu-id="1a9a9-201">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-201">No</span></span>  |  <span data-ttu-id="1a9a9-202">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-202">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="1a9a9-203">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-203">string</span></span>  |  <span data-ttu-id="1a9a9-204">Sim</span><span class="sxs-lookup"><span data-stu-id="1a9a9-204">Yes</span></span>  |  <span data-ttu-id="1a9a9-205">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-205">The name of the parameter.</span></span> <span data-ttu-id="1a9a9-206">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-206">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="1a9a9-207">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-207">string</span></span>  |  <span data-ttu-id="1a9a9-208">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-208">No</span></span>  |  <span data-ttu-id="1a9a9-209">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-209">The data type of the parameter.</span></span> <span data-ttu-id="1a9a9-210">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-210">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="1a9a9-211">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-211">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="1a9a9-212">booliano</span><span class="sxs-lookup"><span data-stu-id="1a9a9-212">boolean</span></span> | <span data-ttu-id="1a9a9-213">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-213">No</span></span> | <span data-ttu-id="1a9a9-214">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-214">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="1a9a9-215">result</span><span class="sxs-lookup"><span data-stu-id="1a9a9-215">result</span></span>

<span data-ttu-id="1a9a9-216">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-216">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1a9a9-217">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-217">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="1a9a9-218">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1a9a9-218">Property</span></span>  |  <span data-ttu-id="1a9a9-219">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="1a9a9-219">Data type</span></span>  |  <span data-ttu-id="1a9a9-220">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1a9a9-220">Required</span></span>  |  <span data-ttu-id="1a9a9-221">Descrição</span><span class="sxs-lookup"><span data-stu-id="1a9a9-221">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="1a9a9-222">string</span><span class="sxs-lookup"><span data-stu-id="1a9a9-222">string</span></span>  |  <span data-ttu-id="1a9a9-223">Não</span><span class="sxs-lookup"><span data-stu-id="1a9a9-223">No</span></span>  |  <span data-ttu-id="1a9a9-224">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-224">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="1a9a9-225">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="1a9a9-225">Next steps</span></span>
<span data-ttu-id="1a9a9-226">Conheça as [práticas recomendadas para nomear sua função](custom-functions-naming.md) ou descubra como [localizar sua função](custom-functions-localize.md) usando o método JSON manuscrito descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="1a9a9-226">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="1a9a9-227">Confira também</span><span class="sxs-lookup"><span data-stu-id="1a9a9-227">See also</span></span>

* [<span data-ttu-id="1a9a9-228">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1a9a9-228">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="1a9a9-229">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1a9a9-229">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* <span data-ttu-id="1a9a9-230">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="1a9a9-230">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="1a9a9-231">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="1a9a9-231">Create custom functions in Excel</span></span>](custom-functions-overview.md)