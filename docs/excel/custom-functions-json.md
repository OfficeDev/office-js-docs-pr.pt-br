---
ms.date: 05/03/2019
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: 92e2b1aaae46d376cc8033b304192d7ce8489fd8
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628071"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="0a10a-103">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0a10a-103">Custom functions metadata</span></span>

<span data-ttu-id="0a10a-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o projeto do suplemento inclui um arquivo de metadados JSON que fornece as informações que o Excel requer para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="0a10a-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0a10a-105">Este arquivo é gerado:</span><span class="sxs-lookup"><span data-stu-id="0a10a-105">This file is generated either:</span></span>

- <span data-ttu-id="0a10a-106">Por você, em um arquivo JSON manuscrito</span><span class="sxs-lookup"><span data-stu-id="0a10a-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="0a10a-107">Nos comentários do JSDoc inseridos no início da função</span><span class="sxs-lookup"><span data-stu-id="0a10a-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="0a10a-108">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="0a10a-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="0a10a-109">Este artigo descreve o formato do arquivo de metadados JSON, supondo que você o esteja escrevendo à mão.</span><span class="sxs-lookup"><span data-stu-id="0a10a-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="0a10a-110">Para obter informações sobre a geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="0a10a-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="0a10a-111">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="0a10a-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="0a10a-112">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="0a10a-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="0a10a-113">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="0a10a-113">Example metadata</span></span>

<span data-ttu-id="0a10a-114">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0a10a-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="0a10a-115">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="0a10a-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="0a10a-116">Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="0a10a-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="0a10a-117">functions</span><span class="sxs-lookup"><span data-stu-id="0a10a-117">functions</span></span> 

<span data-ttu-id="0a10a-118">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0a10a-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="0a10a-119">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="0a10a-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0a10a-120">Propriedade</span><span class="sxs-lookup"><span data-stu-id="0a10a-120">Property</span></span>  |  <span data-ttu-id="0a10a-121">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="0a10a-121">Data type</span></span>  |  <span data-ttu-id="0a10a-122">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0a10a-122">Required</span></span>  |  <span data-ttu-id="0a10a-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="0a10a-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0a10a-124">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-124">string</span></span>  |  <span data-ttu-id="0a10a-125">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-125">No</span></span>  |  <span data-ttu-id="0a10a-126">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="0a10a-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="0a10a-127">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="0a10a-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="0a10a-128">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="0a10a-128">string</span></span>  |   <span data-ttu-id="0a10a-129">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-129">No</span></span>  |  <span data-ttu-id="0a10a-130">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-130">URL that provides information about the function.</span></span> <span data-ttu-id="0a10a-131">(Ela é exibida em um painel de tarefas). Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="0a10a-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="0a10a-132">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-132">string</span></span> | <span data-ttu-id="0a10a-133">Sim</span><span class="sxs-lookup"><span data-stu-id="0a10a-133">Yes</span></span> | <span data-ttu-id="0a10a-134">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-134">A unique ID for the function.</span></span> <span data-ttu-id="0a10a-135">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="0a10a-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="0a10a-136">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-136">string</span></span>  |  <span data-ttu-id="0a10a-137">Sim</span><span class="sxs-lookup"><span data-stu-id="0a10a-137">Yes</span></span>  |  <span data-ttu-id="0a10a-138">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="0a10a-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="0a10a-139">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="0a10a-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="0a10a-140">objeto</span><span class="sxs-lookup"><span data-stu-id="0a10a-140">object</span></span>  |  <span data-ttu-id="0a10a-141">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-141">No</span></span>  |  <span data-ttu-id="0a10a-142">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0a10a-143">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="0a10a-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="0a10a-144">array</span><span class="sxs-lookup"><span data-stu-id="0a10a-144">array</span></span>  |  <span data-ttu-id="0a10a-145">Sim</span><span class="sxs-lookup"><span data-stu-id="0a10a-145">Yes</span></span>  |  <span data-ttu-id="0a10a-146">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="0a10a-147">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="0a10a-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="0a10a-148">object</span><span class="sxs-lookup"><span data-stu-id="0a10a-148">object</span></span>  |  <span data-ttu-id="0a10a-149">Sim</span><span class="sxs-lookup"><span data-stu-id="0a10a-149">Yes</span></span>  |  <span data-ttu-id="0a10a-150">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="0a10a-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0a10a-151">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="0a10a-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="0a10a-152">options</span><span class="sxs-lookup"><span data-stu-id="0a10a-152">options</span></span>

<span data-ttu-id="0a10a-153">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0a10a-154">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="0a10a-155">Propriedade</span><span class="sxs-lookup"><span data-stu-id="0a10a-155">Property</span></span>  |  <span data-ttu-id="0a10a-156">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="0a10a-156">Data type</span></span>  |  <span data-ttu-id="0a10a-157">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0a10a-157">Required</span></span>  |  <span data-ttu-id="0a10a-158">Descrição</span><span class="sxs-lookup"><span data-stu-id="0a10a-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="0a10a-159">booliano</span><span class="sxs-lookup"><span data-stu-id="0a10a-159">boolean</span></span>  |  <span data-ttu-id="0a10a-160">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-160">No</span></span><br/><br/><span data-ttu-id="0a10a-161">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-161">Default value is `false`.</span></span>  |  <span data-ttu-id="0a10a-162">Se o valor for `true`, o Excel chamará o manipulador `onCanceled` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="0a10a-163">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="0a10a-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="0a10a-164">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="0a10a-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="0a10a-165">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="0a10a-166">Para saber mais, confira [Cancelar uma função](custom-functions-web-reqs.md#stream-and-cancel-functions).</span><span class="sxs-lookup"><span data-stu-id="0a10a-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#stream-and-cancel-functions).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="0a10a-167">booliano</span><span class="sxs-lookup"><span data-stu-id="0a10a-167">boolean</span></span> | <span data-ttu-id="0a10a-168">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-168">No</span></span> <br/><br/><span data-ttu-id="0a10a-169">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0a10a-170">Se true, sua função personalizada pode acessar o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0a10a-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="0a10a-171">Para obter o endereço da célula que chamou sua função personalizada, use Context. Address em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0a10a-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="0a10a-172">Para saber mais, confira [determinar quais célula chamada sua função personalizada](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span><span class="sxs-lookup"><span data-stu-id="0a10a-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="0a10a-173">As funções personalizadas não podem ser definidas como streaming e requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="0a10a-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="0a10a-174">Ao usar essa opção, o parâmetro "invocationContext" deve ser o último parâmetro passado em opções.</span><span class="sxs-lookup"><span data-stu-id="0a10a-174">When using this option, the 'invocationContext' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="0a10a-175">booliano</span><span class="sxs-lookup"><span data-stu-id="0a10a-175">boolean</span></span>  |  <span data-ttu-id="0a10a-176">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-176">No</span></span><br/><br/><span data-ttu-id="0a10a-177">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-177">Default value is `false`.</span></span>  |  <span data-ttu-id="0a10a-178">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="0a10a-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="0a10a-179">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="0a10a-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="0a10a-180">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="0a10a-180">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="0a10a-181">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="0a10a-181">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="0a10a-182">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-182">The function should have no `return` statement.</span></span> <span data-ttu-id="0a10a-183">Em vez disso, o valor resultante é passado como o argumento do método de retorno `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-183">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="0a10a-184">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#stream-and-cancel-functions).</span><span class="sxs-lookup"><span data-stu-id="0a10a-184">For more information, see [Streaming functions](custom-functions-web-reqs.md#stream-and-cancel-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="0a10a-185">booliano</span><span class="sxs-lookup"><span data-stu-id="0a10a-185">boolean</span></span> | <span data-ttu-id="0a10a-186">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-186">No</span></span> <br/><br/><span data-ttu-id="0a10a-187">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-187">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0a10a-188">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="0a10a-188">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="0a10a-189">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="0a10a-189">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="0a10a-190">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="0a10a-190">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="0a10a-191">parâmetros</span><span class="sxs-lookup"><span data-stu-id="0a10a-191">parameters</span></span>

<span data-ttu-id="0a10a-192">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="0a10a-192">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="0a10a-193">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="0a10a-193">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0a10a-194">Propriedade</span><span class="sxs-lookup"><span data-stu-id="0a10a-194">Property</span></span>  |  <span data-ttu-id="0a10a-195">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="0a10a-195">Data type</span></span>  |  <span data-ttu-id="0a10a-196">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0a10a-196">Required</span></span>  |  <span data-ttu-id="0a10a-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="0a10a-197">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0a10a-198">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-198">string</span></span>  |  <span data-ttu-id="0a10a-199">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-199">No</span></span> |  <span data-ttu-id="0a10a-200">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="0a10a-200">A description of the parameter.</span></span> <span data-ttu-id="0a10a-201">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="0a10a-201">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="0a10a-202">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="0a10a-202">string</span></span>  |  <span data-ttu-id="0a10a-203">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-203">No</span></span>  |  <span data-ttu-id="0a10a-204">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="0a10a-204">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="0a10a-205">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-205">string</span></span>  |  <span data-ttu-id="0a10a-206">Sim</span><span class="sxs-lookup"><span data-stu-id="0a10a-206">Yes</span></span>  |  <span data-ttu-id="0a10a-207">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="0a10a-207">The name of the parameter.</span></span> <span data-ttu-id="0a10a-208">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="0a10a-208">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="0a10a-209">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-209">string</span></span>  |  <span data-ttu-id="0a10a-210">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-210">No</span></span>  |  <span data-ttu-id="0a10a-211">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="0a10a-211">The data type of the parameter.</span></span> <span data-ttu-id="0a10a-212">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="0a10a-212">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="0a10a-213">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="0a10a-213">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="0a10a-214">booliano</span><span class="sxs-lookup"><span data-stu-id="0a10a-214">boolean</span></span> | <span data-ttu-id="0a10a-215">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-215">No</span></span> | <span data-ttu-id="0a10a-216">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="0a10a-216">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="0a10a-217">result</span><span class="sxs-lookup"><span data-stu-id="0a10a-217">result</span></span>

<span data-ttu-id="0a10a-218">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="0a10a-218">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0a10a-219">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="0a10a-219">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="0a10a-220">Propriedade</span><span class="sxs-lookup"><span data-stu-id="0a10a-220">Property</span></span>  |  <span data-ttu-id="0a10a-221">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="0a10a-221">Data type</span></span>  |  <span data-ttu-id="0a10a-222">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0a10a-222">Required</span></span>  |  <span data-ttu-id="0a10a-223">Descrição</span><span class="sxs-lookup"><span data-stu-id="0a10a-223">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="0a10a-224">string</span><span class="sxs-lookup"><span data-stu-id="0a10a-224">string</span></span>  |  <span data-ttu-id="0a10a-225">Não</span><span class="sxs-lookup"><span data-stu-id="0a10a-225">No</span></span>  |  <span data-ttu-id="0a10a-226">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="0a10a-226">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="0a10a-227">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="0a10a-227">Next steps</span></span>
<span data-ttu-id="0a10a-228">Conheça as [práticas recomendadas para nomear sua função](custom-functions-naming.md) ou descubra como [localizar sua função](custom-functions-localize.md) usando o método JSON manuscrito descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="0a10a-228">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="0a10a-229">Confira também</span><span class="sxs-lookup"><span data-stu-id="0a10a-229">See also</span></span>

* [<span data-ttu-id="0a10a-230">Gerar automaticamente metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0a10a-230">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="0a10a-231">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0a10a-231">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* <span data-ttu-id="0a10a-232">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="0a10a-232">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="0a10a-233">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="0a10a-233">Create custom functions in Excel</span></span>](custom-functions-overview.md)