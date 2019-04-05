---
ms.date: 03/29/2019
description: Defina os metadados de funções personalizadas no Excel.
title: Metadados de funções personalizadas no Excel (visualização)
localization_priority: Normal
ms.openlocfilehash: 28a9a0207f7439af164eb9ca7c4b9ed9e966b3ed
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477548"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="55ace-103">Metadados de funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="55ace-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="55ace-104">Quando você define [funções personalizadas](custom-functions-overview.md) dentro de seu suplemento do Excel, o projeto do suplemento inclui um arquivo de metadados JSON que fornece as informações que o Excel requer para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="55ace-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="55ace-105">Este arquivo é gerado:</span><span class="sxs-lookup"><span data-stu-id="55ace-105">This file is generated either:</span></span>

- <span data-ttu-id="55ace-106">por você, em um arquivo JSON manuscrito</span><span class="sxs-lookup"><span data-stu-id="55ace-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="55ace-107">nos comentários do JSDoc inseridos no início da função</span><span class="sxs-lookup"><span data-stu-id="55ace-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="55ace-108">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="55ace-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="55ace-109">Este artigo descreve o formato do arquivo de metadados JSON, supondo que você o esteja escrevendo à mão.</span><span class="sxs-lookup"><span data-stu-id="55ace-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="55ace-110">Para obter informações sobre a geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="55ace-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="55ace-111">Para saber mais sobre outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="55ace-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="55ace-112">Configurações do servidor no servidor que hospeda o arquivo JSON deve ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para funções personalizadas funcionarem corretamente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="55ace-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="55ace-113">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="55ace-113">Example metadata</span></span>

<span data-ttu-id="55ace-114">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="55ace-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="55ace-115">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="55ace-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="55ace-116">Um exemplo de arquivo JSON completo está disponível no repositório GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="55ace-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="55ace-117">functions</span><span class="sxs-lookup"><span data-stu-id="55ace-117">functions</span></span> 

<span data-ttu-id="55ace-118">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="55ace-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="55ace-119">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="55ace-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="55ace-120">Propriedade</span><span class="sxs-lookup"><span data-stu-id="55ace-120">Property</span></span>  |  <span data-ttu-id="55ace-121">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="55ace-121">Data type</span></span>  |  <span data-ttu-id="55ace-122">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="55ace-122">Required</span></span>  |  <span data-ttu-id="55ace-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="55ace-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="55ace-124">string</span><span class="sxs-lookup"><span data-stu-id="55ace-124">string</span></span>  |  <span data-ttu-id="55ace-125">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-125">No</span></span>  |  <span data-ttu-id="55ace-126">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="55ace-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="55ace-127">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="55ace-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="55ace-128">string</span><span class="sxs-lookup"><span data-stu-id="55ace-128">string</span></span>  |   <span data-ttu-id="55ace-129">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-129">No</span></span>  |  <span data-ttu-id="55ace-130">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="55ace-130">URL that provides information about the function.</span></span> <span data-ttu-id="55ace-131">(Ela é exibida em um painel de tarefas). Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="55ace-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="55ace-132">string</span><span class="sxs-lookup"><span data-stu-id="55ace-132">string</span></span> | <span data-ttu-id="55ace-133">Sim</span><span class="sxs-lookup"><span data-stu-id="55ace-133">Yes</span></span> | <span data-ttu-id="55ace-134">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="55ace-134">A unique ID for the function.</span></span> <span data-ttu-id="55ace-135">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="55ace-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="55ace-136">string</span><span class="sxs-lookup"><span data-stu-id="55ace-136">string</span></span>  |  <span data-ttu-id="55ace-137">Sim</span><span class="sxs-lookup"><span data-stu-id="55ace-137">Yes</span></span>  |  <span data-ttu-id="55ace-138">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="55ace-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="55ace-139">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="55ace-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="55ace-140">objeto</span><span class="sxs-lookup"><span data-stu-id="55ace-140">object</span></span>  |  <span data-ttu-id="55ace-141">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-141">No</span></span>  |  <span data-ttu-id="55ace-142">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="55ace-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="55ace-143">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="55ace-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="55ace-144">array</span><span class="sxs-lookup"><span data-stu-id="55ace-144">array</span></span>  |  <span data-ttu-id="55ace-145">Sim</span><span class="sxs-lookup"><span data-stu-id="55ace-145">Yes</span></span>  |  <span data-ttu-id="55ace-146">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="55ace-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="55ace-147">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="55ace-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="55ace-148">object</span><span class="sxs-lookup"><span data-stu-id="55ace-148">object</span></span>  |  <span data-ttu-id="55ace-149">Sim</span><span class="sxs-lookup"><span data-stu-id="55ace-149">Yes</span></span>  |  <span data-ttu-id="55ace-150">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="55ace-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="55ace-151">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="55ace-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="55ace-152">options</span><span class="sxs-lookup"><span data-stu-id="55ace-152">options</span></span>

<span data-ttu-id="55ace-153">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="55ace-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="55ace-154">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="55ace-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="55ace-155">Propriedade</span><span class="sxs-lookup"><span data-stu-id="55ace-155">Property</span></span>  |  <span data-ttu-id="55ace-156">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="55ace-156">Data type</span></span>  |  <span data-ttu-id="55ace-157">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="55ace-157">Required</span></span>  |  <span data-ttu-id="55ace-158">Descrição</span><span class="sxs-lookup"><span data-stu-id="55ace-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="55ace-159">booliano</span><span class="sxs-lookup"><span data-stu-id="55ace-159">boolean</span></span>  |  <span data-ttu-id="55ace-160">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-160">No</span></span><br/><br/><span data-ttu-id="55ace-161">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="55ace-161">Default value is `false`.</span></span>  |  <span data-ttu-id="55ace-162">Se o valor for `true`, o Excel chamará o manipulador `onCanceled` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="55ace-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="55ace-163">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="55ace-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="55ace-164">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="55ace-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="55ace-165">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="55ace-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="55ace-166">Para saber mais, confira [Cancelar uma função](custom-functions-web-reqs.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="55ace-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="55ace-167">booliano</span><span class="sxs-lookup"><span data-stu-id="55ace-167">boolean</span></span>  |  <span data-ttu-id="55ace-168">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-168">No</span></span><br/><br/><span data-ttu-id="55ace-169">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="55ace-169">Default value is `false`.</span></span>  |  <span data-ttu-id="55ace-170">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="55ace-170">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="55ace-171">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="55ace-171">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="55ace-172">Se você usar essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="55ace-172">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="55ace-173">(***Não*** registre este parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="55ace-173">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="55ace-174">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="55ace-174">The function should have no `return` statement.</span></span> <span data-ttu-id="55ace-175">Em vez disso, o valor resultante é passado como o argumento do método de retorno `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="55ace-175">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="55ace-176">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="55ace-176">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="55ace-177">booliano</span><span class="sxs-lookup"><span data-stu-id="55ace-177">boolean</span></span> | <span data-ttu-id="55ace-178">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-178">No</span></span> <br/><br/><span data-ttu-id="55ace-179">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="55ace-179">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="55ace-180">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="55ace-180">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="55ace-181">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="55ace-181">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="55ace-182">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="55ace-182">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="55ace-183">parâmetros</span><span class="sxs-lookup"><span data-stu-id="55ace-183">parameters</span></span>

<span data-ttu-id="55ace-184">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="55ace-184">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="55ace-185">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="55ace-185">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="55ace-186">Propriedade</span><span class="sxs-lookup"><span data-stu-id="55ace-186">Property</span></span>  |  <span data-ttu-id="55ace-187">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="55ace-187">Data type</span></span>  |  <span data-ttu-id="55ace-188">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="55ace-188">Required</span></span>  |  <span data-ttu-id="55ace-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="55ace-189">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="55ace-190">string</span><span class="sxs-lookup"><span data-stu-id="55ace-190">string</span></span>  |  <span data-ttu-id="55ace-191">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-191">No</span></span> |  <span data-ttu-id="55ace-192">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="55ace-192">A description of the parameter.</span></span> <span data-ttu-id="55ace-193">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="55ace-193">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="55ace-194">string</span><span class="sxs-lookup"><span data-stu-id="55ace-194">string</span></span>  |  <span data-ttu-id="55ace-195">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-195">No</span></span>  |  <span data-ttu-id="55ace-196">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="55ace-196">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="55ace-197">string</span><span class="sxs-lookup"><span data-stu-id="55ace-197">string</span></span>  |  <span data-ttu-id="55ace-198">Sim</span><span class="sxs-lookup"><span data-stu-id="55ace-198">Yes</span></span>  |  <span data-ttu-id="55ace-199">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="55ace-199">The name of the parameter.</span></span> <span data-ttu-id="55ace-200">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="55ace-200">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="55ace-201">string</span><span class="sxs-lookup"><span data-stu-id="55ace-201">string</span></span>  |  <span data-ttu-id="55ace-202">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-202">No</span></span>  |  <span data-ttu-id="55ace-203">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="55ace-203">The data type of the parameter.</span></span> <span data-ttu-id="55ace-204">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="55ace-204">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="55ace-205">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="55ace-205">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="55ace-206">booliano</span><span class="sxs-lookup"><span data-stu-id="55ace-206">boolean</span></span> | <span data-ttu-id="55ace-207">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-207">No</span></span> | <span data-ttu-id="55ace-208">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="55ace-208">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="55ace-209">Se a propriedade `type` de um parâmetro opcional não for especificada ou definida como `any`, é provável que você tenha problemas, como erros de lint em seu IDE e parâmetros opcionais que não serão exibidos quando a função estiver sendo inserida em uma célula no Excel.</span><span class="sxs-lookup"><span data-stu-id="55ace-209">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="55ace-210">A previsão é para ser alterado em dezembro de 2018.</span><span class="sxs-lookup"><span data-stu-id="55ace-210">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="55ace-211">result</span><span class="sxs-lookup"><span data-stu-id="55ace-211">result</span></span>

<span data-ttu-id="55ace-212">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="55ace-212">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="55ace-213">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="55ace-213">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="55ace-214">Propriedade</span><span class="sxs-lookup"><span data-stu-id="55ace-214">Property</span></span>  |  <span data-ttu-id="55ace-215">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="55ace-215">Data type</span></span>  |  <span data-ttu-id="55ace-216">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="55ace-216">Required</span></span>  |  <span data-ttu-id="55ace-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="55ace-217">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="55ace-218">string</span><span class="sxs-lookup"><span data-stu-id="55ace-218">string</span></span>  |  <span data-ttu-id="55ace-219">Não</span><span class="sxs-lookup"><span data-stu-id="55ace-219">No</span></span>  |  <span data-ttu-id="55ace-220">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="55ace-220">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="55ace-221">string</span><span class="sxs-lookup"><span data-stu-id="55ace-221">string</span></span>  |  <span data-ttu-id="55ace-222">Sim</span><span class="sxs-lookup"><span data-stu-id="55ace-222">Yes</span></span>  |  <span data-ttu-id="55ace-223">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="55ace-223">The data type of the parameter.</span></span> <span data-ttu-id="55ace-224">Deve ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="55ace-224">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="55ace-225">Confira também</span><span class="sxs-lookup"><span data-stu-id="55ace-225">See also</span></span>

* [<span data-ttu-id="55ace-226">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="55ace-226">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="55ace-227">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="55ace-227">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="55ace-228">Práticas recomendadas de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="55ace-228">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="55ace-229">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="55ace-229">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="55ace-230">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="55ace-230">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
