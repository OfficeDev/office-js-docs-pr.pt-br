---
ms.date: 09/27/2018
description: Defina metadados para funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
ms.openlocfilehash: e8af13b8855d6c5e1a3b1ce99edb24445e066756
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459235"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="7a2ae-103">Metadados de funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="7a2ae-103">Custom functions metadata</span></span>

<span data-ttu-id="7a2ae-104">Quando você define [funções personalizadas](custom-functions-overview.md) no seu suplemento do Excel, seu projeto de suplemento deve incluir um arquivo de metadados JSON que fornece as informações que o Excel precisa para registrar as funções personalizadas e torná-las disponíveis para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="7a2ae-105">Este artigo descreve o formato do arquivo JSON de metadados.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="7a2ae-106">Para obter informações sobre os outros arquivos que você deve incluir em seu projeto de suplemento para habilitar funções personalizadas, confira [Criação de funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="7a2ae-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="7a2ae-107">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="7a2ae-107">Example metadata</span></span>

<span data-ttu-id="7a2ae-p102">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas. As seções a seguir neste exemplo fornecem informações detalhadas sobre as propriedades individuais nesse exemplo JSON.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="7a2ae-110">Um exemplo completo do arquivo JSON está disponível no [repositório GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="7a2ae-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="7a2ae-111">functions</span><span class="sxs-lookup"><span data-stu-id="7a2ae-111">functions</span></span> 

<span data-ttu-id="7a2ae-p103">A `functions` propriedade é uma matriz de objetos de função personalizada. A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="7a2ae-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="7a2ae-114">Property</span></span>  |  <span data-ttu-id="7a2ae-115">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="7a2ae-115">Data type</span></span>  |  <span data-ttu-id="7a2ae-116">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7a2ae-116">Required</span></span>  |  <span data-ttu-id="7a2ae-117">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a2ae-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="7a2ae-118">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-118">string</span></span>  |  <span data-ttu-id="7a2ae-119">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-119">No</span></span>  |  <span data-ttu-id="7a2ae-120">O nome da função que os usuários finais veem no Excel.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="7a2ae-121">Por exemplo, **Converte um valor Celsius em Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="7a2ae-122">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-122">string</span></span>  |   <span data-ttu-id="7a2ae-123">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-123">No</span></span>  |  <span data-ttu-id="7a2ae-124">URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-124">URL that provides information about the function.</span></span> <span data-ttu-id="7a2ae-125">(É exibida em um painel de tarefas.) Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="7a2ae-126">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-126">string</span></span> | <span data-ttu-id="7a2ae-127">Sim</span><span class="sxs-lookup"><span data-stu-id="7a2ae-127">Yes</span></span> | <span data-ttu-id="7a2ae-p106">Um ID exclusivo para a função. Esse ID não deve ser alterado depois de definido.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p106">A unique ID for the function. This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="7a2ae-130">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-130">string</span></span>  |  <span data-ttu-id="7a2ae-131">Sim</span><span class="sxs-lookup"><span data-stu-id="7a2ae-131">Yes</span></span>  |  <span data-ttu-id="7a2ae-p107">O nome da função que os usuários finais veem no Excel. No Excel, esse nome de função será prefixado pelo namespace das funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="7a2ae-134">objeto</span><span class="sxs-lookup"><span data-stu-id="7a2ae-134">object</span></span>  |  <span data-ttu-id="7a2ae-135">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-135">No</span></span>  |  <span data-ttu-id="7a2ae-p108">Permite personalizar alguns aspectos de como e quando o Excel executa a função. Consulte o [objeto options](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="7a2ae-138">matriz</span><span class="sxs-lookup"><span data-stu-id="7a2ae-138">array</span></span>  |  <span data-ttu-id="7a2ae-139">Sim</span><span class="sxs-lookup"><span data-stu-id="7a2ae-139">Yes</span></span>  |  <span data-ttu-id="7a2ae-p109">Matriz que define os parâmetros de entrada para a função. Consulte a [matriz de parâmetros](#parameters-array) , para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="7a2ae-142">objeto</span><span class="sxs-lookup"><span data-stu-id="7a2ae-142">object</span></span>  |  <span data-ttu-id="7a2ae-143">Sim</span><span class="sxs-lookup"><span data-stu-id="7a2ae-143">Yes</span></span>  |  <span data-ttu-id="7a2ae-p110">Objeto que define o tipo de informação que é retornado pela função. Consulte o [objeto result](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="7a2ae-146">options</span><span class="sxs-lookup"><span data-stu-id="7a2ae-146">options</span></span>

<span data-ttu-id="7a2ae-p111">O objeto  `options` permite personalizar alguns aspectos do como e quando o Excel executa a função. A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="7a2ae-149">Propriedade</span><span class="sxs-lookup"><span data-stu-id="7a2ae-149">Property</span></span>  |  <span data-ttu-id="7a2ae-150">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="7a2ae-150">Data type</span></span>  |  <span data-ttu-id="7a2ae-151">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7a2ae-151">Required</span></span>  |  <span data-ttu-id="7a2ae-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a2ae-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="7a2ae-153">booleano</span><span class="sxs-lookup"><span data-stu-id="7a2ae-153">boolean</span></span>  |  <span data-ttu-id="7a2ae-154">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-154">No</span></span><br/><br/><span data-ttu-id="7a2ae-155">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-155">Default value is 4.</span></span>  |  <span data-ttu-id="7a2ae-p112">Se `true`, o Excel chama o manipulador de `onCanceled` sempre que o usuário realizar uma ação que tem o efeito de cancelar a função; por exemplo, disparando manualmente o recálculo ou editando uma célula referenciada pela função. Se você usar essa opção, o Excel chamará a função JavaScript com o parâmetro adicional `caller`. (***Não*** registre esse parâmetro na propriedade `parameters`). No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`. Para saber mais, confira [Cancelar uma função](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="7a2ae-161">booleano</span><span class="sxs-lookup"><span data-stu-id="7a2ae-161">boolean</span></span>  |  <span data-ttu-id="7a2ae-162">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-162">No</span></span><br/><br/><span data-ttu-id="7a2ae-163">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-163">Default value is 4.</span></span>  |  <span data-ttu-id="7a2ae-p113">Se `true`, a função pode atualizar o valor da célula repetidamente, mesmo quando invocada apenas uma vez. Essa opção é útil para alterar fontes de dados rapidamente, como um preço de ação. Se você usar essa opção, o Excel chamará a função JavaScript com o parâmetro adicional `caller`. (***Não*** registre esse parâmetro na propriedade `parameters` ). A função não deve ter a instrução `return`. Em vez disso, o valor de resultado é passado como o argumento do `caller.setResult` método de retorno de chamada. Para obter mais informações, consulte [Funções de fluxo contínuo](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p113">If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streamed functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="7a2ae-171">parameters</span><span class="sxs-lookup"><span data-stu-id="7a2ae-171">parameters</span></span>

<span data-ttu-id="7a2ae-p114">A propriedade  `parameters` é uma matriz de parâmetros. A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="7a2ae-174">Propriedade</span><span class="sxs-lookup"><span data-stu-id="7a2ae-174">Property</span></span>  |  <span data-ttu-id="7a2ae-175">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="7a2ae-175">Data type</span></span>  |  <span data-ttu-id="7a2ae-176">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7a2ae-176">Required</span></span>  |  <span data-ttu-id="7a2ae-177">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a2ae-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="7a2ae-178">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-178">string</span></span>  |  <span data-ttu-id="7a2ae-179">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-179">No</span></span> |  <span data-ttu-id="7a2ae-180">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="7a2ae-181">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-181">string</span></span>  |  <span data-ttu-id="7a2ae-182">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-182">No</span></span>  |  <span data-ttu-id="7a2ae-183">Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="7a2ae-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="7a2ae-184">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-184">string</span></span>  |  <span data-ttu-id="7a2ae-185">Sim</span><span class="sxs-lookup"><span data-stu-id="7a2ae-185">Yes</span></span>  |  <span data-ttu-id="7a2ae-p115">O nome do parâmetro. Esse nome é exibido no intelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="7a2ae-188">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-188">string</span></span>  |  <span data-ttu-id="7a2ae-189">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-189">No</span></span>  |  <span data-ttu-id="7a2ae-p116">O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="7a2ae-192">result</span><span class="sxs-lookup"><span data-stu-id="7a2ae-192">result</span></span>

<span data-ttu-id="7a2ae-p117">O objeto  `results` define o tipo de informação retornado pela função. A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="7a2ae-195">Propriedade</span><span class="sxs-lookup"><span data-stu-id="7a2ae-195">Property</span></span>  |  <span data-ttu-id="7a2ae-196">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="7a2ae-196">Data type</span></span>  |  <span data-ttu-id="7a2ae-197">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="7a2ae-197">Required</span></span>  |  <span data-ttu-id="7a2ae-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a2ae-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="7a2ae-199">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-199">string</span></span>  |  <span data-ttu-id="7a2ae-200">Não</span><span class="sxs-lookup"><span data-stu-id="7a2ae-200">No</span></span>  |  <span data-ttu-id="7a2ae-201">Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="7a2ae-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="7a2ae-202">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a2ae-202">string</span></span>  |  <span data-ttu-id="7a2ae-203">Sim</span><span class="sxs-lookup"><span data-stu-id="7a2ae-203">Yes</span></span>  |  <span data-ttu-id="7a2ae-p118">O tipo de dados do parâmetro. Deve ser **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="7a2ae-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="7a2ae-206">Confira também</span><span class="sxs-lookup"><span data-stu-id="7a2ae-206">See also</span></span>

* [<span data-ttu-id="7a2ae-207">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="7a2ae-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="7a2ae-208">Runtime para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="7a2ae-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="7a2ae-209">Práticas recomendadas para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="7a2ae-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="7a2ae-210">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="7a2ae-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)