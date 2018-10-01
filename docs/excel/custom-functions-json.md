---
ms.date: 09/27/2018
description: Defina metadados para funções personalizadas no Excel.
title: Metadados para funções personalizadas no Excel
ms.openlocfilehash: a179a9c4bc071200cab1377c5e48913bfc8358cf
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348791"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="26df2-103">Metadados de funções personalizadas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="26df2-103">Custom functions metadata</span></span>

<span data-ttu-id="26df2-p101">Quando você define [funções personalizadas](custom-functions-overview.md) no suplemento do Excel, o projeto de suplemento deve incluir um arquivo de metadados JSON que forneça as informações necessárias para o Excel registrar as funções personalizadas e disponibilizá-las aos usuários finais. Este artigo descreve o formato do arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="26df2-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="26df2-106">Para obter informações sobre os outros arquivos que você deve incluir no projeto de suplemento para habilitar funções personalizadas, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="26df2-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="26df2-107">Exemplo de metadados</span><span class="sxs-lookup"><span data-stu-id="26df2-107">Example metadata</span></span>

<span data-ttu-id="26df2-108">O exemplo a seguir mostra o conteúdo de um arquivo JSON de metadados para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="26df2-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="26df2-109">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais nesse exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="26df2-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="26df2-110">Um exemplo completo de arquivo JSON está disponível no [repositório GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="26df2-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="26df2-111">functions</span><span class="sxs-lookup"><span data-stu-id="26df2-111">functions</span></span> 

<span data-ttu-id="26df2-112">A propriedade `functions` é uma matriz de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="26df2-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="26df2-113">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="26df2-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="26df2-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="26df2-114">Property</span></span>  |  <span data-ttu-id="26df2-115">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="26df2-115">Data type</span></span>  |  <span data-ttu-id="26df2-116">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="26df2-116">Required</span></span>  |  <span data-ttu-id="26df2-117">Descrição</span><span class="sxs-lookup"><span data-stu-id="26df2-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="26df2-118">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-118">string</span></span>  |  <span data-ttu-id="26df2-119">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-119">No</span></span>  |  <span data-ttu-id="26df2-p104">A descrição da função que os usuários finais veem no Excel. Por exemplo, **Converte um valor de Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="26df2-p104">A description of the function that appears in the Excel UI. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="26df2-122">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-122">string</span></span>  |   <span data-ttu-id="26df2-123">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-123">No</span></span>  |  <span data-ttu-id="26df2-p105">URL que fornece informações sobre a função. (Ela é exibida em um painel de tarefas.) Por exemplo, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="26df2-p105">URL where users can get information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="26df2-126">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-126">string</span></span> | <span data-ttu-id="26df2-127">Sim</span><span class="sxs-lookup"><span data-stu-id="26df2-127">Yes</span></span> | <span data-ttu-id="26df2-128">Um ID exclusivo para a função.</span><span class="sxs-lookup"><span data-stu-id="26df2-128">A unique ID for the group.</span></span> <span data-ttu-id="26df2-129">Esse ID não deve ser alterado depois de ser definido.</span><span class="sxs-lookup"><span data-stu-id="26df2-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="26df2-130">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-130">string</span></span>  |  <span data-ttu-id="26df2-131">Sim</span><span class="sxs-lookup"><span data-stu-id="26df2-131">Yes</span></span>  |  <span data-ttu-id="26df2-132">O nome da função que os usuários finais veem no Excel.</span><span class="sxs-lookup"><span data-stu-id="26df2-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="26df2-133">No Excel, esse nome de função terá como prefixo o namespace das funções personalizadas especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="26df2-133">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="26df2-134">object</span><span class="sxs-lookup"><span data-stu-id="26df2-134">object</span></span>  |  <span data-ttu-id="26df2-135">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-135">No</span></span>  |  <span data-ttu-id="26df2-136">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="26df2-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="26df2-137">Confira [objeto options](#options-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="26df2-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="26df2-138">matriz</span><span class="sxs-lookup"><span data-stu-id="26df2-138">array</span></span>  |  <span data-ttu-id="26df2-139">Sim</span><span class="sxs-lookup"><span data-stu-id="26df2-139">Yes</span></span>  |  <span data-ttu-id="26df2-140">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="26df2-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="26df2-141">Confira [matriz de parâmetros](#parameters-array) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="26df2-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="26df2-142">objeto</span><span class="sxs-lookup"><span data-stu-id="26df2-142">object</span></span>  |  <span data-ttu-id="26df2-143">Sim</span><span class="sxs-lookup"><span data-stu-id="26df2-143">Yes</span></span>  |  <span data-ttu-id="26df2-144">Objeto que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="26df2-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="26df2-145">Confira [objeto result](#result-object) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="26df2-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="26df2-146">options</span><span class="sxs-lookup"><span data-stu-id="26df2-146">options</span></span>

<span data-ttu-id="26df2-147">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="26df2-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="26df2-148">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="26df2-148">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="26df2-149">Propriedade</span><span class="sxs-lookup"><span data-stu-id="26df2-149">Property</span></span>  |  <span data-ttu-id="26df2-150">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="26df2-150">Data type</span></span>  |  <span data-ttu-id="26df2-151">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="26df2-151">Required</span></span>  |  <span data-ttu-id="26df2-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="26df2-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="26df2-153">booleano</span><span class="sxs-lookup"><span data-stu-id="26df2-153">boolean</span></span>  |  <span data-ttu-id="26df2-154">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-154">No</span></span><br/><br/><span data-ttu-id="26df2-155">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="26df2-155">Default value is 4.</span></span>  |  <span data-ttu-id="26df2-156">Se for `true`, o Excel chama o manipulador `onCanceled` sempre que o usuário executar uma ação que tenha o efeito de cancelar a função; por exemplo, acionando manualmente o recálculo ou editando uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="26df2-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="26df2-157">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="26df2-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="26df2-158">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="26df2-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="26df2-159">No corpo da função, um manipulador deve ser atribuído ao membro `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="26df2-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="26df2-160">Para obter mais informações, consulte [Cancelamento de uma função](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="26df2-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="26df2-161">booleano</span><span class="sxs-lookup"><span data-stu-id="26df2-161">boolean</span></span>  |  <span data-ttu-id="26df2-162">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-162">No</span></span><br/><br/><span data-ttu-id="26df2-163">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="26df2-163">Default value is 4.</span></span>  |  <span data-ttu-id="26df2-164">Se for `true`, a função pode modificar o valor da célula repetidamente, mesmo quando invocada apenas uma vez.</span><span class="sxs-lookup"><span data-stu-id="26df2-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="26df2-165">Essa opção é útil para fontes de dados que mudam rapidamente, como o preço de uma ação.</span><span class="sxs-lookup"><span data-stu-id="26df2-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="26df2-166">Caso você use essa opção, o Excel chamará a função JavaScript com um parâmetro `caller` adicional.</span><span class="sxs-lookup"><span data-stu-id="26df2-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="26df2-167">(***Não*** registre esse parâmetro na propriedade `parameters`).</span><span class="sxs-lookup"><span data-stu-id="26df2-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="26df2-168">A função não deve ter a instrução `return`.</span><span class="sxs-lookup"><span data-stu-id="26df2-168">The function should have no `return` statement.</span></span> <span data-ttu-id="26df2-169">Em vez disso, o valor do resultado é passado como argumento do método de retorno de chamada `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="26df2-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="26df2-170">Para obter mais informações, consulte [Funções de fluxo contínuo](custom-functions-overview.md#streamed-functions).</span><span class="sxs-lookup"><span data-stu-id="26df2-170">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="26df2-171">parameters</span><span class="sxs-lookup"><span data-stu-id="26df2-171">parameters</span></span>

<span data-ttu-id="26df2-172">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="26df2-172">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="26df2-173">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="26df2-173">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="26df2-174">Propriedade</span><span class="sxs-lookup"><span data-stu-id="26df2-174">Property</span></span>  |  <span data-ttu-id="26df2-175">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="26df2-175">Data type</span></span>  |  <span data-ttu-id="26df2-176">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="26df2-176">Required</span></span>  |  <span data-ttu-id="26df2-177">Descrição</span><span class="sxs-lookup"><span data-stu-id="26df2-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="26df2-178">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-178">string</span></span>  |  <span data-ttu-id="26df2-179">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-179">No</span></span> |  <span data-ttu-id="26df2-180">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="26df2-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="26df2-181">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-181">string</span></span>  |  <span data-ttu-id="26df2-182">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-182">No</span></span>  |  <span data-ttu-id="26df2-183">Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="26df2-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="26df2-184">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-184">string</span></span>  |  <span data-ttu-id="26df2-185">Sim</span><span class="sxs-lookup"><span data-stu-id="26df2-185">Yes</span></span>  |  <span data-ttu-id="26df2-186">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="26df2-186">The name of the parameter.</span></span> <span data-ttu-id="26df2-187">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="26df2-187">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="26df2-188">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-188">string</span></span>  |  <span data-ttu-id="26df2-189">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-189">No</span></span>  |  <span data-ttu-id="26df2-190">O tipo de dado do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="26df2-190">The data type of the parameter.</span></span> <span data-ttu-id="26df2-191">Deve ser **boolean**, **number**ou **string**.</span><span class="sxs-lookup"><span data-stu-id="26df2-191">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="26df2-192">result</span><span class="sxs-lookup"><span data-stu-id="26df2-192">result</span></span>

<span data-ttu-id="26df2-193">O objeto `results` define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="26df2-193">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="26df2-194">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="26df2-194">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="26df2-195">Propriedade</span><span class="sxs-lookup"><span data-stu-id="26df2-195">Property</span></span>  |  <span data-ttu-id="26df2-196">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="26df2-196">Data type</span></span>  |  <span data-ttu-id="26df2-197">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="26df2-197">Required</span></span>  |  <span data-ttu-id="26df2-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="26df2-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="26df2-199">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-199">string</span></span>  |  <span data-ttu-id="26df2-200">Não</span><span class="sxs-lookup"><span data-stu-id="26df2-200">No</span></span>  |  <span data-ttu-id="26df2-201">Deve ser **scalar** (um valor não-matriz) ou **matrix** (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="26df2-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="26df2-202">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="26df2-202">string</span></span>  |  <span data-ttu-id="26df2-203">Sim</span><span class="sxs-lookup"><span data-stu-id="26df2-203">Yes</span></span>  |  <span data-ttu-id="26df2-204">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="26df2-204">The data type of the parameter.</span></span> <span data-ttu-id="26df2-205">Deve ser **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="26df2-205">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="26df2-206">Confira também</span><span class="sxs-lookup"><span data-stu-id="26df2-206">See also</span></span>

* [<span data-ttu-id="26df2-207">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="26df2-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="26df2-208">Runtime para funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="26df2-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="26df2-209">Práticas recomendadas para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="26df2-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="26df2-210">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="26df2-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)