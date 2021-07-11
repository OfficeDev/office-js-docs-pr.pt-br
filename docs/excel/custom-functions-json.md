---
ms.date: 12/22/2020
description: Defina os metadados JSON para funções personalizadas no Excel e associe sua ID de função e propriedades de nome.
title: Criar metadados JSON manualmente para funções personalizadas Excel
localization_priority: Normal
ms.openlocfilehash: c03238d46e8d861307ba0db3d03dafea81aeca51
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349627"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="cb939-103">Criar metadados JSON manualmente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cb939-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="cb939-104">Conforme descrito no artigo visão geral de funções [personalizadas,](custom-functions-overview.md) um projeto de funções personalizadas deve incluir um arquivo de metadados JSON e um arquivo de script (JavaScript ou TypeScript) para registrar uma função, disponibilizando-a para uso.</span><span class="sxs-lookup"><span data-stu-id="cb939-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="cb939-105">As funções personalizadas são registradas quando o usuário executa o complemento pela primeira vez e depois disso estão disponíveis para o mesmo usuário em todas as guias de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb939-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cb939-106">Recomendamos usar a geração automática JSON quando possível, em vez de criar seu próprio arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="cb939-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="cb939-107">A geração automática é menos propensa a erros do usuário e os arquivos `yo office` scaffolded já incluem isso.</span><span class="sxs-lookup"><span data-stu-id="cb939-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="cb939-108">Para obter mais informações sobre marcas JSDoc e o processo de geração automática JSON, consulte [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="cb939-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="cb939-109">No entanto, você pode fazer um projeto de funções personalizadas do zero.</span><span class="sxs-lookup"><span data-stu-id="cb939-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="cb939-110">Esse processo exige que você:</span><span class="sxs-lookup"><span data-stu-id="cb939-110">This process requires you to:</span></span>

- <span data-ttu-id="cb939-111">Escreva seu arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="cb939-111">Write your JSON file.</span></span>
- <span data-ttu-id="cb939-112">Verifique se o arquivo de manifesto está conectado ao arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="cb939-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="cb939-113">Associe suas funções `id` e propriedades no arquivo de script para registrar suas `name` funções.</span><span class="sxs-lookup"><span data-stu-id="cb939-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="cb939-114">A imagem a seguir explica as diferenças entre usar `yo office` arquivos scaffold e escrever JSON do zero.</span><span class="sxs-lookup"><span data-stu-id="cb939-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Imagem das diferenças entre usar Yo Office e escrever seu próprio JSON.](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="cb939-116">Lembre-se de conectar seu manifesto ao arquivo JSON que você criar, por meio da seção em seu arquivo de manifesto XML se `<Resources>` você não usar o `yo office` gerador.</span><span class="sxs-lookup"><span data-stu-id="cb939-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="cb939-117">Autoria de metadados e conexão com o manifesto</span><span class="sxs-lookup"><span data-stu-id="cb939-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="cb939-118">Crie um arquivo JSON em seu projeto e forneça todos os detalhes sobre suas funções nele, como os parâmetros da função.</span><span class="sxs-lookup"><span data-stu-id="cb939-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="cb939-119">Consulte o [exemplo de metadados a](#json-metadata-example) seguir e a referência de [metadados](#metadata-reference) para uma lista completa de propriedades de função.</span><span class="sxs-lookup"><span data-stu-id="cb939-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="cb939-120">Verifique se o arquivo de manifesto XML faz referência ao arquivo JSON na `<Resources>` seção, semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="cb939-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a><span data-ttu-id="cb939-121">Exemplo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="cb939-121">JSON metadata example</span></span>

<span data-ttu-id="cb939-122">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cb939-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="cb939-123">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="cb939-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
      "description": "Count up from zero",
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
      "description": "Get the second highest number from a range",
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
> <span data-ttu-id="cb939-124">Um arquivo JSON de exemplo completo está disponível no [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub histórico de confirmação do repositório.</span><span class="sxs-lookup"><span data-stu-id="cb939-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="cb939-125">Como o projeto foi ajustado para gerar automaticamente o JSON, um exemplo completo de JSON manuscrito só está disponível em versões anteriores do projeto.</span><span class="sxs-lookup"><span data-stu-id="cb939-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="cb939-126">Referência de metadados</span><span class="sxs-lookup"><span data-stu-id="cb939-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="cb939-127">functions</span><span class="sxs-lookup"><span data-stu-id="cb939-127">functions</span></span>

<span data-ttu-id="cb939-128">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cb939-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="cb939-129">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="cb939-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="cb939-130">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cb939-130">Property</span></span>      | <span data-ttu-id="cb939-131">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="cb939-131">Data type</span></span> | <span data-ttu-id="cb939-132">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cb939-132">Required</span></span> | <span data-ttu-id="cb939-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="cb939-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="cb939-134">string</span><span class="sxs-lookup"><span data-stu-id="cb939-134">string</span></span>    | <span data-ttu-id="cb939-135">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-135">No</span></span>       | <span data-ttu-id="cb939-136">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="cb939-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="cb939-137">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="cb939-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="cb939-138">string</span><span class="sxs-lookup"><span data-stu-id="cb939-138">string</span></span>    | <span data-ttu-id="cb939-139">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-139">No</span></span>       | <span data-ttu-id="cb939-140">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="cb939-140">URL that provides information about the function.</span></span> <span data-ttu-id="cb939-141">(Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="cb939-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="cb939-142">string</span><span class="sxs-lookup"><span data-stu-id="cb939-142">string</span></span>    | <span data-ttu-id="cb939-143">Sim</span><span class="sxs-lookup"><span data-stu-id="cb939-143">Yes</span></span>      | <span data-ttu-id="cb939-144">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="cb939-144">A unique ID for the function.</span></span> <span data-ttu-id="cb939-145">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="cb939-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="cb939-146">string</span><span class="sxs-lookup"><span data-stu-id="cb939-146">string</span></span>    | <span data-ttu-id="cb939-147">Sim</span><span class="sxs-lookup"><span data-stu-id="cb939-147">Yes</span></span>      | <span data-ttu-id="cb939-148">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="cb939-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="cb939-149">Em Excel, esse nome de função é prefixado pelo namespace de funções personalizadas especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="cb939-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="cb939-150">object</span><span class="sxs-lookup"><span data-stu-id="cb939-150">object</span></span>    | <span data-ttu-id="cb939-151">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-151">No</span></span>       | <span data-ttu-id="cb939-152">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="cb939-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="cb939-153">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="cb939-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="cb939-154">array</span><span class="sxs-lookup"><span data-stu-id="cb939-154">array</span></span>     | <span data-ttu-id="cb939-155">Sim</span><span class="sxs-lookup"><span data-stu-id="cb939-155">Yes</span></span>      | <span data-ttu-id="cb939-156">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="cb939-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="cb939-157">Consulte [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="cb939-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="cb939-158">object</span><span class="sxs-lookup"><span data-stu-id="cb939-158">object</span></span>    | <span data-ttu-id="cb939-159">Sim</span><span class="sxs-lookup"><span data-stu-id="cb939-159">Yes</span></span>      | <span data-ttu-id="cb939-160">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="cb939-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="cb939-161">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="cb939-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="cb939-162">options</span><span class="sxs-lookup"><span data-stu-id="cb939-162">options</span></span>

<span data-ttu-id="cb939-163">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="cb939-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="cb939-164">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="cb939-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="cb939-165">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cb939-165">Property</span></span>          | <span data-ttu-id="cb939-166">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="cb939-166">Data type</span></span> | <span data-ttu-id="cb939-167">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cb939-167">Required</span></span>                               | <span data-ttu-id="cb939-168">Descrição</span><span class="sxs-lookup"><span data-stu-id="cb939-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="cb939-169">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-169">boolean</span></span>   | <span data-ttu-id="cb939-170">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-170">No</span></span><br/><br/><span data-ttu-id="cb939-171">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="cb939-171">Default value is `false`.</span></span>  | <span data-ttu-id="cb939-172">Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="cb939-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="cb939-173">Funções canceláveis geralmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados.</span><span class="sxs-lookup"><span data-stu-id="cb939-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="cb939-174">Uma função não pode usar as `stream` propriedades `cancelable` e.</span><span class="sxs-lookup"><span data-stu-id="cb939-174">A function can't use both the `stream` and `cancelable` properties.</span></span> |
| `requiresAddress` | <span data-ttu-id="cb939-175">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-175">boolean</span></span>   | <span data-ttu-id="cb939-176">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-176">No</span></span> <br/><br/><span data-ttu-id="cb939-177">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="cb939-177">Default value is `false`.</span></span> | <span data-ttu-id="cb939-178">Se `true` , sua função personalizada pode acessar o endereço da célula que a invocou.</span><span class="sxs-lookup"><span data-stu-id="cb939-178">If `true`, your custom function can access the address of the cell that invoked it.</span></span> <span data-ttu-id="cb939-179">A `address` propriedade do parâmetro [invocação](custom-functions-parameter-options.md#invocation-parameter) contém o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cb939-179">The `address` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="cb939-180">Uma função não pode usar as `stream` propriedades `requiresAddress` e.</span><span class="sxs-lookup"><span data-stu-id="cb939-180">A function can't use both the `stream` and `requiresAddress` properties.</span></span> |
| `requiresParameterAddresses` | <span data-ttu-id="cb939-181">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-181">boolean</span></span>   | <span data-ttu-id="cb939-182">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-182">No</span></span> <br/><br/><span data-ttu-id="cb939-183">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="cb939-183">Default value is `false`.</span></span> | <span data-ttu-id="cb939-184">Se `true` , sua função personalizada pode acessar os endereços dos parâmetros de entrada da função.</span><span class="sxs-lookup"><span data-stu-id="cb939-184">If `true`, your custom function can access the addresses of the function's input parameters.</span></span> <span data-ttu-id="cb939-185">Essa propriedade deve ser usada em combinação com a propriedade do objeto de resultado e deve `dimensionality` ser definida como [](#result) `dimensionality` `matrix` .</span><span class="sxs-lookup"><span data-stu-id="cb939-185">This property must be used in combination with the `dimensionality` property of the [result](#result) object, and `dimensionality` must be set to `matrix`.</span></span> <span data-ttu-id="cb939-186">Consulte [Detectar o endereço de um parâmetro para](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="cb939-186">See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.</span></span> |
| `stream`          | <span data-ttu-id="cb939-187">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-187">boolean</span></span>   | <span data-ttu-id="cb939-188">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-188">No</span></span><br/><br/><span data-ttu-id="cb939-189">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="cb939-189">Default value is `false`.</span></span>  | <span data-ttu-id="cb939-190">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="cb939-190">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="cb939-191">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="cb939-191">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="cb939-192">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="cb939-192">The function should have no `return` statement.</span></span> <span data-ttu-id="cb939-193">Em vez disso, o valor resultante é passado como o argumento do método de retorno `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="cb939-193">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="cb939-194">Para obter mais informações, consulte [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="cb939-194">For more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="cb939-195">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-195">boolean</span></span>   | <span data-ttu-id="cb939-196">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-196">No</span></span> <br/><br/><span data-ttu-id="cb939-197">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="cb939-197">Default value is `false`.</span></span> | <span data-ttu-id="cb939-198">Se , a função será recalculada sempre que Excel recalcular, em vez de somente quando os valores dependentes da fórmula `true` foram alterados.</span><span class="sxs-lookup"><span data-stu-id="cb939-198">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="cb939-199">Uma função não pode usar as `stream` propriedades `volatile` e.</span><span class="sxs-lookup"><span data-stu-id="cb939-199">A function can't use both the `stream` and `volatile` properties.</span></span> <span data-ttu-id="cb939-200">Se as `stream` propriedades `volatile` e estão definidas como , a `true` propriedade volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="cb939-200">If the `stream` and `volatile` properties are both set to `true`, the volatile property will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="cb939-201">parâmetros</span><span class="sxs-lookup"><span data-stu-id="cb939-201">parameters</span></span>

<span data-ttu-id="cb939-202">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cb939-202">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="cb939-203">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="cb939-203">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="cb939-204">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cb939-204">Property</span></span>  |  <span data-ttu-id="cb939-205">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="cb939-205">Data type</span></span>  |  <span data-ttu-id="cb939-206">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cb939-206">Required</span></span>  |  <span data-ttu-id="cb939-207">Descrição</span><span class="sxs-lookup"><span data-stu-id="cb939-207">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="cb939-208">string</span><span class="sxs-lookup"><span data-stu-id="cb939-208">string</span></span>  |  <span data-ttu-id="cb939-209">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-209">No</span></span> |  <span data-ttu-id="cb939-210">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cb939-210">A description of the parameter.</span></span> <span data-ttu-id="cb939-211">Isso é exibido no Excel do IntelliSense.</span><span class="sxs-lookup"><span data-stu-id="cb939-211">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="cb939-212">string</span><span class="sxs-lookup"><span data-stu-id="cb939-212">string</span></span>  |  <span data-ttu-id="cb939-213">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-213">No</span></span>  |  <span data-ttu-id="cb939-214">Deve ser `scalar` (um valor não matriz) `matrix` ou (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="cb939-214">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="cb939-215">string</span><span class="sxs-lookup"><span data-stu-id="cb939-215">string</span></span>  |  <span data-ttu-id="cb939-216">Sim</span><span class="sxs-lookup"><span data-stu-id="cb939-216">Yes</span></span>  |  <span data-ttu-id="cb939-217">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cb939-217">The name of the parameter.</span></span> <span data-ttu-id="cb939-218">Esse nome é exibido no Excel do IntelliSense.</span><span class="sxs-lookup"><span data-stu-id="cb939-218">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="cb939-219">string</span><span class="sxs-lookup"><span data-stu-id="cb939-219">string</span></span>  |  <span data-ttu-id="cb939-220">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-220">No</span></span>  |  <span data-ttu-id="cb939-221">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cb939-221">The data type of the parameter.</span></span> <span data-ttu-id="cb939-222">Pode ser , , ou , o que permite que você `boolean` use qualquer um dos três tipos `number` `string` `any` anteriores.</span><span class="sxs-lookup"><span data-stu-id="cb939-222">Can be `boolean`, `number`, `string`, or `any`, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="cb939-223">Se essa propriedade não for especificada, o tipo de dados será padrão para `any` .</span><span class="sxs-lookup"><span data-stu-id="cb939-223">If this property is not specified, the data type defaults to `any`.</span></span> |
|  `optional`  | <span data-ttu-id="cb939-224">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-224">boolean</span></span> | <span data-ttu-id="cb939-225">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-225">No</span></span> | <span data-ttu-id="cb939-226">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="cb939-226">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="cb939-227">booliano</span><span class="sxs-lookup"><span data-stu-id="cb939-227">boolean</span></span> | <span data-ttu-id="cb939-228">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-228">No</span></span> | <span data-ttu-id="cb939-229">If `true` , parameters populate from a specified array.</span><span class="sxs-lookup"><span data-stu-id="cb939-229">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="cb939-230">Observe que todas as funções de todos os parâmetros repetidos são consideradas parâmetros opcionais por definição.</span><span class="sxs-lookup"><span data-stu-id="cb939-230">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="cb939-231">result</span><span class="sxs-lookup"><span data-stu-id="cb939-231">result</span></span>

<span data-ttu-id="cb939-232">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="cb939-232">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="cb939-233">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="cb939-233">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="cb939-234">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cb939-234">Property</span></span>         | <span data-ttu-id="cb939-235">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="cb939-235">Data type</span></span> | <span data-ttu-id="cb939-236">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cb939-236">Required</span></span> | <span data-ttu-id="cb939-237">Descrição</span><span class="sxs-lookup"><span data-stu-id="cb939-237">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="cb939-238">string</span><span class="sxs-lookup"><span data-stu-id="cb939-238">string</span></span>    | <span data-ttu-id="cb939-239">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-239">No</span></span>       | <span data-ttu-id="cb939-240">Deve ser `scalar` (um valor não matriz) `matrix` ou (uma matriz bidimensional).</span><span class="sxs-lookup"><span data-stu-id="cb939-240">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span> |
| `type` | <span data-ttu-id="cb939-241">string</span><span class="sxs-lookup"><span data-stu-id="cb939-241">string</span></span>    | <span data-ttu-id="cb939-242">Não</span><span class="sxs-lookup"><span data-stu-id="cb939-242">No</span></span>       | <span data-ttu-id="cb939-243">O tipo de dados do resultado.</span><span class="sxs-lookup"><span data-stu-id="cb939-243">The data type of the result.</span></span> <span data-ttu-id="cb939-244">Pode ser , , ou (que permite que você `boolean` use qualquer um dos três tipos `number` `string` `any` anteriores).</span><span class="sxs-lookup"><span data-stu-id="cb939-244">Can be `boolean`, `number`, `string`, or `any` (which allows you to use of any of the previous three types).</span></span> <span data-ttu-id="cb939-245">Se essa propriedade não for especificada, o tipo de dados será padrão para `any` .</span><span class="sxs-lookup"><span data-stu-id="cb939-245">If this property is not specified, the data type defaults to `any`.</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="cb939-246">Associar os nomes de função com metadados JSON</span><span class="sxs-lookup"><span data-stu-id="cb939-246">Associating function names with JSON metadata</span></span>

<span data-ttu-id="cb939-247">Para que uma função funcione corretamente, você precisa associar a propriedade da função `id` à implementação do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cb939-247">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="cb939-248">Certifique-se de que haja uma associação, caso contrário, a função não será registrada e não será Excel.</span><span class="sxs-lookup"><span data-stu-id="cb939-248">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="cb939-249">O exemplo de código a seguir mostra como fazer a associação usando o `CustomFunctions.associate()` método.</span><span class="sxs-lookup"><span data-stu-id="cb939-249">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="cb939-250">A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.</span><span class="sxs-lookup"><span data-stu-id="cb939-250">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="cb939-251">O JSON a seguir mostra os metadados JSON associados à função personalizada anterior código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cb939-251">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

<span data-ttu-id="cb939-252">Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="cb939-252">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="cb939-253">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="cb939-253">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="cb939-254">No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="cb939-254">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="cb939-255">Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="cb939-255">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="cb939-256">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="cb939-256">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="cb939-257">Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.</span><span class="sxs-lookup"><span data-stu-id="cb939-257">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="cb939-258">No arquivo JavaScript, especifique uma associação de função personalizada usando `CustomFunctions.associate` após cada função.</span><span class="sxs-lookup"><span data-stu-id="cb939-258">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="cb939-259">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas no exemplo de código JavaScript anterior.</span><span class="sxs-lookup"><span data-stu-id="cb939-259">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="cb939-260">Os `id` valores e propriedades estão em `name` maiúsculas, o que é uma prática prática prática ao descrever suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cb939-260">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="cb939-261">Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não usando a geração automática.</span><span class="sxs-lookup"><span data-stu-id="cb939-261">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="cb939-262">Para obter mais informações sobre a geração automática, consulte [Metadados JSON](custom-functions-json-autogeneration.md)de geração automática para funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cb939-262">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a><span data-ttu-id="cb939-263">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="cb939-263">Next steps</span></span>

<span data-ttu-id="cb939-264">Aprenda as [práticas recomendadas para nomear sua](custom-functions-naming.md) função ou descubra como [localizar sua](custom-functions-localize.md) função usando o método JSON escrito à mão descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="cb939-264">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="cb939-265">Confira também</span><span class="sxs-lookup"><span data-stu-id="cb939-265">See also</span></span>

- [<span data-ttu-id="cb939-266">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cb939-266">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="cb939-267">Opções de parâmetro de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cb939-267">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="cb939-268">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="cb939-268">Create custom functions in Excel</span></span>](custom-functions-overview.md)
