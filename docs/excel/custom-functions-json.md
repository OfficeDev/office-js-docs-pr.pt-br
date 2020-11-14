---
ms.date: 11/06/2020
description: Definir metadados JSON para funções personalizadas no Excel e associar suas propriedades de ID de função e nome.
title: Criar manualmente metadados JSON para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: adbcbb9d2705a38b1ed9ff5cdffa6162b9d93a9c
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071638"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="15de1-103">Criar manualmente metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="15de1-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="15de1-104">Conforme descrito no artigo [visão geral das funções personalizadas](custom-functions-overview.md) , um projeto de funções personalizadas deve incluir um arquivo de metadados JSON e um arquivo de script (JavaScript ou TypeScript) para registrar uma função, tornando-a disponível para uso.</span><span class="sxs-lookup"><span data-stu-id="15de1-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="15de1-105">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="15de1-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="15de1-106">Recomendamos usar a autogeração JSON quando possível, em vez de criar seu próprio arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="15de1-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="15de1-107">A autogeração está menos sujeita ao erro do usuário e os `yo office` arquivos do estruturado já incluem isso.</span><span class="sxs-lookup"><span data-stu-id="15de1-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="15de1-108">Para obter mais informações sobre marcas JSDoc e o processo de autogeração JSON, consulte [AutoGenerate Metadata JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="15de1-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="15de1-109">No entanto, você pode criar um projeto de funções personalizadas a partir do zero.</span><span class="sxs-lookup"><span data-stu-id="15de1-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="15de1-110">Esse processo exige:</span><span class="sxs-lookup"><span data-stu-id="15de1-110">This process requires you to:</span></span>

- <span data-ttu-id="15de1-111">Escreva o arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="15de1-111">Write your JSON file.</span></span>
- <span data-ttu-id="15de1-112">Verifique se o arquivo de manifesto está conectado ao arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="15de1-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="15de1-113">Associe suas funções `id` e `name` Propriedades no arquivo de script para registrar suas funções.</span><span class="sxs-lookup"><span data-stu-id="15de1-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="15de1-114">A imagem a seguir explica as diferenças entre o uso `yo office` de arquivos do estruturar e a gravação de JSON do zero.</span><span class="sxs-lookup"><span data-stu-id="15de1-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Imagem das diferenças entre usar Yo Office e escrever seu próprio JSON](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="15de1-116">Lembre-se de conectar seu manifesto ao arquivo JSON criado por você, através da `<Resources>` seção no arquivo de manifesto XML, se você não usar o `yo office` gerador.</span><span class="sxs-lookup"><span data-stu-id="15de1-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="15de1-117">Criação de metadados e conexão com o manifesto</span><span class="sxs-lookup"><span data-stu-id="15de1-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="15de1-118">Crie um arquivo JSON em seu projeto e forneça todos os detalhes sobre suas funções nele, como os parâmetros da função.</span><span class="sxs-lookup"><span data-stu-id="15de1-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="15de1-119">Consulte o [exemplo de metadados a seguir](#json-metadata-example) e [a referência de metadados](#metadata-reference) para obter uma lista completa das propriedades de função.</span><span class="sxs-lookup"><span data-stu-id="15de1-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="15de1-120">Verifique se o arquivo de manifesto XML faz referência ao arquivo JSON na `<Resources>` seção, semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="15de1-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="15de1-121">Exemplo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="15de1-121">JSON metadata example</span></span>

<span data-ttu-id="15de1-122">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="15de1-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="15de1-123">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="15de1-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="15de1-124">Um arquivo JSON de exemplo completo está disponível no histórico de confirmação do repositório do GitHub do [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) .</span><span class="sxs-lookup"><span data-stu-id="15de1-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="15de1-125">À medida que o projeto é ajustado para gerar JSON automaticamente, um exemplo completo de JSON manuscrito só está disponível em versões anteriores do projeto.</span><span class="sxs-lookup"><span data-stu-id="15de1-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="15de1-126">Referência de metadados</span><span class="sxs-lookup"><span data-stu-id="15de1-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="15de1-127">functions</span><span class="sxs-lookup"><span data-stu-id="15de1-127">functions</span></span>

<span data-ttu-id="15de1-128">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="15de1-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="15de1-129">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="15de1-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="15de1-130">Propriedade</span><span class="sxs-lookup"><span data-stu-id="15de1-130">Property</span></span>      | <span data-ttu-id="15de1-131">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="15de1-131">Data type</span></span> | <span data-ttu-id="15de1-132">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15de1-132">Required</span></span> | <span data-ttu-id="15de1-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="15de1-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="15de1-134">string</span><span class="sxs-lookup"><span data-stu-id="15de1-134">string</span></span>    | <span data-ttu-id="15de1-135">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-135">No</span></span>       | <span data-ttu-id="15de1-136">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="15de1-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="15de1-137">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="15de1-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="15de1-138">string</span><span class="sxs-lookup"><span data-stu-id="15de1-138">string</span></span>    | <span data-ttu-id="15de1-139">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-139">No</span></span>       | <span data-ttu-id="15de1-140">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="15de1-140">URL that provides information about the function.</span></span> <span data-ttu-id="15de1-141">(Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="15de1-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="15de1-142">string</span><span class="sxs-lookup"><span data-stu-id="15de1-142">string</span></span>    | <span data-ttu-id="15de1-143">Sim</span><span class="sxs-lookup"><span data-stu-id="15de1-143">Yes</span></span>      | <span data-ttu-id="15de1-144">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="15de1-144">A unique ID for the function.</span></span> <span data-ttu-id="15de1-145">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="15de1-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="15de1-146">string</span><span class="sxs-lookup"><span data-stu-id="15de1-146">string</span></span>    | <span data-ttu-id="15de1-147">Sim</span><span class="sxs-lookup"><span data-stu-id="15de1-147">Yes</span></span>      | <span data-ttu-id="15de1-148">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="15de1-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="15de1-149">No Excel, esse nome de função é prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="15de1-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="15de1-150">object</span><span class="sxs-lookup"><span data-stu-id="15de1-150">object</span></span>    | <span data-ttu-id="15de1-151">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-151">No</span></span>       | <span data-ttu-id="15de1-152">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="15de1-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="15de1-153">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="15de1-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="15de1-154">array</span><span class="sxs-lookup"><span data-stu-id="15de1-154">array</span></span>     | <span data-ttu-id="15de1-155">Sim</span><span class="sxs-lookup"><span data-stu-id="15de1-155">Yes</span></span>      | <span data-ttu-id="15de1-156">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="15de1-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="15de1-157">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="15de1-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="15de1-158">object</span><span class="sxs-lookup"><span data-stu-id="15de1-158">object</span></span>    | <span data-ttu-id="15de1-159">Sim</span><span class="sxs-lookup"><span data-stu-id="15de1-159">Yes</span></span>      | <span data-ttu-id="15de1-160">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="15de1-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="15de1-161">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="15de1-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="15de1-162">options</span><span class="sxs-lookup"><span data-stu-id="15de1-162">options</span></span>

<span data-ttu-id="15de1-163">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="15de1-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="15de1-164">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="15de1-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="15de1-165">Propriedade</span><span class="sxs-lookup"><span data-stu-id="15de1-165">Property</span></span>          | <span data-ttu-id="15de1-166">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="15de1-166">Data type</span></span> | <span data-ttu-id="15de1-167">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15de1-167">Required</span></span>                               | <span data-ttu-id="15de1-168">Descrição</span><span class="sxs-lookup"><span data-stu-id="15de1-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="15de1-169">booliano</span><span class="sxs-lookup"><span data-stu-id="15de1-169">boolean</span></span>   | <span data-ttu-id="15de1-170">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-170">No</span></span><br/><br/><span data-ttu-id="15de1-171">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="15de1-171">Default value is `false`.</span></span>  | <span data-ttu-id="15de1-172">Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="15de1-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="15de1-173">As funções de cancelamento normalmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados.</span><span class="sxs-lookup"><span data-stu-id="15de1-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="15de1-174">Uma função não pode ser streaming e cancelamento.</span><span class="sxs-lookup"><span data-stu-id="15de1-174">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="15de1-175">Para obter mais informações, consulte a observação próxima ao final de [fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="15de1-175">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="15de1-176">booliano</span><span class="sxs-lookup"><span data-stu-id="15de1-176">boolean</span></span>   | <span data-ttu-id="15de1-177">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-177">No</span></span> <br/><br/><span data-ttu-id="15de1-178">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="15de1-178">Default value is `false`.</span></span> | <span data-ttu-id="15de1-179">Se `true` , sua função personalizada pode acessar o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="15de1-179">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="15de1-180">Para obter o endereço da célula que chamou sua função personalizada, use Context. Address em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="15de1-180">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="15de1-181">As funções personalizadas não podem ser definidas como streaming e requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="15de1-181">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="15de1-182">Ao usar essa opção, o parâmetro "invocar" deve ser o último parâmetro passado em opções.</span><span class="sxs-lookup"><span data-stu-id="15de1-182">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
| `stream`          | <span data-ttu-id="15de1-183">booliano</span><span class="sxs-lookup"><span data-stu-id="15de1-183">boolean</span></span>   | <span data-ttu-id="15de1-184">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-184">No</span></span><br/><br/><span data-ttu-id="15de1-185">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="15de1-185">Default value is `false`.</span></span>  | <span data-ttu-id="15de1-186">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="15de1-186">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="15de1-187">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="15de1-187">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="15de1-188">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="15de1-188">The function should have no `return` statement.</span></span> <span data-ttu-id="15de1-189">Em vez disso, o valor resultante é passado como o argumento do método de retorno `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="15de1-189">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="15de1-190">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="15de1-190">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="15de1-191">booliano</span><span class="sxs-lookup"><span data-stu-id="15de1-191">boolean</span></span>   | <span data-ttu-id="15de1-192">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-192">No</span></span> <br/><br/><span data-ttu-id="15de1-193">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="15de1-193">Default value is `false`.</span></span> | <span data-ttu-id="15de1-194">Se `true` , a função será recalculada sempre que o Excel for recalculado, em vez de apenas quando os valores dependentes da fórmula tiverem sido alterados.</span><span class="sxs-lookup"><span data-stu-id="15de1-194">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="15de1-195">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="15de1-195">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="15de1-196">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="15de1-196">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="15de1-197">parâmetros</span><span class="sxs-lookup"><span data-stu-id="15de1-197">parameters</span></span>

<span data-ttu-id="15de1-198">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="15de1-198">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="15de1-199">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="15de1-199">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="15de1-200">Propriedade</span><span class="sxs-lookup"><span data-stu-id="15de1-200">Property</span></span>  |  <span data-ttu-id="15de1-201">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="15de1-201">Data type</span></span>  |  <span data-ttu-id="15de1-202">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15de1-202">Required</span></span>  |  <span data-ttu-id="15de1-203">Descrição</span><span class="sxs-lookup"><span data-stu-id="15de1-203">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="15de1-204">string</span><span class="sxs-lookup"><span data-stu-id="15de1-204">string</span></span>  |  <span data-ttu-id="15de1-205">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-205">No</span></span> |  <span data-ttu-id="15de1-206">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="15de1-206">A description of the parameter.</span></span> <span data-ttu-id="15de1-207">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="15de1-207">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="15de1-208">string</span><span class="sxs-lookup"><span data-stu-id="15de1-208">string</span></span>  |  <span data-ttu-id="15de1-209">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-209">No</span></span>  |  <span data-ttu-id="15de1-210">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="15de1-210">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="15de1-211">string</span><span class="sxs-lookup"><span data-stu-id="15de1-211">string</span></span>  |  <span data-ttu-id="15de1-212">Sim</span><span class="sxs-lookup"><span data-stu-id="15de1-212">Yes</span></span>  |  <span data-ttu-id="15de1-213">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="15de1-213">The name of the parameter.</span></span> <span data-ttu-id="15de1-214">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="15de1-214">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="15de1-215">string</span><span class="sxs-lookup"><span data-stu-id="15de1-215">string</span></span>  |  <span data-ttu-id="15de1-216">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-216">No</span></span>  |  <span data-ttu-id="15de1-217">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="15de1-217">The data type of the parameter.</span></span> <span data-ttu-id="15de1-218">Pode ser **booliano** , **número** , **cadeia de caracteres** ou **qualquer** , que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="15de1-218">Can be **boolean** , **number** , **string** , or **any** , which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="15de1-219">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="15de1-219">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="15de1-220">booliano</span><span class="sxs-lookup"><span data-stu-id="15de1-220">boolean</span></span> | <span data-ttu-id="15de1-221">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-221">No</span></span> | <span data-ttu-id="15de1-222">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="15de1-222">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="15de1-223">booliano</span><span class="sxs-lookup"><span data-stu-id="15de1-223">boolean</span></span> | <span data-ttu-id="15de1-224">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-224">No</span></span> | <span data-ttu-id="15de1-225">Se `true` , os parâmetros são preenchidos de uma matriz especificada.</span><span class="sxs-lookup"><span data-stu-id="15de1-225">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="15de1-226">Observe que funções todos os parâmetros de repetição são considerados parâmetros opcionais por definição.</span><span class="sxs-lookup"><span data-stu-id="15de1-226">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="15de1-227">result</span><span class="sxs-lookup"><span data-stu-id="15de1-227">result</span></span>

<span data-ttu-id="15de1-228">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="15de1-228">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="15de1-229">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="15de1-229">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="15de1-230">Propriedade</span><span class="sxs-lookup"><span data-stu-id="15de1-230">Property</span></span>         | <span data-ttu-id="15de1-231">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="15de1-231">Data type</span></span> | <span data-ttu-id="15de1-232">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15de1-232">Required</span></span> | <span data-ttu-id="15de1-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="15de1-233">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="15de1-234">string</span><span class="sxs-lookup"><span data-stu-id="15de1-234">string</span></span>    | <span data-ttu-id="15de1-235">Não</span><span class="sxs-lookup"><span data-stu-id="15de1-235">No</span></span>       | <span data-ttu-id="15de1-236">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="15de1-236">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="15de1-237">Associar os nomes de função com metadados JSON</span><span class="sxs-lookup"><span data-stu-id="15de1-237">Associating function names with JSON metadata</span></span>

<span data-ttu-id="15de1-238">Para que uma função funcione corretamente, você precisa associar a propriedade da função à `id` implementação do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="15de1-238">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="15de1-239">Verifique se há uma associação, caso contrário, a função não será registrada e não é utilizável no Excel.</span><span class="sxs-lookup"><span data-stu-id="15de1-239">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="15de1-240">O exemplo de código a seguir mostra como fazer a Associação usando o `CustomFunctions.associate()` método.</span><span class="sxs-lookup"><span data-stu-id="15de1-240">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="15de1-241">A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.</span><span class="sxs-lookup"><span data-stu-id="15de1-241">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="15de1-242">O JSON a seguir mostra os metadados JSON que estão associados ao código JavaScript da função personalizada anterior.</span><span class="sxs-lookup"><span data-stu-id="15de1-242">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="15de1-243">Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="15de1-243">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="15de1-244">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="15de1-244">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="15de1-245">No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="15de1-245">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="15de1-246">Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="15de1-246">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="15de1-247">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="15de1-247">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="15de1-248">Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.</span><span class="sxs-lookup"><span data-stu-id="15de1-248">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="15de1-249">No arquivo JavaScript, especifique uma associação de função personalizada usando `CustomFunctions.associate` após cada função.</span><span class="sxs-lookup"><span data-stu-id="15de1-249">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="15de1-250">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas no exemplo de código JavaScript anterior.</span><span class="sxs-lookup"><span data-stu-id="15de1-250">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="15de1-251">Os `id` `name` valores de propriedade e estão em letras maiúsculas, o que é uma prática recomendada ao descrever suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="15de1-251">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="15de1-252">Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não usando a autogeração.</span><span class="sxs-lookup"><span data-stu-id="15de1-252">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="15de1-253">Para obter mais informações sobre a autogeração, consulte [AutoGenerate metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="15de1-253">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
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

## <a name="next-steps"></a><span data-ttu-id="15de1-254">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="15de1-254">Next steps</span></span>

<span data-ttu-id="15de1-255">Conheça as [práticas recomendadas para nomear sua função](custom-functions-naming.md) ou descubra como [localizar sua função](custom-functions-localize.md) usando o método JSON manuscrito descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="15de1-255">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="15de1-256">Confira também</span><span class="sxs-lookup"><span data-stu-id="15de1-256">See also</span></span>

- [<span data-ttu-id="15de1-257">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="15de1-257">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="15de1-258">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="15de1-258">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="15de1-259">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="15de1-259">Create custom functions in Excel</span></span>](custom-functions-overview.md)
