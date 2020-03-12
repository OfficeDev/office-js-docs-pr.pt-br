---
ms.date: 01/14/2020
description: Definir metadados JSON para funções personalizadas no Excel e associar suas propriedades de ID de função e nome.
title: Metadados para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: 79f23f83dfd4bff40880cb39edc6ebe9bf2e052e
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596778"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="f5527-103">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f5527-103">Custom functions metadata</span></span>

<span data-ttu-id="f5527-104">Conforme descrito no artigo [visão geral das funções personalizadas](custom-functions-overview.md) , um projeto de funções personalizadas deve incluir um arquivo de metadados JSON e um arquivo de script (JavaScript ou TypeScript) para registrar uma função, tornando-a disponível para uso.</span><span class="sxs-lookup"><span data-stu-id="f5527-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="f5527-105">As funções personalizadas são registradas quando o usuário executa o suplemento pela primeira vez e depois que eles estão disponíveis para o mesmo usuário em todas as pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f5527-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f5527-106">É recomendável que você use a geração automática JSON quando possível, usando `yo office` os arquivos do estruturar, semelhante ao processo mostrado no [tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md) , pois esse processo é mais fácil e menos sujeito ao erro do usuário.</span><span class="sxs-lookup"><span data-stu-id="f5527-106">It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error.</span></span> <span data-ttu-id="f5527-107">Para obter mais informações sobre o processo de geração de arquivo JSON de comentário JSDoc, consulte [GENERATE JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="f5527-107">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="f5527-108">No entanto, você pode tornar um projeto de funções personalizadas a partir do zero; é necessário:</span><span class="sxs-lookup"><span data-stu-id="f5527-108">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="f5527-109">Gravar seu arquivo JSON manualmente</span><span class="sxs-lookup"><span data-stu-id="f5527-109">Write your JSON file by hand</span></span>
- <span data-ttu-id="f5527-110">Verifique se o arquivo de manifesto está conectado ao arquivo JSON de autoria de mão</span><span class="sxs-lookup"><span data-stu-id="f5527-110">Check that your manifest file is connected to your hand-authored JSON file</span></span>
- <span data-ttu-id="f5527-111">Associar suas funções `id` e `name` Propriedades no arquivo de script para registrar suas funções</span><span class="sxs-lookup"><span data-stu-id="f5527-111">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="f5527-112">Este artigo mostrará como realizar todas as três etapas.</span><span class="sxs-lookup"><span data-stu-id="f5527-112">This article will show you how to do all three of these steps.</span></span>

<span data-ttu-id="f5527-113">A imagem a seguir explica as diferenças entre `yo office` o uso de arquivos do estruturar e a gravação de JSON do zero.</span><span class="sxs-lookup"><span data-stu-id="f5527-113">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>
<span data-ttu-id="f5527-114">![Imagem das diferenças entre usar Yo Office e escrever seu próprio JSON](../images/custom-functions-json.png)</span><span class="sxs-lookup"><span data-stu-id="f5527-114">![Image of differences between using Yo Office and writing your own JSON](../images/custom-functions-json.png)</span></span>

> [!NOTE]
> <span data-ttu-id="f5527-115">Ao contrário dos arquivos `yo office` do estruturar, você precisará conectar seu manifesto ao arquivo JSON que você cria, através da `<Resources>` seção no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f5527-115">In contrast with the `yo office` scaffold files, you need to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file.</span></span> <span data-ttu-id="f5527-116">Observe que as configurações do servidor no servidor que hospeda o arquivo JSON devem ter o [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) habilitado para que as funções personalizadas funcionem corretamente no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="f5527-116">Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="f5527-117">Criação de metadados e conexão com o manifesto</span><span class="sxs-lookup"><span data-stu-id="f5527-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="f5527-118">Você precisa criar um arquivo JSON em seu projeto e fornecer todos os detalhes sobre suas funções nele, como os parâmetros da função.</span><span class="sxs-lookup"><span data-stu-id="f5527-118">You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="f5527-119">Consulte o [exemplo de metadados a seguir](#json-metadata-example) e [a referência de metadados](#metadata-reference) para obter uma lista completa das propriedades de função.</span><span class="sxs-lookup"><span data-stu-id="f5527-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="f5527-120">Você também precisa certificar-se de que seu arquivo de manifesto XML faça referência ao `<Resources>` arquivo JSON na seção, semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f5527-120">You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="f5527-121">Exemplo de metadados JSON</span><span class="sxs-lookup"><span data-stu-id="f5527-121">JSON metadata example</span></span>

<span data-ttu-id="f5527-122">O exemplo a seguir mostra o conteúdo de um arquivo de metadados JSON para um suplemento que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f5527-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="f5527-123">As seções que seguem este exemplo fornecem informações detalhadas sobre as propriedades individuais neste exemplo de JSON.</span><span class="sxs-lookup"><span data-stu-id="f5527-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="f5527-124">Um arquivo JSON de exemplo completo está disponível no histórico de confirmação do repositório do GitHub do [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) .</span><span class="sxs-lookup"><span data-stu-id="f5527-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="f5527-125">À medida que o projeto é ajustado para gerar JSON automaticamente, um exemplo completo de JSON manuscrito só está disponível em versões anteriores do projeto.</span><span class="sxs-lookup"><span data-stu-id="f5527-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="f5527-126">Referência de metadados</span><span class="sxs-lookup"><span data-stu-id="f5527-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="f5527-127">functions</span><span class="sxs-lookup"><span data-stu-id="f5527-127">functions</span></span>

<span data-ttu-id="f5527-128">A propriedade `functions` é um conjunto de objetos de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f5527-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="f5527-129">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="f5527-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="f5527-130">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f5527-130">Property</span></span>      | <span data-ttu-id="f5527-131">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="f5527-131">Data type</span></span> | <span data-ttu-id="f5527-132">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f5527-132">Required</span></span> | <span data-ttu-id="f5527-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="f5527-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="f5527-134">string</span><span class="sxs-lookup"><span data-stu-id="f5527-134">string</span></span>    | <span data-ttu-id="f5527-135">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-135">No</span></span>       | <span data-ttu-id="f5527-136">Descrição da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="f5527-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="f5527-137">Por exemplo, **Converte um valor em Celsius para Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="f5527-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="f5527-138">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f5527-138">string</span></span>    | <span data-ttu-id="f5527-139">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-139">No</span></span>       | <span data-ttu-id="f5527-140">A URL que fornece informações sobre a função.</span><span class="sxs-lookup"><span data-stu-id="f5527-140">URL that provides information about the function.</span></span> <span data-ttu-id="f5527-141">(Ela é exibida em um painel de tarefas). Por exemplo, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="f5527-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="f5527-142">string</span><span class="sxs-lookup"><span data-stu-id="f5527-142">string</span></span>    | <span data-ttu-id="f5527-143">Sim</span><span class="sxs-lookup"><span data-stu-id="f5527-143">Yes</span></span>      | <span data-ttu-id="f5527-144">Identificação exclusiva para a função.</span><span class="sxs-lookup"><span data-stu-id="f5527-144">A unique ID for the function.</span></span> <span data-ttu-id="f5527-145">Essa ID pode conter apenas caracteres alfanuméricos e pontos e não deve ser alterada depois de configurada.</span><span class="sxs-lookup"><span data-stu-id="f5527-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="f5527-146">string</span><span class="sxs-lookup"><span data-stu-id="f5527-146">string</span></span>    | <span data-ttu-id="f5527-147">Sim</span><span class="sxs-lookup"><span data-stu-id="f5527-147">Yes</span></span>      | <span data-ttu-id="f5527-148">O nome da função que é exibida aos usuários finais no Excel.</span><span class="sxs-lookup"><span data-stu-id="f5527-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="f5527-149">No Excel, o nome da função será prefixado pelo namespace de funções personalizadas que é especificado no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f5527-149">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="f5527-150">object</span><span class="sxs-lookup"><span data-stu-id="f5527-150">object</span></span>    | <span data-ttu-id="f5527-151">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-151">No</span></span>       | <span data-ttu-id="f5527-152">Permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="f5527-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="f5527-153">Confira [opções](#options) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="f5527-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="f5527-154">array</span><span class="sxs-lookup"><span data-stu-id="f5527-154">array</span></span>     | <span data-ttu-id="f5527-155">Sim</span><span class="sxs-lookup"><span data-stu-id="f5527-155">Yes</span></span>      | <span data-ttu-id="f5527-156">Matriz que define os parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="f5527-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="f5527-157">Confira os [parâmetros](#parameters) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="f5527-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="f5527-158">object</span><span class="sxs-lookup"><span data-stu-id="f5527-158">object</span></span>    | <span data-ttu-id="f5527-159">Sim</span><span class="sxs-lookup"><span data-stu-id="f5527-159">Yes</span></span>      | <span data-ttu-id="f5527-160">Objeto que define o tipo de informação que é retornada pela função do Excel.</span><span class="sxs-lookup"><span data-stu-id="f5527-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="f5527-161">Confira [resultado](#result) para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="f5527-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="f5527-162">options</span><span class="sxs-lookup"><span data-stu-id="f5527-162">options</span></span>

<span data-ttu-id="f5527-163">O objeto `options` permite que você personalize alguns aspectos de como e quando o Excel executa a função.</span><span class="sxs-lookup"><span data-stu-id="f5527-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="f5527-164">A tabela a seguir lista as propriedades do objeto `options`.</span><span class="sxs-lookup"><span data-stu-id="f5527-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="f5527-165">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f5527-165">Property</span></span>          | <span data-ttu-id="f5527-166">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="f5527-166">Data type</span></span> | <span data-ttu-id="f5527-167">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f5527-167">Required</span></span>                               | <span data-ttu-id="f5527-168">Descrição</span><span class="sxs-lookup"><span data-stu-id="f5527-168">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="f5527-169">booliano</span><span class="sxs-lookup"><span data-stu-id="f5527-169">boolean</span></span>   | <span data-ttu-id="f5527-170">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-170">No</span></span><br/><br/><span data-ttu-id="f5527-171">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="f5527-171">Default value is `false`.</span></span>  | <span data-ttu-id="f5527-172">Se o valor for `true`, o Excel chamará o manipulador `CancelableInvocation` sempre que o usuário realizar uma ação que tenha o efeito de cancelar a função, por exemplo, manualmente acionar um recálculo ou editar uma célula referenciada pela função.</span><span class="sxs-lookup"><span data-stu-id="f5527-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="f5527-173">As funções de cancelamento normalmente são usadas apenas para funções assíncronas que retornam um único resultado e precisam lidar com o cancelamento de uma solicitação de dados.</span><span class="sxs-lookup"><span data-stu-id="f5527-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="f5527-174">Uma função não pode ser streaming e cancelamento.</span><span class="sxs-lookup"><span data-stu-id="f5527-174">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="f5527-175">Para obter mais informações, consulte a observação próxima ao final de [fazer uma função de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="f5527-175">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="f5527-176">booliano</span><span class="sxs-lookup"><span data-stu-id="f5527-176">boolean</span></span>   | <span data-ttu-id="f5527-177">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-177">No</span></span> <br/><br/><span data-ttu-id="f5527-178">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="f5527-178">Default value is `false`.</span></span> | <span data-ttu-id="f5527-179">Se `true`, sua função personalizada pode acessar o endereço da célula que invocou sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f5527-179">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f5527-180">Para obter o endereço da célula que chamou sua função personalizada, use Context. Address em sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f5527-180">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="f5527-181">Para saber mais, confira o [parâmetro context da célula de endereçamento](../excel/custom-functions-parameter-options.md#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="f5527-181">For more information, see [Addressing cell's context parameter](../excel/custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span> <span data-ttu-id="f5527-182">As funções personalizadas não podem ser definidas como streaming e requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="f5527-182">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="f5527-183">Ao usar essa opção, o parâmetro "invocar" deve ser o último parâmetro passado em opções.</span><span class="sxs-lookup"><span data-stu-id="f5527-183">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="f5527-184">booliano</span><span class="sxs-lookup"><span data-stu-id="f5527-184">boolean</span></span>   | <span data-ttu-id="f5527-185">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-185">No</span></span><br/><br/><span data-ttu-id="f5527-186">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="f5527-186">Default value is `false`.</span></span>  | <span data-ttu-id="f5527-187">Se o valor for `true`, a função poderá gerar uma saída para a célula de forma repetida, mesmo quando invocada somente uma vez.</span><span class="sxs-lookup"><span data-stu-id="f5527-187">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="f5527-188">Essa opção é útil para fontes de dados que mudam constantemente, como preços de ações.</span><span class="sxs-lookup"><span data-stu-id="f5527-188">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="f5527-189">A função não deve ter instruções `return`.</span><span class="sxs-lookup"><span data-stu-id="f5527-189">The function should have no `return` statement.</span></span> <span data-ttu-id="f5527-190">Em vez disso, o valor resultante é passado como o argumento do método de retorno `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="f5527-190">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="f5527-191">Para saber mais informações, confira [Funções de streaming](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="f5527-191">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="f5527-192">booliano</span><span class="sxs-lookup"><span data-stu-id="f5527-192">boolean</span></span>   | <span data-ttu-id="f5527-193">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-193">No</span></span> <br/><br/><span data-ttu-id="f5527-194">O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="f5527-194">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="f5527-195">Se for `true`, a função será recalculada sempre que o Excel recalcular, em vez de apenas quando os valores dependentes da fórmula forem alterados.</span><span class="sxs-lookup"><span data-stu-id="f5527-195">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="f5527-196">Uma função não pode ser de streaming e volátil ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="f5527-196">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="f5527-197">Se as propriedades `stream` e `volatile` forem definidas como `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="f5527-197">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="f5527-198">parâmetros</span><span class="sxs-lookup"><span data-stu-id="f5527-198">parameters</span></span>

<span data-ttu-id="f5527-199">A propriedade `parameters` é uma matriz de objetos de parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f5527-199">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="f5527-200">A tabela a seguir lista as propriedades de cada objeto.</span><span class="sxs-lookup"><span data-stu-id="f5527-200">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="f5527-201">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f5527-201">Property</span></span>  |  <span data-ttu-id="f5527-202">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="f5527-202">Data type</span></span>  |  <span data-ttu-id="f5527-203">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f5527-203">Required</span></span>  |  <span data-ttu-id="f5527-204">Descrição</span><span class="sxs-lookup"><span data-stu-id="f5527-204">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="f5527-205">string</span><span class="sxs-lookup"><span data-stu-id="f5527-205">string</span></span>  |  <span data-ttu-id="f5527-206">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-206">No</span></span> |  <span data-ttu-id="f5527-207">Uma descrição do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f5527-207">A description of the parameter.</span></span> <span data-ttu-id="f5527-208">Isso é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="f5527-208">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="f5527-209">string</span><span class="sxs-lookup"><span data-stu-id="f5527-209">string</span></span>  |  <span data-ttu-id="f5527-210">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-210">No</span></span>  |  <span data-ttu-id="f5527-211">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="f5527-211">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="f5527-212">string</span><span class="sxs-lookup"><span data-stu-id="f5527-212">string</span></span>  |  <span data-ttu-id="f5527-213">Sim</span><span class="sxs-lookup"><span data-stu-id="f5527-213">Yes</span></span>  |  <span data-ttu-id="f5527-214">O nome do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f5527-214">The name of the parameter.</span></span> <span data-ttu-id="f5527-215">Esse nome é exibido no IntelliSense do Excel.</span><span class="sxs-lookup"><span data-stu-id="f5527-215">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="f5527-216">string</span><span class="sxs-lookup"><span data-stu-id="f5527-216">string</span></span>  |  <span data-ttu-id="f5527-217">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-217">No</span></span>  |  <span data-ttu-id="f5527-218">O tipo de dados do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f5527-218">The data type of the parameter.</span></span> <span data-ttu-id="f5527-219">Pode ser **booliano**, **número**, **cadeia de caracteres** ou **qualquer**, que permita usar qualquer um dos três tipos anteriores.</span><span class="sxs-lookup"><span data-stu-id="f5527-219">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="f5527-220">Se essa propriedade não for especificada, o tipo de dados padrão será **qualquer**.</span><span class="sxs-lookup"><span data-stu-id="f5527-220">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="f5527-221">booliano</span><span class="sxs-lookup"><span data-stu-id="f5527-221">boolean</span></span> | <span data-ttu-id="f5527-222">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-222">No</span></span> | <span data-ttu-id="f5527-223">Se for `true`, o parâmetro será opcional.</span><span class="sxs-lookup"><span data-stu-id="f5527-223">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="f5527-224">booliano</span><span class="sxs-lookup"><span data-stu-id="f5527-224">boolean</span></span> | <span data-ttu-id="f5527-225">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-225">No</span></span> | <span data-ttu-id="f5527-226">Se `true`, os parâmetros serão preenchidos a partir de uma matriz especificada.</span><span class="sxs-lookup"><span data-stu-id="f5527-226">If `true`, parameters will populate from a specified array.</span></span> <span data-ttu-id="f5527-227">Observe que funções todos os parâmetros de repetição são considerados parâmetros opcionais por definição.</span><span class="sxs-lookup"><span data-stu-id="f5527-227">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="f5527-228">result</span><span class="sxs-lookup"><span data-stu-id="f5527-228">result</span></span>

<span data-ttu-id="f5527-229">O objeto `result` que define o tipo de informação que é retornado pela função.</span><span class="sxs-lookup"><span data-stu-id="f5527-229">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="f5527-230">A tabela a seguir lista as propriedades do objeto `result`.</span><span class="sxs-lookup"><span data-stu-id="f5527-230">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="f5527-231">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f5527-231">Property</span></span>         | <span data-ttu-id="f5527-232">Tipo de dados</span><span class="sxs-lookup"><span data-stu-id="f5527-232">Data type</span></span> | <span data-ttu-id="f5527-233">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f5527-233">Required</span></span> | <span data-ttu-id="f5527-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="f5527-234">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="f5527-235">string</span><span class="sxs-lookup"><span data-stu-id="f5527-235">string</span></span>    | <span data-ttu-id="f5527-236">Não</span><span class="sxs-lookup"><span data-stu-id="f5527-236">No</span></span>       | <span data-ttu-id="f5527-237">Deve ser **escalar** (um valor não matriz) ou **matriz** (uma matriz de 2 dimensões).</span><span class="sxs-lookup"><span data-stu-id="f5527-237">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="f5527-238">Associar os nomes de função com metadados JSON</span><span class="sxs-lookup"><span data-stu-id="f5527-238">Associating function names with JSON metadata</span></span>

<span data-ttu-id="f5527-239">Para que uma função funcione corretamente, você precisa associar a propriedade da `id` função à implementação do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f5527-239">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="f5527-240">Verifique se há uma associação, caso contrário, a função não será registrada e não é utilizável no Excel.</span><span class="sxs-lookup"><span data-stu-id="f5527-240">Make sure there is an association, otherwise the function will not be registered and not useable in Excel.</span></span> <span data-ttu-id="f5527-241">O exemplo de código a seguir mostra como fazer a Associação usando `CustomFunctions.associate()` o método.</span><span class="sxs-lookup"><span data-stu-id="f5527-241">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="f5527-242">A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.</span><span class="sxs-lookup"><span data-stu-id="f5527-242">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="f5527-243">O JSON a seguir mostra os metadados JSON que estão associados ao código JavaScript da função personalizada anterior.</span><span class="sxs-lookup"><span data-stu-id="f5527-243">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="f5527-244">Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="f5527-244">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="f5527-245">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="f5527-245">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="f5527-246">No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="f5527-246">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="f5527-247">Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="f5527-247">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="f5527-248">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="f5527-248">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="f5527-249">Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.</span><span class="sxs-lookup"><span data-stu-id="f5527-249">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="f5527-250">No arquivo JavaScript, especifique uma associação de função personalizada usando `CustomFunctions.associate` após cada função.</span><span class="sxs-lookup"><span data-stu-id="f5527-250">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="f5527-251">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f5527-251">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="f5527-252">Os `id` valores `name` de propriedade e estão em letras maiúsculas, o que é uma prática recomendada ao descrever suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f5527-252">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="f5527-253">Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não usando a autogeração.</span><span class="sxs-lookup"><span data-stu-id="f5527-253">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="f5527-254">Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="f5527-254">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="f5527-255">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f5527-255">Next steps</span></span>

<span data-ttu-id="f5527-256">Conheça as [práticas recomendadas para nomear sua função](custom-functions-naming.md) ou descubra como [localizar sua função](custom-functions-localize.md) usando o método JSON manuscrito descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="f5527-256">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="f5527-257">Confira também
</span><span class="sxs-lookup"><span data-stu-id="f5527-257">See also</span></span>

- [<span data-ttu-id="f5527-258">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f5527-258">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="f5527-259">Opções de parâmetros de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f5527-259">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="f5527-260">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="f5527-260">Create custom functions in Excel</span></span>](custom-functions-overview.md)
