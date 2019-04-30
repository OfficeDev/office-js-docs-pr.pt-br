---
ms.date: 04/20/2019
description: Criar funções personalizadas no Excel usando JavaScript.
title: Criar funções personalizadas no Excel (versão prévia)
localization_priority: Priority
ms.openlocfilehash: 634b76ed90a30c7aa8252da346ba3f95684967a4
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353248"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="d0ac0-103">Criar funções personalizadas no Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="d0ac0-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="d0ac0-104">Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="d0ac0-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="d0ac0-106">Este artigo descreve como criar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="d0ac0-107">A ilustração a seguir mostra um usuário final inserindo uma função personalizada em uma célula de uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="d0ac0-108">A função personalizada `CONTOSO.ADD42` foi projetada para adicionar 42 ao par dos números que o usuário especifica como parâmetros de entrada para a função.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="d0ac0-109">O código a seguir define a função personalizada `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="d0ac0-110">A seção [Problemas conhecidos](#known-issues) neste artigo especifica as atuais limitações de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="d0ac0-111">Componentes de um projeto de suplemento de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d0ac0-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="d0ac0-112">Se você usar o [gerador Yo Office](https://github.com/OfficeDev/generator-office) para criar funções personalizadas em um projeto do Excel, você encontrará que cria os arquivos que controlam as funções, o painel de tarefas e o suplemento geral.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="d0ac0-113">Vamos nos concentrar em arquivos que são importantes para funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="d0ac0-113">We'll concentrate on the files that are important to custom functions:</span></span> 

| <span data-ttu-id="d0ac0-114">File</span><span class="sxs-lookup"><span data-stu-id="d0ac0-114">File</span></span> | <span data-ttu-id="d0ac0-115">Formato de arquivo</span><span class="sxs-lookup"><span data-stu-id="d0ac0-115">File format</span></span> | <span data-ttu-id="d0ac0-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="d0ac0-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="d0ac0-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="d0ac0-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="d0ac0-118">ou</span><span class="sxs-lookup"><span data-stu-id="d0ac0-118">or</span></span><br/><span data-ttu-id="d0ac0-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="d0ac0-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="d0ac0-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="d0ac0-120">JavaScript</span></span><br/><span data-ttu-id="d0ac0-121">ou</span><span class="sxs-lookup"><span data-stu-id="d0ac0-121">or</span></span><br/><span data-ttu-id="d0ac0-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="d0ac0-122">TypeScript</span></span> | <span data-ttu-id="d0ac0-123">Contém o código que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="d0ac0-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="d0ac0-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="d0ac0-125">HTML</span><span class="sxs-lookup"><span data-stu-id="d0ac0-125">HTML</span></span> | <span data-ttu-id="d0ac0-126">Fornece uma referência&lt;script&gt;ao arquivo JavaScript que define funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="d0ac0-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="d0ac0-127">**./manifest.xml**</span></span> | <span data-ttu-id="d0ac0-128">XML</span><span class="sxs-lookup"><span data-stu-id="d0ac0-128">XML</span></span> | <span data-ttu-id="d0ac0-129">Especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos JavaScript e HTML listados anteriormente nesta tabela.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="d0ac0-130">Também lista os locais de outros arquivos, que o suplemento pode fazer uso, como os arquivos do painel de tarefas e arquivos de comando.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="d0ac0-131">Arquivo de script</span><span class="sxs-lookup"><span data-stu-id="d0ac0-131">Script file</span></span>

<span data-ttu-id="d0ac0-132">O arquivo de script (**./src/functions/functions.js** ou **./src/functions/functions.ts** no projeto que o gerador Yo Office cria) contém o código que define funções personalizadas, comentários que definem a função, e associa os nomes das funções personalizadas a objetos no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="d0ac0-133">Por exemplo, o código a seguir define funções personalizadas `add` e especifica as informações de mapeamento para as duas funções.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-133">The following code defines the custom function `add`  and then specifies association information for the function.</span></span> <span data-ttu-id="d0ac0-134">Para saber mais, confira [práticas recomendadas de funções personalizados](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-134">For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

<span data-ttu-id="d0ac0-135">O código a seguir também fornece os comentários de código que definem a função.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-135">The following code also provides code comments which define the function.</span></span> <span data-ttu-id="d0ac0-136">O necessário `@customfunction` comentário é declarado primeiro, para indicar que se trata de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-136">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="d0ac0-137">Além disso, observe que dois parâmetros foram declarados, `first` e `second`, que é seguido por suas `description` propriedades.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-137">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="d0ac0-138">Por fim, uma `returns` descrição é fornecida.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-138">Finally, a `returns` description is given.</span></span> <span data-ttu-id="d0ac0-139">Para obter mais informações sobre quais comentários são necessários para sua função personalizada, confira [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-139">For more information about what comments are required for your custom function, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a><span data-ttu-id="d0ac0-140">Arquivo de manifesto</span><span class="sxs-lookup"><span data-stu-id="d0ac0-140">Manifest file</span></span>

<span data-ttu-id="d0ac0-141">O arquivo de manifesto XML para um suplemento que define funções personalizadas (**./manifest.xml** no projeto gerador que Yo Office cria) especifica o namespace para todas as funções personalizadas no suplemento e o local dos arquivos HTML, JavaScript e JSON.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-141">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="d0ac0-142">A marcação XML a seguir mostra um exemplo dos elementos `<ExtensionPoint>` e `<Resources>` que você deve incluir no manifesto de um suplemento para habilitar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-142">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="d0ac0-143">Se estiver usando o gerador Yo Office, seus arquivos de funções personalizadas gerados conterão um arquivo de manifesto mais complexo, que você pode comparar neste [repositório do Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-143">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="d0ac0-144">As URLs especificadas no arquivo de manifesto para as funções personalizadas JavaScript e JSON e arquivos HTML devem estar publicamente acessíveis e ter o mesmo subdomínio.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-144">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
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
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="d0ac0-145">Funções do Excel são anexadas ao namespace especificado no seu arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-145">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="d0ac0-146">O namespace da função vem antes do nome da função e são separados por um ponto.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-146">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="d0ac0-147">Por exemplo, para acionar a função`ADD42` na célula de uma planilha do Excel, você digitaria `=CONTOSO.ADD42`, porque `CONTOSO` é o namespace e `ADD42` é o nome da função especificada no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-147">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="d0ac0-148">O namespace deve ser usado como identificador para o as sua empresa ou suplemento.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-148">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="d0ac0-149">Um namespace pode conter apenas caracteres alfanuméricos e períodos.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-149">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="d0ac0-150">Como declarar uma função volátil</span><span class="sxs-lookup"><span data-stu-id="d0ac0-150">Declaring a volatile function</span></span>

<span data-ttu-id="d0ac0-151">As [funções voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) são funções nas quais o valor muda de momento a momento, mesmo que nenhum dos argumentos da função tenha mudado.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-151">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="d0ac0-152">Essas funções são recalculadas sempre que o Excel recalcular.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-152">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="d0ac0-153">Por exemplo, imagine uma célula que chame a função `NOW`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-153">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="d0ac0-154">Toda vez que `NOW` for chamado, retornará automaticamente a data e a hora atuais.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-154">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="d0ac0-155">O Excel contém várias funções voláteis internas, como `RAND` e `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-155">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="d0ac0-156">Para ver uma lista mais completa de funções voláteis do Excel, confira [Funções voláteis e não voláteis](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-156">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="d0ac0-157">As funções personalizadas permitem que você crie suas próprias funções voláteis, que podem ser úteis ao lidar com datas, horas, números aleatórios e modelagem.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-157">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="d0ac0-158">Por exemplo, as simulações de Monte Carlo exigem a geração de entradas aleatórias para determinar uma solução ideal.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-158">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="d0ac0-159">Para declarar uma função volátil, adicione `"volatile": true` no objeto `options` para a função no arquivo JSON de metadados, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-159">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="d0ac0-160">Observe que uma função não pode ser marcada como `"streaming": true` e `"volatile": true`; em casos em que ambas estejam marcadas com `true`, a opção volátil será ignorada.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-160">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="d0ac0-161">Salvar e compartilhar estado</span><span class="sxs-lookup"><span data-stu-id="d0ac0-161">Saving and sharing state</span></span>

<span data-ttu-id="d0ac0-162">Funções personalizadas podem salvar os dados em variáveis, que podem ser usadas em chamadas subsequentes.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-162">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="d0ac0-163">O estado salvo é útil quando os usuários solicitam a mesma função personalizada usando mais de uma célula, porque todas as ocorrências da função podem acessar o estado.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-163">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="d0ac0-164">Por exemplo, você pode salvar os dados retornados de uma chamada para um recurso da Web para evitar fazer chamadas adicionais para o mesmo recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-164">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="d0ac0-165">O código a seguir mostra uma implementação da função de streaming de temperatura que salva o estado globalmente.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-165">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="d0ac0-166">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="d0ac0-166">Note the following about this code:</span></span>

- <span data-ttu-id="d0ac0-167">A função `streamTemperature` atualiza o valor de temperatura exibido na célula a cada segundo e ele usa a variável `savedTemperatures` como fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-167">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="d0ac0-168">Como `streamTemperature` é uma função de streaming, ela implementa um identificador de cancelamento que será executado quando a função for cancelada.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-168">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="d0ac0-169">Se um usuário ligar a função`streamTemperature` de várias células no Excel, a função `streamTemperature` lê os dados a partir da mesma`savedTemperatures` variável toda vez que ela for executada.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-169">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="d0ac0-170">`refreshTemperature` é uma função de streaming que lê a temperatura de um determinado termômetro a cada segundo e armazena o resultado na variável`savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-170">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="d0ac0-171">Como a função`refreshTemperature` não é exibida para os usuários finais no Excel, não é necessário ser registrado no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-171">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="coauthoring"></a><span data-ttu-id="d0ac0-172">Coautoria</span><span class="sxs-lookup"><span data-stu-id="d0ac0-172">Coauthoring</span></span>

<span data-ttu-id="d0ac0-173">O Excel Online e o Excel para Windows com uma assinatura do Office 365 permitem editar documentos em coautoria, e esse recurso funciona com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-173">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="d0ac0-174">Se a pasta de trabalho usa uma função personalizada, seu colega será solicitado a carregar o suplemento da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-174">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="d0ac0-175">Depois de carregarem o suplemento, a função personalizada compartilhará resultados por meio de coautoria.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-175">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="d0ac0-176">Para saber mais sobre coautoria, confira o tópico [Sobre o recurso de coautoria no Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-176">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="d0ac0-177">Trabalhar com intervalos de dados</span><span class="sxs-lookup"><span data-stu-id="d0ac0-177">Working with ranges of data</span></span>

<span data-ttu-id="d0ac0-178">Sua função personalizada pode aceitar um intervalo de dados como um parâmetro de entrada ou pode retornar um intervalo de dados.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-178">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="d0ac0-179">Em JavaScript, um intervalo de dados é representado como uma matriz bidimensional.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-179">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="d0ac0-180">Por exemplo, suponha que sua função retorne o segundo maior valor de um intervalo de números armazenados no Excel.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-180">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="d0ac0-181">A função a seguir aceita o parâmetro `values`, que é um tipo de parâmetro `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-181">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="d0ac0-182">Observe que, nos metadados JSON dessa função, você deve definir o parâmetro `type` propriedade para `matrix`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-182">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="d0ac0-183">Determinar quais células chamadas de sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="d0ac0-183">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="d0ac0-184">Em alguns casos, você precisará obter o endereço da célula invocada na sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-184">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="d0ac0-185">Isso pode ser útil para os seguintes tipos de cenários:</span><span class="sxs-lookup"><span data-stu-id="d0ac0-185">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="d0ac0-186">Formatação de intervalos: Use o endereço da célula como a chave para armazenar informações em [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-186">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="d0ac0-187">Em seguida, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) no Excel para carregar a chave de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-187">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="d0ac0-188">Exibição de valores armazenados em cache: se sua função for usada offline, exiba valores armazenados em cache de `AsyncStorage` usando `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-188">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="d0ac0-189">Reconciliação: Use o endereço da célula para descobrir uma célula de origem para ajudá-lo a reconciliar onde o processamento está ocorrendo.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-189">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="d0ac0-190">As informações sobre o endereço de uma célula serão expostas somente se `requiresAddress` estiver marcado como `true` no arquivo de metadados JSON da função.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-190">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="d0ac0-191">A seguir, um exemplo disso para se você fosse gravar esse arquivo JSON manualmente.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-191">The following sample gives an example of this if you were to write this JSON file by hand.</span></span> <span data-ttu-id="d0ac0-192">Você também pode usar a tag `@requiresAddress` se gerar automaticamente seu arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-192">You can also use the `@requiresAddress` tag if automatically generating your JSON file.</span></span> <span data-ttu-id="d0ac0-193">Para mais detalhes, confira [Geração automática do JSON](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-193">For more details, see [JSON Autogeneration](custom-functions-json-autogeneration.md).</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="d0ac0-194">No arquivo de script (**./src/customfunctions.js** ou **./src/customfunctions.ts**), também será necessário adicionar uma função `getAddress` para encontrar o endereço de uma célula.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-194">In the script file (**./src/functions/functions.js** or **./src/functions/functions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="d0ac0-195">Essa função pode ter parâmetros, conforme mostrado no exemplo a seguir como `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-195">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="d0ac0-196">O último parâmetro sempre será `invocationContext`, um objeto com o local da célula que o Excel passa quando `requiresAddress` é marcado como `true` no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-196">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="d0ac0-197">Por padrão, os valores retornados de uma função `getAddress` seguem o formato abaixo: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-197">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="d0ac0-198">Por exemplo, se uma função foi chamada de uma planilha nomeada Despesas na célula B2, o valor retornado seria `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="d0ac0-198">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="d0ac0-199">Problemas conhecidos</span><span class="sxs-lookup"><span data-stu-id="d0ac0-199">Known issues</span></span>

<span data-ttu-id="d0ac0-200">Veja os problemas conhecidos no nosso [GitHub de funções do Excel personalizado repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-200">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="d0ac0-201">Confira também</span><span class="sxs-lookup"><span data-stu-id="d0ac0-201">See also</span></span>

* [<span data-ttu-id="d0ac0-202">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d0ac0-202">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d0ac0-203">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="d0ac0-203">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="d0ac0-204">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="d0ac0-204">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="d0ac0-205">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d0ac0-205">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="d0ac0-206">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="d0ac0-206">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="d0ac0-207">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d0ac0-207">Custom functions debugging</span></span>](custom-functions-debugging.md)
