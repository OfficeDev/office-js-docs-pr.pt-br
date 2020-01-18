---
title: Conceitos avançados de programação com a API JavaScript do Excel
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 76308b6ce04dfcaa09e9006373caf07744572112
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217335"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="2a2c8-102">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2a2c8-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="2a2c8-103">Este artigo se baseia nas informações contidas em [conceitos fundamentais de programação API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avançados que são essenciais para a criação de suplementos complexos para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-103">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="2a2c8-104">APIs Office.js para Excel</span><span class="sxs-lookup"><span data-stu-id="2a2c8-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="2a2c8-105">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:</span><span class="sxs-lookup"><span data-stu-id="2a2c8-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="2a2c8-106">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="2a2c8-107">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="2a2c8-108">Enquanto você provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, você também usará objetos na API comum.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-108">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="2a2c8-109">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="2a2c8-109">For example:</span></span>

- <span data-ttu-id="2a2c8-p102">[Contexto](/javascript/api/office/office.context): o objeto **Context** representa o ambiente de tempo de execução do suplemento e oferece acesso aos principais objetos da API. Ele consiste em detalhes da configuração da pasta de trabalho, como `contentLanguage` e `officeTheme`, além de fornecer informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se o conjunto de requisitos especificado é suportado pelo aplicativo Excel onde o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p102">[Context](/javascript/api/office/office.context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="2a2c8-113">[Document](/javascript/api/office/office.document): O objeto **Document** fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo Excel onde o suplemento está em execução.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-113">[Document](/javascript/api/office/office.document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="2a2c8-114">A imagem a seguir ilustra quando você pode usar a API JavaScript do Excel ou as APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-114">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Imagem das diferentes entre a API JS do Excel e as APIs comuns](../images/excel-js-api-common-api.png)

## <a name="requirement-sets"></a><span data-ttu-id="2a2c8-116">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="2a2c8-116">Requirement sets</span></span>

<span data-ttu-id="2a2c8-p103">Os conjuntos de requisitos são grupos nomeados de membros da API. Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um host do Office dá suporte às APIs necessárias ao suplemento. Para identificar os conjuntos de requisitos específicos que estão disponíveis em cada plataforma suportada, confira [Conjuntos de requisitos da API JavaScript do Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p103">Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="2a2c8-120">Verificando o suporte ao conjunto de requisitos no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="2a2c8-120">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="2a2c8-121">O exemplo de código a seguir mostra como determinar se o aplicativo host, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-121">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="2a2c8-122">Definindo o suporte ao conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-122">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="2a2c8-p104">Você pode usar o [elemento Requirements](/office/dev/add-ins/reference/manifest/requirements) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado. Se a plataforma ou o host do Office não der suporte aos conjuntos de requisitos ou aos métodos de API que são especificados no elemento **Requirements** do manifesto, o suplemento não será executado nesse host ou plataforma e não será exibido na lista de suplementos que são mostrados em **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p104">You can use the [Requirements element](/office/dev/add-ins/reference/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="2a2c8-125">O exemplo de código a seguir mostra o elemento **Requirements** em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos host do Office que dão suporte ao conjunto de requisitos ExcelApi, versão 1.3 ou superior.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-125">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="2a2c8-126">Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel Online, Windows e iPad, é recomendável verificar o suporte a requisitos no tempo de execução, em vez de definir o suporte ao conjunto de requisitos no manifesto.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-126">To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="2a2c8-127">Conjuntos de requisitos para a API comum Office.js</span><span class="sxs-lookup"><span data-stu-id="2a2c8-127">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="2a2c8-128">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2a2c8-128">For information about Common API requirement sets, see [Office Common API requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="2a2c8-129">Carregando as propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-129">Loading the properties of an object</span></span>

<span data-ttu-id="2a2c8-p105">Chamar o método `load()` em um objeto JavaScript do Excel orienta a API a carregar o objeto na memória do JavaScript quando o método `sync()` é executado. O método `load()` aceita uma cadeia de caracteres que contém nomes de propriedades delimitados por vírgulas a serem carregados ou um objeto que especifica propriedades a serem carregadas, opções de paginação, etc.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p105">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs. The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

> [!NOTE]
> <span data-ttu-id="2a2c8-p106">Se você chamar o método `load()` em um objeto (ou uma coleção) sem especificar qualquer parâmetro, todas as propriedades escalares do objeto (ou todas as propriedades escalares de todos os objetos na coleção) serão carregadas. Para reduzir a quantidade de transferência de dados entre o aplicativo host e o suplemento do Excel, você deve evitar chamar o método `load()` sem especificar explicitamente quais propriedades carregar.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p106">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="2a2c8-134">Detalhes do método</span><span class="sxs-lookup"><span data-stu-id="2a2c8-134">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="2a2c8-135">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="2a2c8-135">load(param: object)</span></span>

<span data-ttu-id="2a2c8-136">Preenche o objeto proxy criado na camada JavaScript com os valores da propriedade e do objeto especificados pelos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-136">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="2a2c8-137">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2a2c8-137">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="2a2c8-138">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2a2c8-138">Parameters</span></span>

|<span data-ttu-id="2a2c8-139">**Parâmetro**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-139">**Parameter**</span></span>|<span data-ttu-id="2a2c8-140">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-140">**Type**</span></span>|<span data-ttu-id="2a2c8-141">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-141">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="2a2c8-142">objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-142">object</span></span>|<span data-ttu-id="2a2c8-p107">Opcional. Aceita nomes de propriedade como cadeia de caracteres delimitada por vírgula ou uma matriz. Também é possível passar um objeto para definir as propriedades da seleção e de navegação (conforme mostrado no exemplo abaixo).</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p107">Optional. Accepts property names as comma-delimited string or an array. An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="2a2c8-146">Retorna</span><span class="sxs-lookup"><span data-stu-id="2a2c8-146">Returns</span></span>

<span data-ttu-id="2a2c8-147">nulo</span><span class="sxs-lookup"><span data-stu-id="2a2c8-147">void</span></span>

#### <a name="example"></a><span data-ttu-id="2a2c8-148">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2a2c8-148">Example</span></span>

<span data-ttu-id="2a2c8-p108">O exemplo de código a seguir define as propriedades de um intervalo do Excel, copiando as propriedades de outro intervalo. Observe que o objeto de origem deve ser carregado primeiro para que seus valores de propriedade possam ser acessados e gravados no intervalo de destino. Este exemplo pressupõe que há dados nos dois intervalos (**B2:E2** e **B7:E7**) e que os dois intervalos são inicialmente formatados de modo diferente.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p108">The following code sample sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first, before its property values can be accessed and written to the target range. This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a><span data-ttu-id="2a2c8-152">Carregar propriedades de opção</span><span class="sxs-lookup"><span data-stu-id="2a2c8-152">Load option properties</span></span>

<span data-ttu-id="2a2c8-153">Como uma alternativa para passar uma cadeia de caracteres delimitada por vírgulas ou uma matriz ao chamar o método `load()`, você pode passar um objeto que contém as propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-153">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="2a2c8-154">**Propriedade**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-154">**Property**</span></span>|<span data-ttu-id="2a2c8-155">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-155">**Type**</span></span>|<span data-ttu-id="2a2c8-156">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-156">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="2a2c8-157">objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-157">object</span></span>|<span data-ttu-id="2a2c8-p109">Inclui uma lista delimitada por vírgula ou uma matriz de nomes de propriedade. Opcional.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p109">Contains a comma-delimited list or an array of scalar property names. Optional.</span></span>|
|`expand`|<span data-ttu-id="2a2c8-160">objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-160">object</span></span>|<span data-ttu-id="2a2c8-p110">Inclui uma lista delimitada por vírgula ou uma matriz de nomes de propriedade de navegação. Opcional.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p110">Contains a comma-delimited list or an array of navigational property names. Optional.</span></span>|
|`top`|<span data-ttu-id="2a2c8-163">int</span><span class="sxs-lookup"><span data-stu-id="2a2c8-163">int</span></span>| <span data-ttu-id="2a2c8-p111">Especifica o número máximo de itens da coleção que podem ser incluídos no resultado. Opcional. Você só pode usar essa opção quando usar a opção de notação de objeto.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="2a2c8-167">int</span><span class="sxs-lookup"><span data-stu-id="2a2c8-167">int</span></span>|<span data-ttu-id="2a2c8-p112">Determina o número de itens da coleção que devem ser ignorados e não incluídos no resultado. Quando a propriedade `top` for especificada, o conjunto de resultados será iniciado depois de ignorar o número de itens especificado. Opcional. Você só pode usar esta opção ao usar a opção de notação de objeto.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="2a2c8-172">O exemplo de código a seguir carrega uma coleção de planilhas selecionando a `name`propriedade e o `address`do intervalo usado para cada planilha na coleção.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-172">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="2a2c8-173">Ele também especifica que apenas as cinco planilhas principais na coleção devem ser carregadas.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-173">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="2a2c8-174">Você poderia processar o próximo conjunto de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributo.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-174">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="2a2c8-175">Propriedades escalares e de navegação</span><span class="sxs-lookup"><span data-stu-id="2a2c8-175">Scalar and navigation properties</span></span>

<span data-ttu-id="2a2c8-176">Há duas categorias de propriedades: **escalar** e de **navegação**.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-176">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="2a2c8-177">As propriedades escalares são tipos atribuíveis, como cadeias de caracteres, inteiros e estruturas JSON.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-177">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="2a2c8-178">As propriedades de navegação são objetos Somente Leitura e coleções de objetos que têm seus campos atribuídos, em vez de atribuir diretamente a propriedade.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-178">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="2a2c8-179">Por exemplo, os membros `name` e `position` no objeto [Planilha](/javascript/api/excel/excel.worksheet) são propriedades escalares, enquanto `protection` e `tables` são propriedades de navegação.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-179">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span> <span data-ttu-id="2a2c8-180">`prompt` no objeto [DataValidation](/javascript/api/excel/excel.datavalidation) é um exemplo de uma propriedade escalar que deve ser definida usando um objeto JSON (`dv.prompt = { title: "MyPrompt"}`), em vez de definir as subpropriedades (`dv.prompt.title = "MyPrompt" // will not set the title`).</span><span class="sxs-lookup"><span data-stu-id="2a2c8-180">`prompt` on the [DataValidation](/javascript/api/excel/excel.datavalidation) object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="2a2c8-181">Propriedades escalares e propriedades de navegação com `object.load()`</span><span class="sxs-lookup"><span data-stu-id="2a2c8-181">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="2a2c8-p115">Chamar o método `object.load()` sem parâmetros especificados carregará todas as propriedades escalares do objeto; as propriedades de navegação do objeto não serão carregadas. Além disso, as propriedades de navegação não podem ser carregadas diretamente. Em vez disso, você deve usar o método `load()` para fazer referência às propriedades escalares individuais na propriedade de navegação desejada. Por exemplo, para carregar o nome da fonte de um intervalo, você deve especificar as propriedades de navegação **format** e **font** como o caminho para a propriedade **name**:</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p115">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded. Additionally, navigation properties cannot be loaded directly. Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property. For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="2a2c8-p116">com a API JavaScript do Excel, é possível definir propriedades escalares de uma propriedade de navegação percorrendo o caminho. Por exemplo, é possível definir o tamanho da fonte de um intervalo usando `someRange.format.font.size = 10;`. Não é necessário carregar a propriedade antes de configurá-la.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p116">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path. For example, you could set the font size for a range by using `someRange.format.font.size = 10;`. You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="2a2c8-189">Definindo propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-189">Setting properties of an object</span></span>

<span data-ttu-id="2a2c8-p117">A definição de propriedades em um objeto com propriedades de navegação aninhadas pode ser uma tarefa complicada. Como uma alternativa para definir propriedades individuais usando caminhos de navegação, conforme descrito acima, você pode usar o método `object.set()` que está disponível em todos os objetos na API JavaScript do Excel. Com esse método, é possível definir várias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que são estruturadas, como as propriedades do objeto no qual o método é chamado.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p117">Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="2a2c8-p118">O método `set()` é implementado apenas para objetos nas APIs JavaScript do Office específicas de host, como a API JavaScript do Excel. As APIs comuns (compartilhadas) não dão suporte a esse método.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p118">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API. The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="2a2c8-195">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="2a2c8-195">set (properties: object, options: object)</span></span>

<span data-ttu-id="2a2c8-p119">As propriedades do objeto em que o método é chamado são definidas com os mesmos valores das propriedades correspondentes do objeto transmitido. Se o parâmetro `properties` for um objeto JavaScript, as propriedades do objeto transmitido que correspondem à propriedade de somente leitura no objeto em que o método é chamado serão ignoradas ou causarão uma exceção, dependendo do parâmetro `options`.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="2a2c8-198">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2a2c8-198">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="2a2c8-199">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2a2c8-199">Parameters</span></span>

|<span data-ttu-id="2a2c8-200">**Parâmetro**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-200">**Parameter**</span></span>|<span data-ttu-id="2a2c8-201">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-201">**Type**</span></span>|<span data-ttu-id="2a2c8-202">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="2a2c8-202">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="2a2c8-203">objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-203">object</span></span>|<span data-ttu-id="2a2c8-204">Um objeto do mesmo tipo de objeto do Office.js no qual o método é chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o método é chamado.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-204">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="2a2c8-205">objeto</span><span class="sxs-lookup"><span data-stu-id="2a2c8-205">object</span></span>|<span data-ttu-id="2a2c8-p120">Opcional. Só pode ser transmitido quando o primeiro parâmetro é um objeto JavaScript. O objeto pode conter a seguinte propriedade: `throwOnReadOnly?: boolean` (O padrão é `true`: indicar um erro se o objeto JavaScript transmitido incluir propriedades de somente leitura.)</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="2a2c8-209">Retorna</span><span class="sxs-lookup"><span data-stu-id="2a2c8-209">Returns</span></span>

<span data-ttu-id="2a2c8-210">nulo</span><span class="sxs-lookup"><span data-stu-id="2a2c8-210">void</span></span>

#### <a name="example"></a><span data-ttu-id="2a2c8-211">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2a2c8-211">Example</span></span>

<span data-ttu-id="2a2c8-p121">O exemplo de código a seguir define várias propriedades do formato de um intervalo chamando o método `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto **Range**. Este exemplo supõe que há dados no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a><span data-ttu-id="2a2c8-214">Métodos &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="2a2c8-214">&#42;OrNullObject methods</span></span>

<span data-ttu-id="2a2c8-p122">Muitos métodos da API JavaScript do Excel retornarão uma exceção quando a condição da API não for atendida. Por exemplo, se você tentar obter uma planilha especificando um nome de planilha que não existe na pasta de trabalho, o método `getItem()` retornará uma exceção `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p122">Many Excel JavaScript API methods will return an exception when the condition of the API is not met. For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="2a2c8-p123">Em vez de implementar a lógica complexa de tratamento de exceção para cenários como este, você pode usar a variante do método `*OrNullObject` que está disponível para vários métodos na API JavaScript do Excel. Um método `*OrNullObject` retornará um objeto nulo (não o `null` do JavaScript), em vez de emitir uma exceção se o item especificado não existir. Por exemplo, você pode chamar o método `getItemOrNullObject()` em uma coleção, como **Worksheets**, para tentar recuperar um item da coleção. O método `getItemOrNullObject()` retornará o item especificado se ele existir; caso contrário, ele retornará um objeto nulo. O objeto nulo que é retornado contém a propriedade booliana `isNullObject`, que você pode avaliar para determinar se o objeto existe.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p123">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API. An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="2a2c8-p124">O exemplo de código a seguir tenta recuperar uma planilha chamada "Data" usando o método `getItemOrNullObject()`. Se o método retornar um objeto nulo, uma nova folha precisará ser criada para que as ações possam ser tomadas na folha.</span><span class="sxs-lookup"><span data-stu-id="2a2c8-p124">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a><span data-ttu-id="2a2c8-224">Confira também</span><span class="sxs-lookup"><span data-stu-id="2a2c8-224">See also</span></span>

* [<span data-ttu-id="2a2c8-225">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2a2c8-225">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="2a2c8-226">Exemplos de código de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="2a2c8-226">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="2a2c8-227">Otimização de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2a2c8-227">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="2a2c8-228">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2a2c8-228">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
