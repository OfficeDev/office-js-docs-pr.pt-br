---
title: Conceitos avançados de programação com a API JavaScript do Excel
description: Saiba como um suplemento do Excel interage com objetos no Excel usando modelos de objeto da API JavaScript para Office.
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: 81602f48231f20b50a454134bc789dfdee2bbc12
ms.sourcegitcommit: 4f2f1c0a8ee777a43bb28efa226684261f4c4b9f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081393"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="e8b36-103">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e8b36-103">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="e8b36-104">Este artigo se baseia nas informações contidas em [conceitos fundamentais de programação API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avançados que são essenciais para a criação de suplementos complexos para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="e8b36-104">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="e8b36-105">APIs Office.js para Excel</span><span class="sxs-lookup"><span data-stu-id="e8b36-105">Office.js APIs for Excel</span></span>

<span data-ttu-id="e8b36-106">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="e8b36-106">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="e8b36-107">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="e8b36-107">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="e8b36-108">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="e8b36-108">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="e8b36-109">Enquanto você provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, você também usará objetos na API comum.</span><span class="sxs-lookup"><span data-stu-id="e8b36-109">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="e8b36-110">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="e8b36-110">For example:</span></span>

- <span data-ttu-id="e8b36-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span><span class="sxs-lookup"><span data-stu-id="e8b36-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="e8b36-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span><span class="sxs-lookup"><span data-stu-id="e8b36-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="e8b36-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span><span class="sxs-lookup"><span data-stu-id="e8b36-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="e8b36-114">[Documento](/javascript/api/office/office.document): o objeto `Document` fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo do Excel em que o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="e8b36-114">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="e8b36-115">A imagem a seguir ilustra quando você pode usar a API JavaScript do Excel ou as APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="e8b36-115">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Imagem das diferentes entre a API JS do Excel e as APIs comuns](../images/excel-js-api-common-api.png)

## <a name="requirement-sets"></a><span data-ttu-id="e8b36-117">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="e8b36-117">Requirement sets</span></span>

<span data-ttu-id="e8b36-118">Requirement sets are named groups of API members.</span><span class="sxs-lookup"><span data-stu-id="e8b36-118">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="e8b36-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span><span class="sxs-lookup"><span data-stu-id="e8b36-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="e8b36-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e8b36-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="e8b36-121">Verificando o suporte ao conjunto de requisitos no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="e8b36-121">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="e8b36-122">O exemplo de código a seguir mostra como determinar se o aplicativo host, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.</span><span class="sxs-lookup"><span data-stu-id="e8b36-122">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="e8b36-123">Definindo o suporte ao conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="e8b36-123">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="e8b36-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span><span class="sxs-lookup"><span data-stu-id="e8b36-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="e8b36-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span><span class="sxs-lookup"><span data-stu-id="e8b36-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="e8b36-126">O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos host do Office que dão suporte ao conjunto de requisitos ExcelApi, versão 1.3 ou superior.</span><span class="sxs-lookup"><span data-stu-id="e8b36-126">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="e8b36-127">Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel Online, Windows e iPad, é recomendável verificar o suporte a requisitos no tempo de execução, em vez de definir o suporte ao conjunto de requisitos no manifesto.</span><span class="sxs-lookup"><span data-stu-id="e8b36-127">To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="e8b36-128">Conjuntos de requisitos para a API comum Office.js</span><span class="sxs-lookup"><span data-stu-id="e8b36-128">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="e8b36-129">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e8b36-129">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="e8b36-130">Carregando as propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-130">Loading the properties of an object</span></span>

<span data-ttu-id="e8b36-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span><span class="sxs-lookup"><span data-stu-id="e8b36-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="e8b36-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span><span class="sxs-lookup"><span data-stu-id="e8b36-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

### <a name="method-details"></a><span data-ttu-id="e8b36-133">Detalhes do método</span><span class="sxs-lookup"><span data-stu-id="e8b36-133">Method details</span></span>

#### `load(propertyNames?: string | string[])`

<span data-ttu-id="e8b36-134">Coloca um comando na fila para carregar as propriedades especificadas do objeto.</span><span class="sxs-lookup"><span data-stu-id="e8b36-134">Queues up a command to load the specified properties of the object.</span></span> <span data-ttu-id="e8b36-135">Você deve chamar `context.sync()` antes de ler as propriedades.</span><span class="sxs-lookup"><span data-stu-id="e8b36-135">You must call `context.sync()` before reading the properties.</span></span>

#### <a name="syntax"></a><span data-ttu-id="e8b36-136">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e8b36-136">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="e8b36-137">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e8b36-137">Parameters</span></span>

|<span data-ttu-id="e8b36-138">**Parâmetro**</span><span class="sxs-lookup"><span data-stu-id="e8b36-138">**Parameter**</span></span>|<span data-ttu-id="e8b36-139">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="e8b36-139">**Type**</span></span>|<span data-ttu-id="e8b36-140">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="e8b36-140">**Description**</span></span>|
|:------------|:-------|:----------|
|`propertyNames`|<span data-ttu-id="e8b36-141">objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-141">object</span></span>|<span data-ttu-id="e8b36-142">Opcional.</span><span class="sxs-lookup"><span data-stu-id="e8b36-142">Optional.</span></span> <span data-ttu-id="e8b36-143">Aceita nomes de propriedade como uma matriz ou cadeia de caracteres delimitada por vírgulas.</span><span class="sxs-lookup"><span data-stu-id="e8b36-143">Accepts property names as comma-delimited string or an array.</span></span>|

#### <a name="returns"></a><span data-ttu-id="e8b36-144">Retorna</span><span class="sxs-lookup"><span data-stu-id="e8b36-144">Returns</span></span>

<span data-ttu-id="e8b36-145">nulo</span><span class="sxs-lookup"><span data-stu-id="e8b36-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="e8b36-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e8b36-146">Example</span></span>

<span data-ttu-id="e8b36-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span><span class="sxs-lookup"><span data-stu-id="e8b36-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="e8b36-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span><span class="sxs-lookup"><span data-stu-id="e8b36-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="e8b36-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span><span class="sxs-lookup"><span data-stu-id="e8b36-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="e8b36-150">Carregar propriedades de opção</span><span class="sxs-lookup"><span data-stu-id="e8b36-150">Load option properties</span></span>

<span data-ttu-id="e8b36-151">Como uma alternativa para passar uma cadeia de caracteres delimitada por vírgulas ou uma matriz ao chamar o método `load()`, você pode passar um objeto que contém as propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e8b36-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="e8b36-152">**Propriedade**</span><span class="sxs-lookup"><span data-stu-id="e8b36-152">**Property**</span></span>|<span data-ttu-id="e8b36-153">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="e8b36-153">**Type**</span></span>|<span data-ttu-id="e8b36-154">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="e8b36-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="e8b36-155">objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-155">object</span></span>|<span data-ttu-id="e8b36-156">Contains a comma-delimited list or an array of scalar property names.</span><span class="sxs-lookup"><span data-stu-id="e8b36-156">Contains a comma-delimited list or an array of scalar property names.</span></span> <span data-ttu-id="e8b36-157">Optional.</span><span class="sxs-lookup"><span data-stu-id="e8b36-157">Optional.</span></span>|
|`expand`|<span data-ttu-id="e8b36-158">objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-158">object</span></span>|<span data-ttu-id="e8b36-159">Contains a comma-delimited list or an array of navigational property names.</span><span class="sxs-lookup"><span data-stu-id="e8b36-159">Contains a comma-delimited list or an array of navigational property names.</span></span> <span data-ttu-id="e8b36-160">Optional.</span><span class="sxs-lookup"><span data-stu-id="e8b36-160">Optional.</span></span>|
|`top`|<span data-ttu-id="e8b36-161">int</span><span class="sxs-lookup"><span data-stu-id="e8b36-161">int</span></span>| <span data-ttu-id="e8b36-162">Specifies the maximum number of collection items that can be included in the result.</span><span class="sxs-lookup"><span data-stu-id="e8b36-162">Specifies the maximum number of collection items that can be included in the result.</span></span> <span data-ttu-id="e8b36-163">Optional.</span><span class="sxs-lookup"><span data-stu-id="e8b36-163">Optional.</span></span> <span data-ttu-id="e8b36-164">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="e8b36-164">You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="e8b36-165">int</span><span class="sxs-lookup"><span data-stu-id="e8b36-165">int</span></span>|<span data-ttu-id="e8b36-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span><span class="sxs-lookup"><span data-stu-id="e8b36-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span></span> <span data-ttu-id="e8b36-167">If `top` is specified, the result set will start after skipping the specified number of items.</span><span class="sxs-lookup"><span data-stu-id="e8b36-167">If `top` is specified, the result set will start after skipping the specified number of items.</span></span> <span data-ttu-id="e8b36-168">Optional.</span><span class="sxs-lookup"><span data-stu-id="e8b36-168">Optional.</span></span> <span data-ttu-id="e8b36-169">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="e8b36-169">You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="e8b36-170">O exemplo de código a seguir carrega uma coleção de planilhas selecionando a `name`propriedade e o `address`do intervalo usado para cada planilha na coleção.</span><span class="sxs-lookup"><span data-stu-id="e8b36-170">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="e8b36-171">Ele também especifica que apenas as cinco planilhas principais na coleção devem ser carregadas.</span><span class="sxs-lookup"><span data-stu-id="e8b36-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="e8b36-172">Você poderia processar o próximo conjunto de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributo.</span><span class="sxs-lookup"><span data-stu-id="e8b36-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

### <a name="calling-load-without-parameters"></a><span data-ttu-id="e8b36-173">Chamando `load` sem parâmetros</span><span class="sxs-lookup"><span data-stu-id="e8b36-173">Calling `load` without parameters</span></span>

<span data-ttu-id="e8b36-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span><span class="sxs-lookup"><span data-stu-id="e8b36-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="e8b36-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span><span class="sxs-lookup"><span data-stu-id="e8b36-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e8b36-176">A quantidade de dados retornados por uma declaração `load` sem parâmetros pode exceder os limites de tamanho do serviço.</span><span class="sxs-lookup"><span data-stu-id="e8b36-176">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="e8b36-177">Para reduzir os riscos a suplementos mais antigos, algumas propriedades não são retornadas por `load` sem a solicitação explícita.</span><span class="sxs-lookup"><span data-stu-id="e8b36-177">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="e8b36-178">As seguintes propriedades são excluídas dessas operações de carregamento:</span><span class="sxs-lookup"><span data-stu-id="e8b36-178">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="e8b36-179">Propriedades escalares e de navegação</span><span class="sxs-lookup"><span data-stu-id="e8b36-179">Scalar and navigation properties</span></span>

<span data-ttu-id="e8b36-180">Há duas categorias de propriedades: **escalar** e de **navegação**.</span><span class="sxs-lookup"><span data-stu-id="e8b36-180">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="e8b36-181">As propriedades escalares são tipos atribuíveis, como cadeias de caracteres, inteiros e estruturas JSON.</span><span class="sxs-lookup"><span data-stu-id="e8b36-181">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="e8b36-182">As propriedades de navegação são objetos Somente Leitura e coleções de objetos que têm seus campos atribuídos, em vez de atribuir diretamente a propriedade.</span><span class="sxs-lookup"><span data-stu-id="e8b36-182">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="e8b36-183">Por exemplo, os membros `name` e `position` no objeto [Planilha](/javascript/api/excel/excel.worksheet) são propriedades escalares, enquanto `protection` e `tables` são propriedades de navegação.</span><span class="sxs-lookup"><span data-stu-id="e8b36-183">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span> <span data-ttu-id="e8b36-184">`prompt` no objeto [DataValidation](/javascript/api/excel/excel.datavalidation) é um exemplo de uma propriedade escalar que deve ser definida usando um objeto JSON (`dv.prompt = { title: "MyPrompt"}`), em vez de definir as subpropriedades (`dv.prompt.title = "MyPrompt" // will not set the title`).</span><span class="sxs-lookup"><span data-stu-id="e8b36-184">`prompt` on the [DataValidation](/javascript/api/excel/excel.datavalidation) object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="e8b36-185">Propriedades escalares e propriedades de navegação com `object.load()`</span><span class="sxs-lookup"><span data-stu-id="e8b36-185">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="e8b36-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span><span class="sxs-lookup"><span data-stu-id="e8b36-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="e8b36-187">Additionally, navigation properties cannot be loaded directly.</span><span class="sxs-lookup"><span data-stu-id="e8b36-187">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="e8b36-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span><span class="sxs-lookup"><span data-stu-id="e8b36-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="e8b36-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span><span class="sxs-lookup"><span data-stu-id="e8b36-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="e8b36-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span><span class="sxs-lookup"><span data-stu-id="e8b36-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="e8b36-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="e8b36-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="e8b36-192">You do not need to load the property before you set it.</span><span class="sxs-lookup"><span data-stu-id="e8b36-192">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="e8b36-193">Definindo propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-193">Setting properties of an object</span></span>

<span data-ttu-id="e8b36-194">Setting properties on an object with nested navigation properties can be cumbersome.</span><span class="sxs-lookup"><span data-stu-id="e8b36-194">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="e8b36-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="e8b36-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="e8b36-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span><span class="sxs-lookup"><span data-stu-id="e8b36-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="e8b36-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="e8b36-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="e8b36-198">The common (shared) APIs do not support this method.</span><span class="sxs-lookup"><span data-stu-id="e8b36-198">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="e8b36-199">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="e8b36-199">set (properties: object, options: object)</span></span>

<span data-ttu-id="e8b36-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span><span class="sxs-lookup"><span data-stu-id="e8b36-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span></span> <span data-ttu-id="e8b36-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span><span class="sxs-lookup"><span data-stu-id="e8b36-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="e8b36-202">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e8b36-202">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="e8b36-203">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e8b36-203">Parameters</span></span>

|<span data-ttu-id="e8b36-204">**Parâmetro**</span><span class="sxs-lookup"><span data-stu-id="e8b36-204">**Parameter**</span></span>|<span data-ttu-id="e8b36-205">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="e8b36-205">**Type**</span></span>|<span data-ttu-id="e8b36-206">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="e8b36-206">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="e8b36-207">objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-207">object</span></span>|<span data-ttu-id="e8b36-208">Um objeto do mesmo tipo de objeto do Office.js no qual o método é chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o método é chamado.</span><span class="sxs-lookup"><span data-stu-id="e8b36-208">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="e8b36-209">objeto</span><span class="sxs-lookup"><span data-stu-id="e8b36-209">object</span></span>|<span data-ttu-id="e8b36-210">Optional.</span><span class="sxs-lookup"><span data-stu-id="e8b36-210">Optional.</span></span> <span data-ttu-id="e8b36-211">Can only be passed when the first parameter is a JavaScript object.</span><span class="sxs-lookup"><span data-stu-id="e8b36-211">Can only be passed when the first parameter is a JavaScript object.</span></span> <span data-ttu-id="e8b36-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span><span class="sxs-lookup"><span data-stu-id="e8b36-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="e8b36-213">Retorna</span><span class="sxs-lookup"><span data-stu-id="e8b36-213">Returns</span></span>

<span data-ttu-id="e8b36-214">nulo</span><span class="sxs-lookup"><span data-stu-id="e8b36-214">void</span></span>

#### <a name="example"></a><span data-ttu-id="e8b36-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e8b36-215">Example</span></span>

<span data-ttu-id="e8b36-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span><span class="sxs-lookup"><span data-stu-id="e8b36-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span></span> <span data-ttu-id="e8b36-217">This example assumes that there is data in range **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="e8b36-217">This example assumes that there is data in range **B2:E2**.</span></span>

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

## <a name="42ornullobject-methods"></a><span data-ttu-id="e8b36-218">Métodos &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="e8b36-218">&#42;OrNullObject methods</span></span>

<span data-ttu-id="e8b36-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span><span class="sxs-lookup"><span data-stu-id="e8b36-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="e8b36-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span><span class="sxs-lookup"><span data-stu-id="e8b36-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="e8b36-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="e8b36-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="e8b36-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span><span class="sxs-lookup"><span data-stu-id="e8b36-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="e8b36-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span><span class="sxs-lookup"><span data-stu-id="e8b36-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="e8b36-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span><span class="sxs-lookup"><span data-stu-id="e8b36-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="e8b36-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span><span class="sxs-lookup"><span data-stu-id="e8b36-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="e8b36-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span><span class="sxs-lookup"><span data-stu-id="e8b36-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="e8b36-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span><span class="sxs-lookup"><span data-stu-id="e8b36-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e8b36-228">Confira também</span><span class="sxs-lookup"><span data-stu-id="e8b36-228">See also</span></span>

* [<span data-ttu-id="e8b36-229">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e8b36-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="e8b36-230">Exemplos de código de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="e8b36-230">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="e8b36-231">Otimização de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e8b36-231">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="e8b36-232">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e8b36-232">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
