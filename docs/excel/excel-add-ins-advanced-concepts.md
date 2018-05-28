---
title: Conceitos avan?ados da API JavaScript do Excel
description: ''
ms.date: 1/18/2018
ms.openlocfilehash: 89db69e124475c882448a2105837787ce2c84753
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-advanced-concepts"></a><span data-ttu-id="f9975-102">Conceitos avan?ados da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f9975-102">Excel JavaScript API advanced concepts</span></span>

<span data-ttu-id="f9975-103">Este artigo foi criado com base nas informa??es em [conceitos principais da API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avan?ados que s?o essenciais para criar suplementos complexos para o Excel 2016.</span><span class="sxs-lookup"><span data-stu-id="f9975-103">This article builds upon the information in [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016.</span></span> 

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="f9975-104">APIs Office.js para Excel</span><span class="sxs-lookup"><span data-stu-id="f9975-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="f9975-105">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:</span><span class="sxs-lookup"><span data-stu-id="f9975-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="f9975-106">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) fornece objetos fortemente tipados que voc? pode usar para acessar planilhas, intervalos, tabelas, gr?ficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="f9975-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="f9975-107">**APIs comuns**: introduzidas com o Office 2013, as APIs comuns (tamb?m conhecidas como a [API Compartilhada](https://dev.office.com/reference/add-ins/javascript-api-for-office)) podem ser usadas para acessar recursos como interface de usu?rio, caixas de di?logo e configura??es de cliente, que s?o comuns entre v?rios tipos de aplicativos host, como Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f9975-107">**Common APIs**: Introduced with Office 2013, the common APIs (also referred to as the [Shared API](https://dev.office.com/reference/add-ins/javascript-api-for-office)) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="f9975-p101">Enquanto voc? provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, voc? tamb?m usar? objetos na API Compartilhada. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="f9975-p101">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016, you'll also use objects in the Shared API. For example:</span></span>

- <span data-ttu-id="f9975-110">[Contexto](https://dev.office.com/reference/add-ins/shared/context): o objeto **Context** representa o ambiente de tempo de execu??o do suplemento e oferece acesso aos principais objetos da API.</span><span class="sxs-lookup"><span data-stu-id="f9975-110">[Context](https://dev.office.com/reference/add-ins/shared/context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="f9975-111">Ele consiste em detalhes da configura??o da pasta de trabalho, como `contentLanguage` e `officeTheme`, al?m de fornecer informa??es sobre o ambiente de tempo de execu??o do suplemento, como `host` e `platform`.</span><span class="sxs-lookup"><span data-stu-id="f9975-111">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="f9975-112">Al?m disso, ele fornece o m?todo `requirements.isSetSupported()`, que voc? pode usar para verificar se o conjunto de requisitos especificado ? suportado pelo aplicativo Excel onde o suplemento est? sendo executado.</span><span class="sxs-lookup"><span data-stu-id="f9975-112">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="f9975-113">[Document](https://dev.office.com/reference/add-ins/shared/document): O objeto **Document** fornece o m?todo `getFileAsync()`, que voc? pode usar para baixar o arquivo Excel onde o suplemento est? em execu??o.</span><span class="sxs-lookup"><span data-stu-id="f9975-113">[Document](https://dev.office.com/reference/add-ins/shared/document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="f9975-114">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="f9975-114">Requirement sets</span></span>

<span data-ttu-id="f9975-115">Os conjuntos de requisitos s?o grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="f9975-115">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="f9975-116">Um Suplemento do Office pode executar uma verifica??o de tempo de execu??o ou usar conjuntos de requisitos especificados no manifesto para determinar se um host do Office d? suporte ?s APIs necess?rias ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="f9975-116">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="f9975-117">Para identificar os conjuntos de requisitos espec?ficos que est?o dispon?veis em cada plataforma suportada, confira [Conjuntos de requisitos da API JavaScript do Excel](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="f9975-117">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="f9975-118">Verificando o suporte ao conjunto de requisitos no tempo de execu??o</span><span class="sxs-lookup"><span data-stu-id="f9975-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="f9975-119">O exemplo de c?digo a seguir mostra como determinar se o aplicativo host, onde o suplemento est? em execu??o, d? suporte ao conjunto de requisitos da API especificado.</span><span class="sxs-lookup"><span data-stu-id="f9975-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="f9975-120">Definindo o suporte ao conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="f9975-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="f9975-121">Voc? pode usar o [elemento Requirements](https://dev.office.com/reference/add-ins/manifest/requirements) no manifesto do suplemento para especificar os conjuntos de requisitos m?nimos e/ou os m?todos de API exigidos pelo suplemento para ser ativado.</span><span class="sxs-lookup"><span data-stu-id="f9975-121">You can use the [Requirements element](https://dev.office.com/reference/add-ins/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="f9975-122">Se a plataforma ou o host do Office n?o der suporte aos conjuntos de requisitos ou aos m?todos de API que s?o especificados no elemento **Requirements** do manifesto, o suplemento n?o ser? executado nesse host ou plataforma e n?o ser? exibido na lista de suplementos que s?o mostrados em **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="f9975-122">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="f9975-123">O exemplo de c?digo a seguir mostra o elemento **Requirements** em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos host do Office que d?o suporte ao conjunto de requisitos ExcelApi, vers?o 1.3 ou superior.</span><span class="sxs-lookup"><span data-stu-id="f9975-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="f9975-124">Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel para Windows, Excel Online e Excel para iPad, ? recomend?vel verificar o suporte a requisitos no tempo de execu??o, em vez de definir o suporte ao conjunto de requisitos no manifesto.</span><span class="sxs-lookup"><span data-stu-id="f9975-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="f9975-125">Conjuntos de requisitos para a API comum Office.js</span><span class="sxs-lookup"><span data-stu-id="f9975-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="f9975-126">Para saber mais sobre conjuntos de requisitos de API comum, confira [Conjuntos de requisitos de API comum do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="f9975-126">For information about common API requirement sets, see [Office common API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="f9975-127">Carregando as propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-127">Loading the properties of an object</span></span>

<span data-ttu-id="f9975-128">Chamar o m?todo `load()` em um objeto JavaScript do Excel orienta a API a carregar o objeto na mem?ria do JavaScript quando o m?todo `sync()` ? executado.</span><span class="sxs-lookup"><span data-stu-id="f9975-128">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="f9975-129">O m?todo `load()` aceita uma cadeia de caracteres que cont?m nomes de propriedades delimitados por v?rgulas a serem carregados ou um objeto que especifica propriedades a serem carregadas, op??es de pagina??o, etc.</span><span class="sxs-lookup"><span data-stu-id="f9975-129">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="f9975-130">Se voc? chamar o m?todo `load()` em um objeto (ou uma cole??o) sem especificar qualquer par?metro, todas as propriedades escalares do objeto (ou todas as propriedades escalares de todos os objetos na cole??o) ser?o carregadas.</span><span class="sxs-lookup"><span data-stu-id="f9975-130">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="f9975-131">Para reduzir a quantidade de transfer?ncia de dados entre o aplicativo host e o suplemento do Excel, voc? deve evitar chamar o m?todo `load()` sem especificar explicitamente quais propriedades carregar.</span><span class="sxs-lookup"><span data-stu-id="f9975-131">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="f9975-132">Detalhes do m?todo</span><span class="sxs-lookup"><span data-stu-id="f9975-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="f9975-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="f9975-133">load(param: object)</span></span>

<span data-ttu-id="f9975-134">Preenche o objeto proxy criado na camada JavaScript com os valores da propriedade e do objeto especificados pelos par?metros.</span><span class="sxs-lookup"><span data-stu-id="f9975-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="f9975-135">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f9975-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="f9975-136">Par?metros</span><span class="sxs-lookup"><span data-stu-id="f9975-136">Parameters</span></span>

|<span data-ttu-id="f9975-137">**Par?metro**</span><span class="sxs-lookup"><span data-stu-id="f9975-137">**Parameter**</span></span>|<span data-ttu-id="f9975-138">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f9975-138">**Type**</span></span>|<span data-ttu-id="f9975-139">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="f9975-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="f9975-140">objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-140">object</span></span>|<span data-ttu-id="f9975-141">Opcional.</span><span class="sxs-lookup"><span data-stu-id="f9975-141">Optional.</span></span> <span data-ttu-id="f9975-142">Aceita nomes de par?metro e de rela??o como uma matriz ou cadeia de caracteres delimitada por v?rgulas.</span><span class="sxs-lookup"><span data-stu-id="f9975-142">Accepts parameter and relationship names as comma-delimited string or an array.</span></span> <span data-ttu-id="f9975-143">Um objeto tamb?m pode ser passado para definir as propriedades de navega??o e sele??o (conforme mostrado no exemplo abaixo).</span><span class="sxs-lookup"><span data-stu-id="f9975-143">An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="f9975-144">Retorna</span><span class="sxs-lookup"><span data-stu-id="f9975-144">Returns</span></span>

<span data-ttu-id="f9975-145">nulo</span><span class="sxs-lookup"><span data-stu-id="f9975-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="f9975-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9975-146">Example</span></span>

<span data-ttu-id="f9975-147">O exemplo de c?digo a seguir define as propriedades de um intervalo do Excel, copiando as propriedades de outro intervalo.</span><span class="sxs-lookup"><span data-stu-id="f9975-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="f9975-148">Observe que o objeto de origem deve ser carregado primeiro para que seus valores de propriedade possam ser acessados e gravados no intervalo de destino.</span><span class="sxs-lookup"><span data-stu-id="f9975-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="f9975-149">Este exemplo pressup?e que h? dados nos dois intervalos (**B2:E2** e **B7:E7**) e que os dois intervalos s?o inicialmente formatados de modo diferente.</span><span class="sxs-lookup"><span data-stu-id="f9975-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="f9975-150">Carregar propriedades de op??o</span><span class="sxs-lookup"><span data-stu-id="f9975-150">Load option properties</span></span>

<span data-ttu-id="f9975-151">Como uma alternativa para passar uma cadeia de caracteres delimitada por v?rgulas ou uma matriz ao chamar o m?todo `load()`, voc? pode passar um objeto que cont?m as propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9975-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="f9975-152">**Propriedade**</span><span class="sxs-lookup"><span data-stu-id="f9975-152">**Property**</span></span>|<span data-ttu-id="f9975-153">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f9975-153">**Type**</span></span>|<span data-ttu-id="f9975-154">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="f9975-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="f9975-155">objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-155">object</span></span>|<span data-ttu-id="f9975-p109">Inclui uma lista delimitada por v?rgula ou uma matriz de nomes de par?metro/rela??o. Opcional.</span><span class="sxs-lookup"><span data-stu-id="f9975-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="f9975-158">objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-158">object</span></span>|<span data-ttu-id="f9975-p110">Inclui uma lista delimitada por v?rgula ou uma matriz de nomes de rela??o. Opcional.</span><span class="sxs-lookup"><span data-stu-id="f9975-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="f9975-161">int</span><span class="sxs-lookup"><span data-stu-id="f9975-161">int</span></span>| <span data-ttu-id="f9975-p111">Especifica o n?mero m?ximo de itens da cole??o que podem ser inclu?dos no resultado. Opcional. Voc? s? pode usar essa op??o quando usar a op??o de nota??o de objeto.</span><span class="sxs-lookup"><span data-stu-id="f9975-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="f9975-165">int</span><span class="sxs-lookup"><span data-stu-id="f9975-165">int</span></span>|<span data-ttu-id="f9975-p112">Determina o n?mero de itens da cole??o que devem ser ignorados e n?o inclu?dos no resultado. Quando a propriedade `top` for especificada, o conjunto de resultados ser? iniciado depois de ignorar o n?mero de itens especificado. Opcional. Voc? s? pode usar esta op??o ao usar a op??o de nota??o de objeto.</span><span class="sxs-lookup"><span data-stu-id="f9975-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="f9975-170">O exemplo de c?digo a seguir carrega uma cole??o de planilhas selecionando a propriedade `name` e o `address` do intervalo usado para cada planilha na cole??o.</span><span class="sxs-lookup"><span data-stu-id="f9975-170">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="f9975-171">Ele tamb?m especifica que apenas as cinco planilhas principais na cole??o devem ser carregadas.</span><span class="sxs-lookup"><span data-stu-id="f9975-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="f9975-172">Voc? poderia processar o pr?ximo conjunto de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributo.</span><span class="sxs-lookup"><span data-stu-id="f9975-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="f9975-173">Propriedades escalares e de navega??o</span><span class="sxs-lookup"><span data-stu-id="f9975-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="f9975-174">Na documenta??o de refer?ncia da API JavaScript do Excel, voc? pode notar que os membros do objeto s?o agrupados em duas categorias: **propriedades** e **rela??es**.</span><span class="sxs-lookup"><span data-stu-id="f9975-174">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**.</span></span> <span data-ttu-id="f9975-175">Uma propriedade de um objeto ? um membro escalar como uma cadeia de caracteres, um n?mero inteiro ou um valor booliano, enquanto uma rela??o de um objeto (tamb?m conhecida como uma propriedade de navega??o) ? um membro que ? ou um objeto ou uma cole??o de objetos.</span><span class="sxs-lookup"><span data-stu-id="f9975-175">A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects.</span></span> <span data-ttu-id="f9975-176">Por exemplo, os membros `name` e `position` no objeto [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) s?o propriedades escalares, enquanto `protection` e `tables` s?o rela??es (propriedades de navega??o).</span><span class="sxs-lookup"><span data-stu-id="f9975-176">For example, `name` and `position` members on the [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="f9975-177">Propriedades escalares e propriedades de navega??o com `object.load()`</span><span class="sxs-lookup"><span data-stu-id="f9975-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="f9975-178">Chamar o m?todo `object.load()` sem par?metros especificados carregar? todas as propriedades escalares do objeto; as propriedades de navega??o do objeto n?o ser?o carregadas.</span><span class="sxs-lookup"><span data-stu-id="f9975-178">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="f9975-179">Al?m disso, as propriedades de navega??o n?o podem ser carregadas diretamente.</span><span class="sxs-lookup"><span data-stu-id="f9975-179">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="f9975-180">Em vez disso, voc? deve usar o m?todo `load()` para fazer refer?ncia ?s propriedades escalares individuais na propriedade de navega??o desejada.</span><span class="sxs-lookup"><span data-stu-id="f9975-180">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="f9975-181">Por exemplo, para carregar o nome da fonte de um intervalo, voc? deve especificar as propriedades de navega??o **format** e **font** como o caminho para a propriedade **name**:</span><span class="sxs-lookup"><span data-stu-id="f9975-181">For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="f9975-182">com a API JavaScript do Excel, ? poss?vel definir propriedades escalares de uma propriedade de navega??o percorrendo o caminho.</span><span class="sxs-lookup"><span data-stu-id="f9975-182">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="f9975-183">Por exemplo, ? poss?vel definir o tamanho da fonte de um intervalo usando `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="f9975-183">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="f9975-184">N?o ? necess?rio carregar a propriedade antes de configur?-la.</span><span class="sxs-lookup"><span data-stu-id="f9975-184">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="f9975-185">Definindo propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-185">Setting properties of an object</span></span>

<span data-ttu-id="f9975-186">A defini??o de propriedades em um objeto com propriedades de navega??o aninhadas pode ser uma tarefa complicada.</span><span class="sxs-lookup"><span data-stu-id="f9975-186">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="f9975-187">Como uma alternativa para definir propriedades individuais usando caminhos de navega??o, conforme descrito acima, voc? pode usar o m?todo `object.set()` que est? dispon?vel em todos os objetos na API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f9975-187">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="f9975-188">Com esse m?todo, ? poss?vel definir v?rias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que s?o estruturadas, como as propriedades do objeto no qual o m?todo ? chamado.</span><span class="sxs-lookup"><span data-stu-id="f9975-188">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="f9975-189">O m?todo `set()` ? implementado apenas para objetos nas APIs JavaScript do Office espec?ficas de host, como a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f9975-189">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="f9975-190">As APIs comuns (compartilhadas) n?o d?o suporte a esse m?todo.</span><span class="sxs-lookup"><span data-stu-id="f9975-190">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="f9975-191">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="f9975-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="f9975-p119">As propriedades do objeto em que o m?todo ? chamado s?o definidas com os mesmos valores das propriedades correspondentes do objeto transmitido. Se o par?metro `properties` for um objeto JavaScript, as propriedades do objeto transmitido que correspondem ? propriedade de somente leitura no objeto em que o m?todo ? chamado ser?o ignoradas ou causar?o uma exce??o, dependendo do par?metro `options`.</span><span class="sxs-lookup"><span data-stu-id="f9975-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="f9975-194">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f9975-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="f9975-195">Par?metros</span><span class="sxs-lookup"><span data-stu-id="f9975-195">Parameters</span></span>

|<span data-ttu-id="f9975-196">**Par?metro**</span><span class="sxs-lookup"><span data-stu-id="f9975-196">**Parameter**</span></span>|<span data-ttu-id="f9975-197">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f9975-197">**Type**</span></span>|<span data-ttu-id="f9975-198">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="f9975-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="f9975-199">objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-199">object</span></span>|<span data-ttu-id="f9975-200">Um objeto do mesmo tipo de objeto do Office.js no qual o m?todo ? chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o m?todo ? chamado.</span><span class="sxs-lookup"><span data-stu-id="f9975-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="f9975-201">objeto</span><span class="sxs-lookup"><span data-stu-id="f9975-201">object</span></span>|<span data-ttu-id="f9975-p120">Opcional. S? pode ser transmitido quando o primeiro par?metro ? um objeto JavaScript. O objeto pode conter a seguinte propriedade: `throwOnReadOnly?: boolean` (O padr?o ? `true`: indicar um erro se o objeto JavaScript transmitido incluir propriedades de somente leitura.)</span><span class="sxs-lookup"><span data-stu-id="f9975-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="f9975-205">Retorna</span><span class="sxs-lookup"><span data-stu-id="f9975-205">Returns</span></span>

<span data-ttu-id="f9975-206">nulo</span><span class="sxs-lookup"><span data-stu-id="f9975-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="f9975-207">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9975-207">Example</span></span>

<span data-ttu-id="f9975-p121">O exemplo de c?digo a seguir define v?rias propriedades do formato de um intervalo chamando o m?todo `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto **Range**. Este exemplo sup?e que h? dados no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="f9975-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="f9975-210">M?todos \*OrNullObject</span><span class="sxs-lookup"><span data-stu-id="f9975-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="f9975-211">Muitos m?todos da API JavaScript do Excel retornar?o uma exce??o quando a condi??o da API n?o for atendida.</span><span class="sxs-lookup"><span data-stu-id="f9975-211">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="f9975-212">Por exemplo, se voc? tentar obter uma planilha especificando um nome de planilha que n?o existe na pasta de trabalho, o m?todo `getItem()` retornar? uma exce??o `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="f9975-212">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="f9975-213">Em vez de implementar a l?gica complexa de tratamento de exce??o para cen?rios como este, voc? pode usar a variante do m?todo `*OrNullObject` que est? dispon?vel para v?rios m?todos na API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f9975-213">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="f9975-214">Um m?todo `*OrNullObject` retornar? um objeto nulo (n?o o `null` do JavaScript), em vez de emitir uma exce??o se o item especificado n?o existir.</span><span class="sxs-lookup"><span data-stu-id="f9975-214">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="f9975-215">Por exemplo, voc? pode chamar o m?todo `getItemOrNullObject()` em uma cole??o, como **Worksheets**, para tentar recuperar um item da cole??o.</span><span class="sxs-lookup"><span data-stu-id="f9975-215">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="f9975-216">O m?todo `getItemOrNullObject()` retornar? o item especificado se ele existir; caso contr?rio, ele retornar? um objeto nulo.</span><span class="sxs-lookup"><span data-stu-id="f9975-216">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="f9975-217">O objeto nulo que ? retornado cont?m a propriedade booliana `isNullObject`, que voc? pode avaliar para determinar se o objeto existe.</span><span class="sxs-lookup"><span data-stu-id="f9975-217">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="f9975-218">O exemplo de c?digo a seguir tenta recuperar uma planilha chamada "Data" usando o m?todo `getItemOrNullObject()`.</span><span class="sxs-lookup"><span data-stu-id="f9975-218">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="f9975-219">Se o m?todo retornar um objeto nulo, uma nova folha precisar? ser criada para que as a??es possam ser tomadas na folha.</span><span class="sxs-lookup"><span data-stu-id="f9975-219">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="f9975-220">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="f9975-220">See also</span></span>
 
* [<span data-ttu-id="f9975-221">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f9975-221">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="f9975-222">Exemplos de c?digo de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="f9975-222">Excel add-ins code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="f9975-223">Otimiza??o de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f9975-223">Excel JavaScript API performance optimization</span></span>](https://dev.office.com/reference/add-ins/excel/performance.md)
* [<span data-ttu-id="f9975-224">Refer?ncia da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f9975-224">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
