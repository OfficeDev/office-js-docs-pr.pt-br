---
title: Conceitos avançados de programação com a API JavaScript do Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 09f2d95e4cf7631b519f00cddee265dbf697e07e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505885"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="dc1c3-102">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dc1c3-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="dc1c3-103">Este artigo foi criado com base nas informações em [Conceitos fundamentais da API JavaScript do Excel](excel-add-ins-core-concepts.md) para descrever alguns dos conceitos mais avançados que são essenciais para criar suplementos complexos para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-103">This article builds upon the information in [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="dc1c3-104">APIs Office.js para Excel</span><span class="sxs-lookup"><span data-stu-id="dc1c3-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="dc1c3-105">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:</span><span class="sxs-lookup"><span data-stu-id="dc1c3-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="dc1c3-106">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="dc1c3-107">**APIs comuns**: introduzidas com o Office 2013, as APIs comuns (também conhecidas como [API Compartilhada](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)) podem ser usadas para acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos host, como Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-107">**Common APIs**: Introduced with Office 2013, the common APIs (also referred to as the [Shared API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="dc1c3-p101">Embora você provavelmente usará a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016 ou posterior , você também usará objetos na API Compartilhada. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p101">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Shared API. For example:</span></span>

- <span data-ttu-id="dc1c3-p102">[Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js): o objeto **Context** representa o ambiente em tempo de execução do suplemento e fornece acesso aos principais objetos da API. Ele consiste em detalhes de configuração da pasta de trabalho como `contentLanguage` e `officeTheme` e também fornece informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se o conjunto de requisitos especificado tem suporte no aplicativo Excel em que o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p102">[Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="dc1c3-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js): o objeto **Document** fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo Excel onde o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="dc1c3-114">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="dc1c3-114">Requirement sets</span></span>

<span data-ttu-id="dc1c3-p103">Conjuntos de requisitos são grupos nomeados de membros da API. Um suplemento do Office pode executar uma verificação em tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um host do Office suporta as APIs que o suplemento precisa. Para identificar os conjuntos de requisitos específico que estão disponíveis em cada plataforma compatível, confira [Conjuntos de requisitos da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p103">Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="dc1c3-118">Verificar compatibilidade com o conjunto de requisitos em tempo de execução</span><span class="sxs-lookup"><span data-stu-id="dc1c3-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="dc1c3-119">O exemplo de código a seguir mostra como determinar se o aplicativo host, onde o suplemento está sendo executado é compatível com o conjunto de requisitos especificado pela API .</span><span class="sxs-lookup"><span data-stu-id="dc1c3-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="dc1c3-120">Definir a compatibilidade com o conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="dc1c3-p104">Você pode usar o [elemento Requirements](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements?view=office-js) no manifesto do suplemento para especificar o conjunto mínimo de requisitos e/ou os métodos de API que o seu suplemento necessita para ser ativado. Se o host do Office ou outra plataforma não for compatível com o  conjunto de requisitos ou métodos da API especificados no elemento **Requirements** do manifesto, o suplemento não será executado nessa plataforma ou host e não será exibido na lista **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p104">You can use the [Requirements element](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements?view=office-js) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="dc1c3-123">O exemplo de código a seguir mostra o elemento **Requirements** em um manifesto de suplemento que especifica que o suplemento deve ser carregado em todos os aplicativos host do Office compatíveis com o  conjunto de requisitos ExcelApi, versão 1.3 ou superior.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="dc1c3-124">Para disponibilizar seu suplemento em todas as plataformas de um host do Office, como Excel para Windows, Excel Online e Excel para iPad, é recomendável verificar a compatibilidade com os requisitos em tempo de execução, em vez de defini-lo no manifesto.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="dc1c3-125">Conjuntos de requisitos para a API comum Office.js</span><span class="sxs-lookup"><span data-stu-id="dc1c3-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="dc1c3-126">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="dc1c3-126">For information about common API requirement sets, see [Office common API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="dc1c3-127">Carregar as propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-127">Loading the properties of an object</span></span>

<span data-ttu-id="dc1c3-p105">Chamar o método `load()` em um objeto JavaScript do Excel instrui a API para carregar o objeto na memória do JavaScript quando o método `sync()` for executado. O método `load()` aceita uma sequência de caracteres que contém os nomes das propriedades que devem ser carregadas delimitadas por vírgulas ou um objeto que especifica as propriedades a carregar, opções de paginação, etc.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p105">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs. The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="dc1c3-p106">Se você chamar o método `load()` em um objeto (ou coleção) sem especificar nenhum parâmetro, todas as propriedades escalares do objeto (ou todas as propriedades escalares de todos os objetos da coleção) serão carregadas. Para reduzir o volume de dados transferidos entre o Excel e o suplemento, você deve evitar chamadas ao método `load()`  sem especificar explicitamente quais propriedades para carregar.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p106">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="dc1c3-132">Detalhes do método</span><span class="sxs-lookup"><span data-stu-id="dc1c3-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="dc1c3-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="dc1c3-133">load(param: object)</span></span>

<span data-ttu-id="dc1c3-134">Preenche o objeto de proxy criado na camada JavaScript com os valores de propriedades e objetos especificados pelos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="dc1c3-135">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="dc1c3-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="dc1c3-136">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="dc1c3-136">Parameters</span></span>

|<span data-ttu-id="dc1c3-137">**Parâmetro**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-137">**Parameter**</span></span>|<span data-ttu-id="dc1c3-138">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-138">**Type**</span></span>|<span data-ttu-id="dc1c3-139">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="dc1c3-140">objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-140">object</span></span>|<span data-ttu-id="dc1c3-p107">Opcional. Aceita nomes de parâmetro e relacionamento como uma sequência de caracteres delimitadas por vírgulas ou uma matriz. Um objeto também pode ser passado para definir as propriedades de seleção e navegação (conforme mostrado no exemplo a seguir).</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p107">Optional. Accepts parameter and relationship names as comma-delimited string or an array. An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="dc1c3-144">Retorna</span><span class="sxs-lookup"><span data-stu-id="dc1c3-144">Returns</span></span>

<span data-ttu-id="dc1c3-145">nulo</span><span class="sxs-lookup"><span data-stu-id="dc1c3-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="dc1c3-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dc1c3-146">Example</span></span>

<span data-ttu-id="dc1c3-p108">O exemplo de código a seguir define as propriedades de um intervalo do Excel, copiando as propriedades de outro intervalo. Observe que o objeto de origem deve ser carregado primeiro para que os valores de suas propriedades possam ser acessados e gravados no intervalo de destino. Este exemplo pressupõe que haja dados os dois intervalos (**B2:E2** e **B7:E7**) e os dois intervalos tenham inicialmente uma formatação diferente.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p108">The following code sample sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first, before its property values can be accessed and written to the target range. This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="dc1c3-150">Opções de propriedades de carga</span><span class="sxs-lookup"><span data-stu-id="dc1c3-150">Load option properties</span></span>

<span data-ttu-id="dc1c3-151">Como alternativa para passar uma sequência de caracteres delimitada por vírgulas ou uma matriz ao chamar o método `load()`, você pode passar um objeto que contém as propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="dc1c3-152">**Propriedade**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-152">**Property**</span></span>|<span data-ttu-id="dc1c3-153">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-153">**Type**</span></span>|<span data-ttu-id="dc1c3-154">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="dc1c3-155">objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-155">object</span></span>|<span data-ttu-id="dc1c3-p109">Inclui uma lista delimitada por vírgulas ou uma matriz de nomes de parâmetro/relacionamentos. Opcional.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="dc1c3-158">objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-158">object</span></span>|<span data-ttu-id="dc1c3-p110">Inclui uma lista delimitada por vírgulas ou uma matriz de nomes de relacionamentos. Opcional.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="dc1c3-161">int</span><span class="sxs-lookup"><span data-stu-id="dc1c3-161">int</span></span>| <span data-ttu-id="dc1c3-p111">Especifica o número máximo de itens da coleção que podem ser incluídos no resultado. Opcional. Você só pode usar essa opção quando usar a opção de notação de objeto.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="dc1c3-165">int</span><span class="sxs-lookup"><span data-stu-id="dc1c3-165">int</span></span>|<span data-ttu-id="dc1c3-p112">Determina o número de itens da coleção que devem ser ignorados e não incluídos no resultado. Quando a propriedade `top` for especificada, o conjunto de resultados será iniciado depois de ignorar o número de itens especificado. Opcional. Você só pode usar esta opção ao usar a opção de notação de objeto.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="dc1c3-p113">O exemplo de código a seguir carrega uma coleção planilhas selecionando a propriedade `name` e o intervalo `address` do intervalo utilizado cada planilha na coleção. Ela também especifica que apenas as cinco primeiras planilhas na coleção devem ser carregadas. Você poderia processar o conjunto seguinte de cinco planilhas especificando `top: 10` e `skip: 5` como valores de atributos.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p113">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection. It also specifies that only the top five worksheets in the collection should be loaded. You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="dc1c3-173">Propriedades escalares e de navegação</span><span class="sxs-lookup"><span data-stu-id="dc1c3-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="dc1c3-p114">Na documentação de referência de API JavaScript do Excel, você pode observar que os objetos membros são agrupados em duas categorias: **propriedades** e **relacionamentos**. Uma propriedade de um objeto é um membro escalar como uma sequência de caracteres, um número inteiro ou um valor booleano, enquanto o relacionamento de um objeto (também conhecido como propriedade de navegação) é um membro que é um objeto ou a coleção de objetos. Por exemplo, os membros `name` e `position` do objeto [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) são propriedades escalares, enquanto `protection` e `tables` são relacionamentos (propriedades de navegação).</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p114">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**. A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects. For example, `name` and `position` members on the [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="dc1c3-177">Propriedades escalares e propriedades de navegação com `object.load()`</span><span class="sxs-lookup"><span data-stu-id="dc1c3-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="dc1c3-p115">Chamar o método `object.load()` sem especificar parâmetros carregará todas as propriedades escalares do objeto; as propriedades de navegação do objeto não serão carregadas. Além disso, não é possível carregar as propriedades de navegação diretamente. Em vez disso, você deve usar o método  `load()` para referenciar propriedades escalares individualmente dentro das propriedade de navegação desejadas. Por exemplo, para carregar o nome da fonte de um intervalo, você deve especificar as propriedades de navegação **format** e **font** como caminho para a propriedade **name** :</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p115">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded. Additionally, navigation properties cannot be loaded directly. Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property. For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="dc1c3-p116">Com a API JavaScript do Excel, você pode definir propriedades escalares de uma propriedade de navegação percorrendo o caminho. Por exemplo, você pode definir o tamanho da fonte para um intervalo usando `someRange.format.font.size = 10;`. Você não precisará carregar a propriedade antes de defini-la.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p116">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path. For example, you could set the font size for a range by using `someRange.format.font.size = 10;`. You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="dc1c3-185">Definir propriedades de um objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-185">Setting properties of an object</span></span>

<span data-ttu-id="dc1c3-p117">Definir propriedades em um objeto com propriedades de navegação aninhadas pode ser complicado. Como alternativa à definição de propriedades individuais usando caminhos de navegação, conforme descrito acima, você pode usar o método `object.set()` que está disponível em todos os objetos na API JavaScript do Excel. Com este método, você pode definir várias propriedades de um objeto ao mesmo tempo, passando um outro objeto do mesmo tipo da Office.js ou um objeto JavaScript com propriedades que são estruturadas como as propriedades do objeto no qual o método é chamado.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p117">Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="dc1c3-p118">O método `set()` é implementado apenas para objetos dentro das APIs JavaScript do Office específicas do host, como a API JavaScript do Excel. As APIs (compartilhadas) comuns não suportam esse método.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p118">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API. The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="dc1c3-191">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="dc1c3-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="dc1c3-p119">As propriedades do objeto em que o método é chamado são definidas com os mesmos valores das propriedades correspondentes do objeto passado. Se o parâmetro `properties` for um objeto JavaScript, as propriedades do objeto passado que correspondem à propriedades somente para leitura no objeto em que o método é chamado serão ignoradas ou causarão uma exceção, dependendo do parâmetro `options`.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="dc1c3-194">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="dc1c3-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="dc1c3-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="dc1c3-195">Parameters</span></span>

|<span data-ttu-id="dc1c3-196">**Parâmetro**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-196">**Parameter**</span></span>|<span data-ttu-id="dc1c3-197">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-197">**Type**</span></span>|<span data-ttu-id="dc1c3-198">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="dc1c3-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="dc1c3-199">objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-199">object</span></span>|<span data-ttu-id="dc1c3-200">Um objeto do mesmo tipo de objeto da Office.js no qual o método é chamado ou um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura do objeto no qual o método é chamado.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="dc1c3-201">objeto</span><span class="sxs-lookup"><span data-stu-id="dc1c3-201">object</span></span>|<span data-ttu-id="dc1c3-p120">Opcional. Só pode ser passado quando o primeiro parâmetro é um objeto JavaScript. O objeto pode conter a seguinte propriedade: `throwOnReadOnly?: boolean` (O padrão é `true`: indicar um erro se o objeto JavaScript passado incluir propriedades com acesso de somente leitura.)</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="dc1c3-205">Retorna</span><span class="sxs-lookup"><span data-stu-id="dc1c3-205">Returns</span></span>

<span data-ttu-id="dc1c3-206">nulo</span><span class="sxs-lookup"><span data-stu-id="dc1c3-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="dc1c3-207">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dc1c3-207">Example</span></span>

<span data-ttu-id="dc1c3-p121">O exemplo de código a seguir define várias propriedades de formatação de um intervalo chamando o método `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto **Range**. Este exemplo supõe que há dados no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="dc1c3-210">Métodos \*OrNullObject</span><span class="sxs-lookup"><span data-stu-id="dc1c3-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="dc1c3-p122">Muitos métodos da API JavaScript do Excel retornam uma exceção quando a condição da API não é atendida. Por exemplo, se você tentar obter uma planilha, especificando um nome de planilha que não existe na pasta de trabalho, o método  `getItem()` retornará uma exceção de `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p122">Many Excel JavaScript API methods will return an exception when the condition of the API is not met. For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="dc1c3-p123">Em vez de implementar uma lógica complexa de tratamento de exceções em cenários como esse, você pode usar a variante do método `*OrNullObject`  que está disponível para vários métodos na API JavaScript do Excel. Um método `*OrNullObject` retornará um objeto nulo (não o JavaScript `null`) em vez de gerar uma exceção se o item especificado não existir. Por exemplo, você pode chamar o método `getItemOrNullObject()` em um conjunto como **planilhas** para tentar recuperar um item da coleção. O método  `getItemOrNullObject()` retorna o item especificado se ele existir. Caso contrário, retorna um objeto nulo. O objeto nulo retornado contém a propriedade booleana `isNullObject` que pode ser avaliada para determinar se o objeto existe.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p123">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API. An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="dc1c3-p124">O exemplo de código a seguir tenta recuperar uma planilha chamada "Data" usando o método `getItemOrNullObject()` . Se o método retornar um objeto nulo, uma nova planilha deve ser criada para que as ações sejam executadas.</span><span class="sxs-lookup"><span data-stu-id="dc1c3-p124">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="dc1c3-220">Confira também</span><span class="sxs-lookup"><span data-stu-id="dc1c3-220">See also</span></span>
 
* [<span data-ttu-id="dc1c3-221">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dc1c3-221">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="dc1c3-222">Exemplos de códigos de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="dc1c3-222">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="dc1c3-223">Otimização de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dc1c3-223">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="dc1c3-224">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dc1c3-224">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
