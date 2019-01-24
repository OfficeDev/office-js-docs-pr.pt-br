---
title: Especificar hosts do Office e requisitos de API
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 3ea0116d8c8e9dcf685db349d78ed43af91b1620
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386817"
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="d3f06-102">Especificar hosts do Office e requisitos de API</span><span class="sxs-lookup"><span data-stu-id="d3f06-102">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="d3f06-p101">Seu Suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado. Por exemplo, o suplemento pode:</span><span class="sxs-lookup"><span data-stu-id="d3f06-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="d3f06-105">Executar em um único aplicativo do Office (Word ou Excel) ou diversos aplicativos.</span><span class="sxs-lookup"><span data-stu-id="d3f06-105">Run in a single Office application (Word or Excel), or several applications.</span></span>
    
- <span data-ttu-id="d3f06-p102">Usar as APIs de JavaScript que estão disponíveis apenas em algumas versões do Office. Por exemplo, você pode usar as APIs JavaScript do Excel em um suplemento executado no Excel 2016.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span> 
    
- <span data-ttu-id="d3f06-108">Executar apenas nas versões do Office que oferecem suporte a membros da API que seu suplemento usa.</span><span class="sxs-lookup"><span data-stu-id="d3f06-108">Run only in versions of Office that support API members that your add-in uses.</span></span>
    
<span data-ttu-id="d3f06-109">Este artigo ajuda você a entender quais opções você deve escolher para garantir que seu suplemento funcione conforme o esperado e atinja o público mais amplo possível.</span><span class="sxs-lookup"><span data-stu-id="d3f06-109">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="d3f06-110">Confira uma visão avançada da compatibilidade atual dos suplementos do Office no momento na página [Disponibilidade de hosts e plataformas de suplementos do Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="d3f06-110">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span> 

<span data-ttu-id="d3f06-111">A tabela a seguir lista os principais conceitos discutidos neste artigo.</span><span class="sxs-lookup"><span data-stu-id="d3f06-111">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="d3f06-112">**Conceito**</span><span class="sxs-lookup"><span data-stu-id="d3f06-112">**Concept**</span></span>|<span data-ttu-id="d3f06-113">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="d3f06-113">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="d3f06-114">Aplicativo do Office, aplicativo host do Office, host do Office ou host</span><span class="sxs-lookup"><span data-stu-id="d3f06-114">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="d3f06-115">O aplicativo do Office usado para executar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d3f06-115">The Office application used to run your add-in.</span></span> <span data-ttu-id="d3f06-116">Por exemplo, Word, Word Online, Excel etc.</span><span class="sxs-lookup"><span data-stu-id="d3f06-116">For example, Word, Word Online, Excel, and so on.</span></span>|
|<span data-ttu-id="d3f06-117">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d3f06-117">Platform</span></span>|<span data-ttu-id="d3f06-118">Onde o host do Office é executado, por exemplo, no Office Online ou no Office para iPad.</span><span class="sxs-lookup"><span data-stu-id="d3f06-118">Where the Office host runs, such as Office Online or Office for iPad.</span></span>|
|<span data-ttu-id="d3f06-119">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="d3f06-119">Requirement set</span></span>|<span data-ttu-id="d3f06-p104">Um grupo nomeado de membros relacionados da API. Os suplementos usam conjuntos de requisitos para determinar se o host do Office oferece suporte a membros da API usados por seu suplemento. É mais fácil testar se há suporte para um conjunto de requisitos do que o suporte para membros individuais da API. O suporte a um conjunto de requisitos varia de acordo com o host do Office e a versão do host do Office. </span><span class="sxs-lookup"><span data-stu-id="d3f06-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="d3f06-124">Conjuntos de requisitos são especificados no arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="d3f06-124">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="d3f06-125">Ao especificar conjuntos de requisitos no manifesto, você estabelece o nível mínimo de suporte à API que o host do Office deve fornecer a fim de executar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d3f06-125">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in.</span></span> <span data-ttu-id="d3f06-126">Hosts do Office que não há suporte para conjuntos de requisitos especificados no manifesto não é possível executar o suplemento e o suplemento não será exibido em <span class="ui">Meus suplementos</span>. Isso restringirá onde o suplemento está disponível. Usar a verificação do tempo de execução do código.</span><span class="sxs-lookup"><span data-stu-id="d3f06-126">Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.In code using runtime checks.</span></span> <span data-ttu-id="d3f06-127">Para obter uma lista completa de conjuntos de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d3f06-127">For the complete list of requirement sets, see [Office Add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>|
|<span data-ttu-id="d3f06-128">Verificação no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="d3f06-128">Runtime check</span></span>|<span data-ttu-id="d3f06-p106">Um teste é executado no tempo de execução para determinar se o host do Office que está executando seu suplemento oferece suporte aos conjuntos de requisitos ou métodos usados por seu suplemento. Para executar uma verificação no tempo de execução, use uma instrução **if** com o método **isSetSupported**, os conjuntos de requisito ou os nomes de método que não fazem parte de um conjunto de requisitos. Use as verificações no tempo de execução para garantir que seu suplemento alcance o maior número de clientes. Ao contrário dos conjuntos de requisitos, as verificações no tempo de execução não especificam o nível mínimo de suporte à API exigido do host do Office para que seu suplemento possa ser executado. Em vez disso, use a instrução **if** para determinar se há suporte para um membro da API. Se houver, você poderá proporcionar mais funcionalidade em seu suplemento. Seu suplemento sempre aparecerá em **Meus Suplementos** ao usar verificações no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p106">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in. To perform a runtime check, you use an  **if** statement with the **isSetSupported** method, the requirement sets, or the method names that aren't part of a requirement set.Use runtime checks to ensure that your add-in reaches the broadest number of customers. Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run. Instead, you use the  **if** statement to determine whether an API member is supported. If it is, you can provide additional functionality in your add-in. Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="d3f06-135">Antes de começar</span><span class="sxs-lookup"><span data-stu-id="d3f06-135">Before you begin</span></span>

<span data-ttu-id="d3f06-p107">O suplemento deve usar a versão mais recente do esquema de manifesto de suplemento. Se você usar as verificações no tempo de execução em seu suplemento, use a biblioteca mais recente da API JavaScript para Office (office.js).</span><span class="sxs-lookup"><span data-stu-id="d3f06-p107">Your add-in must use the most current version of the add-in manifest schema. If you use runtime checks in your add-in, ensure that you use the latest JavaScript API for Office (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="d3f06-138">Especificar o esquema de manifesto de suplemento mais recente</span><span class="sxs-lookup"><span data-stu-id="d3f06-138">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="d3f06-p108">Seu manifesto de suplemento deve usar a versão 1.1 do esquema de manifesto de suplemento. Defina o elemento **OfficeApp** no manifesto do seu suplemento da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p108">Your add-in's manifest must use version 1.1 of the add-in manifest schema. Set the  **OfficeApp** element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a><span data-ttu-id="d3f06-141">Especificar a biblioteca de API JavaScript para Office mais recente</span><span class="sxs-lookup"><span data-stu-id="d3f06-141">Specify the latest JavaScript API for Office library</span></span>

<span data-ttu-id="d3f06-p109">Se você usar as verificações no tempo de execução, faça referência à versão mais recente da biblioteca de API JavaScript para Office na CDN (rede de distribuição de conteúdo). Para tanto, adicione a seguinte marca `script` ao código HTML. O uso de `/1/` na URL da CDN garante a referência à versão mais recente do Office.js.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p109">If you use runtime checks, reference the most current version of the JavaScript API for Office library from the content delivery network (CDN). To do this, add the following  `script` tag to your HTML. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="d3f06-145">Opções para especificar os hosts do Office ou requisitos de API</span><span class="sxs-lookup"><span data-stu-id="d3f06-145">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="d3f06-p110">Ao especificar os hosts do Office ou os requisitos de API, há vários fatores a considerar. O diagrama a seguir mostra como decidir sobre qual técnica usar em seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![Escolha a melhor opção para o seu suplemento ao especificar os hosts do Office ou os requisitos de API](../images/options-for-office-hosts.png)

- <span data-ttu-id="d3f06-149">Se o seu suplemento for executado em um host do Office, defina o elemento **Hosts** no manifesto.</span><span class="sxs-lookup"><span data-stu-id="d3f06-149">If your add-in runs in one Office host, set the **Hosts** element in the manifest.</span></span> <span data-ttu-id="d3f06-150">Para saber mais, confira [Definir o elemento Hosts](#set-the-hosts-element).</span><span class="sxs-lookup"><span data-stu-id="d3f06-150">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>
    
- <span data-ttu-id="d3f06-p112">Para definir o conjunto de requisitos mínimos ou os membros da API que devem receber suporte de um host do Office para que seu suplemento seja executado, defina o elemento **Requirements** no manifesto. Para saber mais, confira [Definir o elemento Requirements no manifesto](#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="d3f06-p112">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the  **Requirements** element in the manifest. For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>
    
- <span data-ttu-id="d3f06-p113">Se você quiser fornecer outras funcionalidades caso conjuntos de requisitos ou membros da API específicos estejam disponíveis no host do Office, execute uma verificação no tempo de execução no código JavaScript do seu suplemento. Por exemplo, se o seu suplemento for executado no Excel 2016, use os membros da nova API JavaScript para Excel a fim de fornecer outras funcionalidades. Para saber mais, confira [Usar verificações de tempo de execução em seu código JavaScript](#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="d3f06-p113">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code. For example, if your add-in runs in Excel 2016, use API members from the new JavaScript API for Excel to provide additional functionality. For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>
    
## <a name="set-the-hosts-element"></a><span data-ttu-id="d3f06-156">Definir o elemento Hosts</span><span class="sxs-lookup"><span data-stu-id="d3f06-156">Set the Hosts element</span></span>

<span data-ttu-id="d3f06-p114">Para fazer seu suplemento ser executado em um aplicativo host do Office, use os elementos  **Hosts** e **Host** no manifesto. Se você não especificar o elemento **Hosts**, o suplemento será executado em todos os hosts.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p114">To make your add-in run in one Office host application, use the  **Hosts** and **Host** elements in the manifest. If you don't specify the **Hosts** element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="d3f06-159">Por exemplo, a declaração de **Hosts** e **Host** a seguir especifica que o suplemento funcionará com qualquer versão do Excel, o que inclui o Excel para Windows, o Excel Online e o Excel para iPad.</span><span class="sxs-lookup"><span data-stu-id="d3f06-159">For example, the following  **Hosts** and **Host** declaration specifies that the add-in will work with any release of Excel, which includes Excel for Windows, Excel Online, and Excel for iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="d3f06-p115">O elemento **Hosts** pode conter um ou mais elementos **Host**. O elemento **Host** especifica o host do Office exigido por seu suplemento. O atributo **Name** é obrigatório e pode ser definido com um dos valores a seguir.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p115">The  **Hosts** element can contain one or more **Host** elements. The **Host** element specifies the Office host your add-in requires. The **Name** attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="d3f06-163">Nome</span><span class="sxs-lookup"><span data-stu-id="d3f06-163">Name</span></span>          | <span data-ttu-id="d3f06-164">Aplicativos host do Office</span><span class="sxs-lookup"><span data-stu-id="d3f06-164">Office host applications</span></span>                      |
|:--------------|:----------------------------------------------|
| <span data-ttu-id="d3f06-165">Banco de dados</span><span class="sxs-lookup"><span data-stu-id="d3f06-165">Database</span></span>      | <span data-ttu-id="d3f06-166">Aplicativos Web do Access</span><span class="sxs-lookup"><span data-stu-id="d3f06-166">Access web apps</span></span>                               |
| <span data-ttu-id="d3f06-167">Documento</span><span class="sxs-lookup"><span data-stu-id="d3f06-167">Document</span></span>      | <span data-ttu-id="d3f06-168">Word para Windows, Mac, iPad e Online</span><span class="sxs-lookup"><span data-stu-id="d3f06-168">Word for Windows, Mac, iPad and Online</span></span>        |
| <span data-ttu-id="d3f06-169">Caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d3f06-169">Mailbox</span></span>       | <span data-ttu-id="d3f06-170">Outlook para Windows, Mac, Web e Outlook.com</span><span class="sxs-lookup"><span data-stu-id="d3f06-170">Outlook for Windows, Mac, Web and Outlook.com</span></span> | 
| <span data-ttu-id="d3f06-171">Apresentação</span><span class="sxs-lookup"><span data-stu-id="d3f06-171">Presentation</span></span>  | <span data-ttu-id="d3f06-172">PowerPoint para Windows, Mac, iPad e Online</span><span class="sxs-lookup"><span data-stu-id="d3f06-172">PowerPoint for Windows, Mac, iPad and Online</span></span>  |
| <span data-ttu-id="d3f06-173">Project</span><span class="sxs-lookup"><span data-stu-id="d3f06-173">Project</span></span>       | <span data-ttu-id="d3f06-174">Project</span><span class="sxs-lookup"><span data-stu-id="d3f06-174">Project</span></span>                                       |
| <span data-ttu-id="d3f06-175">Pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="d3f06-175">Workbook</span></span>      | <span data-ttu-id="d3f06-176">Excel para Windows, Mac, iPad e Online</span><span class="sxs-lookup"><span data-stu-id="d3f06-176">Excel Windows, Mac, iPad and Online</span></span>           |

> [!NOTE]
> <span data-ttu-id="d3f06-p116">O atributo `Name` especifica o aplicativo host do Office que pode executar seu suplemento. Há suporte para hosts do Office em várias plataformas, que são executados em computadores, navegadores da Web, tablets e dispositivos móveis. Você não pode especificar qual plataforma pode ser usada para executar seu suplemento. Por exemplo, se você especificar `Mailbox`, o Outlook e o Outlook Web App podem ser usados para executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p116">The  `Name` attribute specifies the Office host application that can run your add-in. Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You can't specify which platform can be used to run your add-in. For example, if you specify `Mailbox`, both Outlook and Outlook Web App can be used to run your add-in.</span></span> 


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="d3f06-181">Definir o elemento Requirements no manifesto</span><span class="sxs-lookup"><span data-stu-id="d3f06-181">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="d3f06-p117">O elemento **Requirements** especifica os conjuntos de requisitos mínimos ou os membros da API que devem receber suporte de um host do Office para que seu suplemento seja executado. O elemento **Requirements** pode especificar conjuntos de requisitos e métodos individuais usados em seu suplemento. Na versão 1.1 do esquema de manifesto de suplemento, o elemento **Requirements** é opcional para todos os suplementos, exceto para os suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p117">The  **Requirements** element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in. The **Requirements** element can specify both requirement sets and individual methods used in your add-in. In version 1.1 of the add-in manifest schema, the **Requirements** element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="d3f06-p118">Use o elemento **Requirements** apenas para especificar conjuntos de requisitos ou membros de API cruciais ao seu suplemento. Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no elemento **Requirements**, o suplemento não será executado no host ou na plataforma e não será exibido em **Meus Suplementos**. Em vez disso, recomendamos que você disponibilize seu suplemento em todas as plataformas de um host do Office, como o Excel para Windows, o Excel Online e o Excel para iPad. Para disponibilizar seu suplemento em _todos_ os hosts e plataformas do Office, use verificações no tempo de execução em vez do elemento **Requirements**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p118">Only use the **Requirements** element to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the **Requirements** element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the **Requirements** element.</span></span>

<span data-ttu-id="d3f06-188">O exemplo de código a seguir mostra um suplemento que carrega em todos os aplicativos host do Office que oferecem suporte ao seguinte:</span><span class="sxs-lookup"><span data-stu-id="d3f06-188">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="d3f06-189">O conjunto de requisitos **TableBindings**, que tem uma versão mínima de 1.1.</span><span class="sxs-lookup"><span data-stu-id="d3f06-189">**TableBindings** requirement set, which has a minimum version of 1.1.</span></span>
    
-  <span data-ttu-id="d3f06-190">O conjunto de requisitos **OOXML**, que tem uma versão mínima de 1.1.</span><span class="sxs-lookup"><span data-stu-id="d3f06-190">**OOXML** requirement set, which has a minimum version of 1.1.</span></span>
    
-  <span data-ttu-id="d3f06-191">O método **Document.getSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-191">**Document.getSelectedDataAsync** method.</span></span>

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- <span data-ttu-id="d3f06-192">O elemento **Requirements** contém os elementos filhos **Sets** e **Methods**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-192">The  **Requirements** element contains the **Sets** and **Methods** child elements.</span></span>
    
- <span data-ttu-id="d3f06-p119">O elemento **conjuntos** pode conter um ou mais elementos **Definir**. **DefaultMinVersion** especifica o valor padrão de **MinVersion** para todos os elementos filhos de **Definir**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p119">The  **Sets** element can contain one or more **Set** elements. **DefaultMinVersion** specifies the default **MinVersion** value of all child **Set** elements.</span></span>
    
- <span data-ttu-id="d3f06-195">O elemento **Definir** especifica os conjuntos de requisitos que devem receber suporte do host do Office para que o suplemento seja executado.</span><span class="sxs-lookup"><span data-stu-id="d3f06-195">The  **Set** element specifies requirement sets that the Office host must support to run the add-in.</span></span> <span data-ttu-id="d3f06-196">O atributo **Nome** especifica o nome do conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="d3f06-196">The **Name** attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="d3f06-197">A **MinVersion** especifica a versão mínima do conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="d3f06-197">The **MinVersion** specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="d3f06-198">A **MinVersion** substitui o valor de **DefaultMinVersion**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-198">**MinVersion** overrides the value of **DefaultMinVersion**.</span></span> <span data-ttu-id="d3f06-199">Para saber mais sobre os conjuntos de requisito e sobre as versões de conjuntos de requisitos aos quais membros de sua API pertencem, confira [Conjuntos de requisitos de suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d3f06-199">For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>
    
- <span data-ttu-id="d3f06-p121">O elemento **métodos** pode conter um ou mais elementos **métodos**. Você não pode usar o elemento **métodos** com suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p121">The  **Methods** element can contain one or more **Method** elements. You can't use the **Methods** element with Outlook add-ins.</span></span>
    
- <span data-ttu-id="d3f06-p122">O elemento **Methods** especifica um método individual que deve receber suporte no host do Office em que o suplemento é executado. O atributo **Name** é obrigatório e especifica o nome do método qualificado com seu objeto pai.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p122">The  **Method** element specifies an individual method that must be supported in the Office host where your add-in runs. The **Name** attribute is required and specifies the name of the method qualified with its parent object.</span></span>
    

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="d3f06-204">Usar verificações no tempo de execução em seu código JavaScript</span><span class="sxs-lookup"><span data-stu-id="d3f06-204">Use runtime checks in your JavaScript code</span></span>


<span data-ttu-id="d3f06-p123">Se certos conjuntos de requisitos recebem suporte do host do Office, você pode proporcionar outras funcionalidades em seu suplemento. Por exemplo, pode usar a nova API JavaScript para Word em seu suplemento existente se o seu suplemento for executado no Word 2016. Para fazer isso, use o método **isSetSupported** com o nome do conjunto de requisitos. **isSetSupported** determinado, no tempo de execução, se o host do Office que está executando o suplemento dá suporte ao conjunto de requisitos. Se houver suporte para o conjunto de requisitos, **isSetSupported** retorna **true** e executa o código adicional que usa os membros da API desse conjunto de requisitos. Se o host do Office não dá suporte ao conjunto de requisitos, **isSetSupported** retorna **false** e o código adicional não é executado. O código a seguir mostra a sintaxe a ser usada com **isSetSupported**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-p123">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host. For example, you might want to use the new Word JavaScript APIs Word in your existing add-in if your add-in runs in Word 2016. To do this, you use the  **isSetSupported** method with the name of the requirement set. **isSetSupported** determines, at runtime, whether the Office host running the add-in supports the requirement set. If the requirement set is supported, **isSetSupported** returns **true** and runs the additional code that uses the API members from that requirement set. If the Office host doesn't support the requirement set, **isSetSupported** returns **false** and the additional code won't run. The following code shows the syntax to use with **isSetSupported**.</span></span>


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  <span data-ttu-id="d3f06-212">_RequirementSetName_ (obrigatório) é uma cadeia de caracteres que representa o nome do conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="d3f06-212">_RequirementSetName_ (required) is a string that represents the name of the requirement set.</span></span> <span data-ttu-id="d3f06-213">Para saber mais sobre os conjuntos de requisitos disponíveis, confira [Conjuntos de requisitos de Suplemento do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d3f06-213">For more information about available requirement sets, see [Office Add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>
    
-  <span data-ttu-id="d3f06-214">_VersionNumber_ (opcional) é a versão do conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="d3f06-214">_VersionNumber_ (optional) is the version of the requirement set.</span></span>
    
<span data-ttu-id="d3f06-p125">No Excel 2016 ou no Word 2016, use **isSetSupported** com os conjuntos de requisitos **ExcelAPI** ou **WordAPI**. O método **isSetSupported** e os conjuntos de requisitos **ExcelAP**I e **WordAPI** estão disponíveis no Office.js mais recente na CDN. Se você não usa o Office.js da CDN, seu suplemento pode gerar exceções, pois **isSetSupported** fica indefinido. Para saber mais, confira [Especificar a biblioteca de API JavaScript para Office mais recente](#specify-the-latest-javascript-api-for-office-library).</span><span class="sxs-lookup"><span data-stu-id="d3f06-p125">In Excel 2016 or Word 2016, use  **isSetSupported** with the **ExcelAPI** or **WordAPI** requirement sets. The **isSetSupported** method, and the **ExcelAPI** and **WordAPI** requirement sets, are available in the latest Office.js file available from the CDN. If you don't use Office.js from the CDN, your add-in might generate exceptions because **isSetSupported** will be undefined. For more information, see [Specify the latest JavaScript API for Office library](#specify-the-latest-javascript-api-for-office-library).</span></span> 


> [!NOTE]
> <span data-ttu-id="d3f06-p126">**isSetSupported** não funciona no Outlook ou no Outlook Web App. Para usar uma verificação no tempo de execução no Outlook ou no Outlook Web App, use a técnica descrita em [Verificações no tempo de execução usando métodos que não fazem parte de um conjunto de requisitos](#runtime-checks-using-methods-not-in-a-requirement-set).</span><span class="sxs-lookup"><span data-stu-id="d3f06-p126">**isSetSupported** does not work in Outlook or Outlook Web App. To use a runtime check in Outlook or Outlook Web App, use the technique described in [Runtime checks using methods not in a requirement set](#runtime-checks-using-methods-not-in-a-requirement-set).</span></span>

<span data-ttu-id="d3f06-221">O exemplo de código a seguir mostra como um suplemento pode fornecer outras funcionalidades para hosts do Office diferentes que podem dar suporte a conjuntos de requisitos ou membros de API diferentes.</span><span class="sxs-lookup"><span data-stu-id="d3f06-221">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>




```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
{
    // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
      // Run code that uses API members from the CustomXmlParts requirement set.
}
else 
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="d3f06-222">Verificações no tempo de execução usando métodos que não fazem parte de um conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="d3f06-222">Runtime checks using methods not in a requirement set</span></span>


<span data-ttu-id="d3f06-223">Alguns membros de API não pertencem a conjuntos de requisitos.</span><span class="sxs-lookup"><span data-stu-id="d3f06-223">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="d3f06-224">Isso aplica-se apenas a membros da API que fazem parte do namespace [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) (qualquer coisa abaixo de Office.), não a membros de API que pertencem a namespaces da API JavaScript para Word (qualquer coisa em Word.) ou da [Referência sobre a API JavaScript para suplementos do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) (qualquer coisa em Excel.).</span><span class="sxs-lookup"><span data-stu-id="d3f06-224">This only applies to API members that are part of the [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) namespace (anything under Office.), not API members that belong to the Word JavaScript API (anything in Word.) or [Excel add-ins JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) (anything in Excel.) namespaces.</span></span> <span data-ttu-id="d3f06-225">Quando seu suplemento depende de um método que não faz parte de um conjunto de requisitos, é possível usar a verificação no tempo de execução para determinar se o método tem suporte no host do Office, conforme mostra o exemplo de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="d3f06-225">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example.</span></span> <span data-ttu-id="d3f06-226">Para obter uma lista completa dos métodos que não pertencem a um conjunto de requisitos, confira [Conjuntos de requisitos de Suplemento do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d3f06-226">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>


> [!NOTE]
> <span data-ttu-id="d3f06-227">Recomendamos limitar o uso desse tipo de verificação no tempo de execução no código de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d3f06-227">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="d3f06-228">O exemplo de código a seguir verifica se o host oferece suporte a **document.setSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="d3f06-228">The following code example checks whether the host supports  **document.setSelectedDataAsync**.</span></span>




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="d3f06-229">Confira também</span><span class="sxs-lookup"><span data-stu-id="d3f06-229">See also</span></span>

- [<span data-ttu-id="d3f06-230">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d3f06-230">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="d3f06-231">Conjuntos de requisitos de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="d3f06-231">Office Add-in requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="d3f06-232">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="d3f06-232">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
