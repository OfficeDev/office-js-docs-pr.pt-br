---
title: Abrir automaticamente um painel de tarefas com um documento
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: a231255200d6edd1fc923a82711c8c24819bf914
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448774"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a><span data-ttu-id="fb1b2-102">Abrir automaticamente um painel de tarefas com um documento</span><span class="sxs-lookup"><span data-stu-id="fb1b2-102">Automatically open a task pane with a document</span></span>

<span data-ttu-id="fb1b2-p101">Você pode usar comandos de suplemento no seu Suplemento do Office para estender a interface do usuário do Office adicionando botões à faixa de opções do Office. Quando os usuários clicam no botão de comando, ocorre uma ação, como abrir um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p101">You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office ribbon. When users click your command button, an action occurs, such as opening a task pane.</span></span>

<span data-ttu-id="fb1b2-105">Alguns cenários exigem que um painel de tarefas seja exibido automaticamente ao abrir um documento, sem a interação explícita do usuário.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-105">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction.</span></span> <span data-ttu-id="fb1b2-106">Você pode usar o recurso autoopen do painel de tarefas, apresentado no conjunto de requisitos AddInCommands 1.1, para abrir automaticamente um painel de tarefas quando necessário.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-106">You can use the autoopen task pane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it.</span></span>


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a><span data-ttu-id="fb1b2-107">De que forma o recurso autoopen é diferente da inserção de um painel de tarefas?</span><span class="sxs-lookup"><span data-stu-id="fb1b2-107">How is the autoopen feature different from inserting a task pane?</span></span>

<span data-ttu-id="fb1b2-p103">Quando um usuário lançar suplementos que não usam comandos de suplemento, por exemplo, suplementos que são executados no Office 2013, eles serão inseridos no documento e persistirão nesse documento. Como resultado, quando outros usuários abrem o documento, é solicitado que eles instalem o suplemento, e o painel de tarefas abrirá. O desafio com esse modelo é que, em muitos casos, os usuários não querem que o suplemento persista no documento. Por exemplo, um aluno que usa um suplemento de dicionário em um documento do Word pode não querer que seus colegas ou professores sejam avisados para instalar esse suplemento quando abrirem o documento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p103">When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document. As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens. The challenge with this model is that in many cases, users don’t want the add-in to persist in the document. For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.</span></span>

<span data-ttu-id="fb1b2-112">Com o recurso autoopen, você pode explicitamente definir, ou permitir que o usuário defina, se um suplemento do painel de tarefas irá persistir em um documento específico.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-112">With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document.</span></span>

## <a name="support-and-availability"></a><span data-ttu-id="fb1b2-113">Suporte e disponibilidade</span><span class="sxs-lookup"><span data-stu-id="fb1b2-113">Support and availability</span></span>

<span data-ttu-id="fb1b2-114">O recurso autoopen é atualmente</span><span class="sxs-lookup"><span data-stu-id="fb1b2-114">The autoopen feature is currently</span></span> <!-- in **developer preview** and it is only --> <span data-ttu-id="fb1b2-115">suportado pelos seguintes produtos e plataformas.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-115">supported in the following products and platforms.</span></span>

|<span data-ttu-id="fb1b2-116">**Produtos**</span><span class="sxs-lookup"><span data-stu-id="fb1b2-116">**Products**</span></span>|<span data-ttu-id="fb1b2-117">**Plataformas**</span><span class="sxs-lookup"><span data-stu-id="fb1b2-117">**Platforms**</span></span>|
|:-----------|:------------|
|<ul><li><span data-ttu-id="fb1b2-118">Word</span><span class="sxs-lookup"><span data-stu-id="fb1b2-118">Word</span></span></li><li><span data-ttu-id="fb1b2-119">Excel</span><span class="sxs-lookup"><span data-stu-id="fb1b2-119">Excel</span></span></li><li><span data-ttu-id="fb1b2-120">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fb1b2-120">PowerPoint</span></span></li></ul>|<span data-ttu-id="fb1b2-121">Plataformas compatíveis com todos os produtos:</span><span class="sxs-lookup"><span data-stu-id="fb1b2-121">Supported platforms for all products:</span></span><ul><li><span data-ttu-id="fb1b2-p104">Office para Windows Desktop. Build 16.0.8121.1000+</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p104">Office for Windows Desktop. Build 16.0.8121.1000+</span></span></li><li><span data-ttu-id="fb1b2-p105">Office para Mac. Versão 15.34.17051500+</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p105">Office for Mac. Build 15.34.17051500+</span></span></li><li><span data-ttu-id="fb1b2-126">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb1b2-126">Office Online</span></span></li></ul>|


## <a name="best-practices"></a><span data-ttu-id="fb1b2-127">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="fb1b2-127">Best practices</span></span>

<span data-ttu-id="fb1b2-128">Aplique as seguintes práticas recomendadas ao usar o recurso autoopen:</span><span class="sxs-lookup"><span data-stu-id="fb1b2-128">Apply the following best practices when you use the autoopen feature:</span></span>

- <span data-ttu-id="fb1b2-129">Use o recurso autoopen quando ele auxiliar a eficiência dos usuários do seu suplemento, como</span><span class="sxs-lookup"><span data-stu-id="fb1b2-129">Use the autoopen feature when it will help make your add-in users more efficient, such as:</span></span>
  - <span data-ttu-id="fb1b2-p106">Quando o documento precisa do suplemento para funcionar corretamente. Por exemplo, uma planilha que inclui valores de ações que são atualizados periodicamente por um suplemento. O suplemento deverá abrir automaticamente quando a planilha for aberta para manter os valores atualizados.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p106">When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span></span>
  - <span data-ttu-id="fb1b2-p107">Quando é muito provável que o usuário sempre utilizará o suplemento com um determinado documento. Por exemplo, um suplemento que ajuda os usuários a preencher ou alterar dados em um documento puxando informações de um sistema de back-end.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p107">When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span></span>
- <span data-ttu-id="fb1b2-p108">Permita que os usuários ativem ou desativem o recurso autoopen. Inclua uma opção em sua interface de usuário para que eles possam escolher quando não querem mais que o suplemento abra automaticamente no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p108">Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span></span>  
- <span data-ttu-id="fb1b2-137">Use a detecção de configuração de exigência para determinar se o recurso autoopen está disponível e fornecer um comportamento de fallback se ele não estiver disponível.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-137">Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn’t.</span></span>
- <span data-ttu-id="fb1b2-p109">Não use o recurso autoopen para aumentar artificialmente o uso do seu suplemento. Se não faz sentido seu suplemento abrir automaticamente em determinados documentos, esse recurso pode incomodar os usuários.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p109">Don't use the autoopen feature to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to open automatically with certain documents, this feature can annoy users.</span></span>

    > [!NOTE]
    > <span data-ttu-id="fb1b2-140">Se a Microsoft detectar abuso do recurso autoopen, seu suplemento poderá ser rejeitado no AppSource.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-140">If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource.</span></span>

- <span data-ttu-id="fb1b2-p110">Não use esse recurso para fixar vários painéis de tarefas. Você só pode definir um painel do suplemento para abrir automaticamente com um documento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p110">Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.</span></span>  

## <a name="implementation"></a><span data-ttu-id="fb1b2-143">Implementação</span><span class="sxs-lookup"><span data-stu-id="fb1b2-143">Implementation</span></span>

<span data-ttu-id="fb1b2-144">Para implementar o recurso autoopen:</span><span class="sxs-lookup"><span data-stu-id="fb1b2-144">To implement the autoopen feature:</span></span>

- <span data-ttu-id="fb1b2-145">Especifique o painel de tarefas a ser aberto automaticamente.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-145">Specify the task pane to be opened automatically.</span></span>
- <span data-ttu-id="fb1b2-146">Marque o documento para abrir o painel de tarefas automaticamente.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-146">Tag the document to automatically open the task pane.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fb1b2-p111">O painel que você designar para abrir automaticamente só será aberto se o suplemento já estiver instalado no dispositivo do usuário. Se o usuário não tiver o suplemento instalado quando abrir um documento, o recurso autoopen não funcionará, e a configuração será ignorada. Se você também exigir que o suplemento seja distribuído com o documento, será preciso definir a propriedade de visibilidade como 1. Isso só pode ser feito usando OpenXML. Um exemplo será fornecido posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p111">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span></span>

### <a name="step-1-specify-the-task-pane-to-open"></a><span data-ttu-id="fb1b2-150">Etapa 1: especificar o painel de tarefas que será aberto</span><span class="sxs-lookup"><span data-stu-id="fb1b2-150">Step 1: Specify the task pane to open</span></span>

<span data-ttu-id="fb1b2-p112">Para especificar o painel de tarefas que será aberto automaticamente, defina o valor [TaskpaneId](/office/dev/add-ins/reference/manifest/action#taskpaneid) para **Office.AutoShowTaskpaneWithDocument**. Você só pode definir esse valor em um painel de tarefas. Se você definir esse valor em vários painéis de tarefas, a primeira ocorrência do valor será reconhecida e as outras serão ignoradas.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p112">To specify the task pane to open automatically, set the [TaskpaneId](/office/dev/add-ins/reference/manifest/action#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span></span>

<span data-ttu-id="fb1b2-154">O exemplo a seguir mostra o valor TaskPaneId configurado para Office.AutoShowTaskpaneWithDocument.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-154">The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a><span data-ttu-id="fb1b2-155">Etapa 2: marcar o documento para abrir o painel de tarefas automaticamente</span><span class="sxs-lookup"><span data-stu-id="fb1b2-155">Step 2: Tag the document to automatically open the task pane</span></span>

<span data-ttu-id="fb1b2-p113">Você pode marcar o documento para acionar o recurso autoopen de duas maneiras. Escolha a alternativa que funciona melhor para o seu cenário.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p113">You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.</span></span>  


#### <a name="tag-the-document-on-the-client-side"></a><span data-ttu-id="fb1b2-158">Marcar o documento no lado do cliente</span><span class="sxs-lookup"><span data-stu-id="fb1b2-158">Tag the document on the client side</span></span>

<span data-ttu-id="fb1b2-159">Use o método [settings.set](/javascript/api/office/office.settings) do Office.js para configurar o **Office.AutoShowTaskpaneWithDocument** para **true**, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-159">Use the Office.js [settings.set](/javascript/api/office/office.settings) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.</span></span>

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

<span data-ttu-id="fb1b2-160">Use esse método se você precisar marcar o documento como parte da interação com o suplemento (por exemplo, assim que o usuário criar uma ligação ou escolher uma opção para indicar que deseja que o painel abra automaticamente).</span><span class="sxs-lookup"><span data-stu-id="fb1b2-160">Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).</span></span>

#### <a name="use-open-xml-to-tag-the-document"></a><span data-ttu-id="fb1b2-161">Usar Open XML para marcar o documento</span><span class="sxs-lookup"><span data-stu-id="fb1b2-161">Use Open XML to tag the document</span></span>

<span data-ttu-id="fb1b2-p114">Você pode usar o Open XML para criar ou modificar um documento e adicionar a marcação XML do Open Office apropriada para acionar o recurso autoopen. Veja um exemplo de como fazer isso em [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p114">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span></span>

<span data-ttu-id="fb1b2-164">Adicione duas partes do Open XML ao documento:</span><span class="sxs-lookup"><span data-stu-id="fb1b2-164">Add two Open XML parts to the document:</span></span>

- <span data-ttu-id="fb1b2-165">Uma parte `webextension`</span><span class="sxs-lookup"><span data-stu-id="fb1b2-165">A `webextension` part</span></span>
- <span data-ttu-id="fb1b2-166">Uma parte `taskpane`</span><span class="sxs-lookup"><span data-stu-id="fb1b2-166">A `taskpane` part</span></span>

<span data-ttu-id="fb1b2-167">O exemplo a seguir mostra como adicionar a parte `webextension`.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-167">The following example shows how to add the `webextension` part.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
   <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="fb1b2-168">A parte `webextension` inclui um conjunto de propriedades e uma propriedade chamada **Office.AutoShowTaskpaneWithDocument** que deve ser definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-168">The `webextension` part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.</span></span>

<span data-ttu-id="fb1b2-169">A parte `webextension` também inclui uma referência para a loja ou o catálogo com atributos para `id`, `storeType`, `store` e `version`.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-169">The `webextension` part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`.</span></span> <span data-ttu-id="fb1b2-170">Dos valores `storeType`, somente quatro são relevantes para o recurso autoopen.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-170">Of the `storeType` values, only four are relevant to the autoopen feature.</span></span> <span data-ttu-id="fb1b2-171">Os valores dos outros três atributos dependem do valor de `storeType`, conforme mostrado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-171">The values for the other three attributes depend on the value for `storeType`, as shown in the following table.</span></span>

| <span data-ttu-id="fb1b2-172">**valor `storeType`**</span><span class="sxs-lookup"><span data-stu-id="fb1b2-172">**`storeType` value**</span></span> | <span data-ttu-id="fb1b2-173">**valor `id`**</span><span class="sxs-lookup"><span data-stu-id="fb1b2-173">**`id` value**</span></span>    |<span data-ttu-id="fb1b2-174">**valor `store`**</span><span class="sxs-lookup"><span data-stu-id="fb1b2-174">**`store` value**</span></span> | <span data-ttu-id="fb1b2-175">**valor `version`**</span><span class="sxs-lookup"><span data-stu-id="fb1b2-175">**`version` value**</span></span>|
|:---------------|:---------------|:---------------|:---------------|
|<span data-ttu-id="fb1b2-176">OMEX (AppSource)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-176">OMEX (AppSource)</span></span>|<span data-ttu-id="fb1b2-177">A ID do ativo do suplemento no AppSource (confira a observação)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-177">The AppSource asset ID of the add-in (see Note)</span></span>|<span data-ttu-id="fb1b2-178">A localidade do AppSource, por exemplo, "pt-br".</span><span class="sxs-lookup"><span data-stu-id="fb1b2-178">The locale of AppSource; for example, "en-us".</span></span>|<span data-ttu-id="fb1b2-179">A versão no catálogo do AppSource (confira a observação)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-179">The version in the AppSource catalog (see Note)</span></span>|
|<span data-ttu-id="fb1b2-180">FileSystem (um compartilhamento de rede)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-180">FileSystem (a network share)</span></span>|<span data-ttu-id="fb1b2-181">O GUID do suplemento no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-181">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="fb1b2-182">O caminho do compartilhamento de rede. Por exemplo, "\\\\Meu Computador\\Minha Pasta Compartilhada".</span><span class="sxs-lookup"><span data-stu-id="fb1b2-182">The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span>|<span data-ttu-id="fb1b2-183">A versão no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-183">The version in the add-in manifest.</span></span>|
|<span data-ttu-id="fb1b2-184">EXCatalog (implantação por meio do servidor Exchange)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-184">EXCatalog (deployment via the Exchange server)</span></span> |<span data-ttu-id="fb1b2-185">O GUID do suplemento no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-185">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="fb1b2-186">"EXCatalog".</span><span class="sxs-lookup"><span data-stu-id="fb1b2-186">"EXCatalog".</span></span> <span data-ttu-id="fb1b2-187">A linha EXCatalog deve ser usada com o suplemento que usa a Implantação Centralizada no Centro de administração do Office 365.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-187">EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Office 365 admin center.</span></span>|<span data-ttu-id="fb1b2-188">A versão no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-188">The version in the add-in manifest.</span></span>
|<span data-ttu-id="fb1b2-189">Registro (registro de sistema)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-189">Registry (System registry)</span></span>|<span data-ttu-id="fb1b2-190">O GUID do suplemento no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-190">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="fb1b2-191">"developer"</span><span class="sxs-lookup"><span data-stu-id="fb1b2-191">"developer"</span></span>|<span data-ttu-id="fb1b2-192">A versão no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-192">The version in the add-in manifest.</span></span>|

> [!NOTE]
> <span data-ttu-id="fb1b2-p117">Para localizar a ID de ativos e a versão de um suplemento no AppSource, vá para a página inicial do suplemento no AppSource. A ID de ativo aparece na barra de endereços no navegador. A versão aparece na seção **Detalhes** da página.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p117">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.</span></span>

<span data-ttu-id="fb1b2-196">Saiba mais sobre a marcação webextension em [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span><span class="sxs-lookup"><span data-stu-id="fb1b2-196">For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span></span>

<span data-ttu-id="fb1b2-197">O exemplo a seguir mostra como adicionar a parte `taskpane`.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-197">The following example shows how to add the `taskpane` part.</span></span>

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

<span data-ttu-id="fb1b2-198">Observe que neste exemplo, o atributo `visibility` está definido como "0".</span><span class="sxs-lookup"><span data-stu-id="fb1b2-198">Note that in this example, the `visibility` attribute is set to "0".</span></span> <span data-ttu-id="fb1b2-199">Isso significa que, após adicionar as partes webextension e `taskpane`, a primeira vez que o documento for aberto, o usuário deve instalar o suplemento clicando no botão **Suplemento** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-199">This means that after the webextension and `taskpane` parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon.</span></span> <span data-ttu-id="fb1b2-200">Depois disso, o painel de tarefas do suplemento abre automaticamente quando o arquivo for aberto.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-200">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="fb1b2-201">E, ao definir `visibility` como "0", é possível usar o Office.js para permitir que os usuários ativem ou desativem o recurso autoopen.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-201">Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="fb1b2-202">Especificamente, seu script define a configuração de documento **Office.AutoShowTaskpaneWithDocument** como `true` ou `false`.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-202">Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`.</span></span> <span data-ttu-id="fb1b2-203">(Saiba mais em [Marcar o documento no lado do cliente](#tag-the-document-on-the-client-side).)</span><span class="sxs-lookup"><span data-stu-id="fb1b2-203">(For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).)</span></span>

<span data-ttu-id="fb1b2-p119">Se o elemento `visibility` é definido como "1", o painel de tarefas abrirá automaticamente na primeira vez em que o documento for aberto. O usuário é solicitado a confiar no suplemento e, quando a confiança é concedida, o suplemento é aberto. Depois disso, o painel de tarefas do suplemento abrirá automaticamente quando o arquivo for aberto. Entretanto, ao definir `visibility` como "1", não é possível usar o Office.js para permitir que os usuários ativem ou desativem o recurso autoopen.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p119">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span></span>

<span data-ttu-id="fb1b2-208">Definir o `visibility` como "1" é uma boa opção quando o suplemento e o modelo ou o conteúdo do documento são muito estreitamente integrados de modo que o usuário não poderia optar por cancelar o recurso autoopen.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-208">Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature.</span></span>

> [!NOTE]
> <span data-ttu-id="fb1b2-p120">Se quiser distribuir seu suplemento com o documento, para que os usuários sejam solicitados a instalá-lo, você deverá definir a propriedade de visibilidade para 1. Isso só pode ser feito pelo Open XML.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p120">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.</span></span>

<span data-ttu-id="fb1b2-p121">Uma maneira fácil de escrever o XML é primeiro executar seu suplemento e [marcar o documento no lado do cliente](#tag-the-document-on-the-client-side) para escrever o valor e, em seguida, salvar o documento e inspecionar o XML que é gerado. O Office detectará e fornecerá os valores de atributo apropriados. Você também pode usar a [Ferramenta de Produtividade Open XML SDK 2.5](https://www.microsoft.com/download/details.aspx?id=30425) para gerar o código C# para adicionar por meio de programação a marcação com base no XML que você gerou.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-p121">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span></span>

## <a name="test-and-verify-opening-task-panes"></a><span data-ttu-id="fb1b2-214">Testar e verificar a abertura de painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="fb1b2-214">Test and verify opening task panes</span></span>

<span data-ttu-id="fb1b2-215">Você pode implantar uma versão de teste do suplemento que abre automaticamente um painel de tarefas usando a implantação centralizada por meio do Centro de administração do Office 365.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-215">You can deploy a test version of your add-in that will automatically open a task pane using Centralized Deployment via the Office 365 admin center.</span></span> <span data-ttu-id="fb1b2-216">O exemplo a seguir mostra como os suplementos são inseridos do catálogo de Implantação Centralizada usando a versão de armazenamento EXCatalog.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-216">The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="fb1b2-217">Para testar o exemplo anterior, considere participar do [Programa para Desenvolvedores do Office 365](/office/developer-program/office-365-developer-program) e inscreva-se para uma [conta de desenvolvedor do Office 365](https://developer.microsoft.com/office/dev-program), caso ainda não tenha uma assinatura do Office 365.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-217">To test the previous example, please consider joining the [Office 365 Developer Program](/office/developer-program/office-365-developer-program) and signing up for an [Office 365 developer account](https://developer.microsoft.com/office/dev-program) if you don't already own an Office 365 subscription.</span></span> <span data-ttu-id="fb1b2-218">Você pode realmente testar a Implantação Centralizada e verificar se o suplemento funciona como esperado.</span><span class="sxs-lookup"><span data-stu-id="fb1b2-218">You can actually test drive Centralized Deployment and verify that your add-in works as expected.</span></span>


## <a name="see-also"></a><span data-ttu-id="fb1b2-219">Confira também</span><span class="sxs-lookup"><span data-stu-id="fb1b2-219">See also</span></span>

<span data-ttu-id="fb1b2-220">Para ver um exemplo que mostra como usar o recurso autoopen, consulte os [exemplos de comandos do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).</span><span class="sxs-lookup"><span data-stu-id="fb1b2-220">For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).</span></span>
<span data-ttu-id="fb1b2-221">[Ingressar no Programa para Desenvolvedores do Office 365](/office/developer-program/office-365-developer-program).</span><span class="sxs-lookup"><span data-stu-id="fb1b2-221">[Join the Office 365 developer program](/office/developer-program/office-365-developer-program).</span></span>
