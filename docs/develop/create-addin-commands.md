---
title: Criar comandos de suplemento em seu manifesto para Excel, PowerPoint e Word
description: Use VersionOverrides in your manifest to define add-in commands for Excel, PowerPoint, and Word. Use add-in commands to create UI elements, add buttons or lists, and perform actions.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 3bcd3c6e07cdb9899601403e68e80e8d609d2e6e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093711"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a><span data-ttu-id="c9f6c-104">Criar comandos de suplemento em seu manifesto para Excel, PowerPoint e Word</span><span class="sxs-lookup"><span data-stu-id="c9f6c-104">Create add-in commands in your manifest for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="c9f6c-105">Use **[VersionOverrides](../reference/manifest/versionoverrides.md)** in your manifest to define add-in commands for Excel, PowerPoint, and Word.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-105">Use **[VersionOverrides](../reference/manifest/versionoverrides.md)** in your manifest to define add-in commands for Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="c9f6c-106">Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-106">Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions.</span></span> <span data-ttu-id="c9f6c-107">You can use add-in commands to:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-107">You can use add-in commands to:</span></span>

- <span data-ttu-id="c9f6c-108">Criar elementos de interface do usuário ou pontos de entrada que facilitam o uso da funcionalidade dos suplementos.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-108">Create UI elements or entry points that make your add-in's functionality easier to use.</span></span>
- <span data-ttu-id="c9f6c-109">Adicionar botões ou uma lista suspensa de botões à faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-109">Add buttons or a drop-down list of buttons to the ribbon.</span></span>
- <span data-ttu-id="c9f6c-110">Adicionar itens de menu individuais — cada um contendo submenus opcionais — aos menus de contexto específicos (atalho).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-110">Add individual menu items — each containing optional submenus — to specific context (shortcut) menus.</span></span>
- <span data-ttu-id="c9f6c-111">Perform actions when your add-in command is chosen.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-111">Perform actions when your add-in command is chosen.</span></span> <span data-ttu-id="c9f6c-112">You can:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-112">You can:</span></span>
  - <span data-ttu-id="c9f6c-113">Show one or more task pane add-ins for users to interact with.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-113">Show one or more task pane add-ins for users to interact with.</span></span> <span data-ttu-id="c9f6c-114">Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-114">Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span></span>

     <span data-ttu-id="c9f6c-115">*ou*</span><span class="sxs-lookup"><span data-stu-id="c9f6c-115">*or*</span></span>

  - <span data-ttu-id="c9f6c-116">Executar código JavaScript, que normalmente é executado sem exibir qualquer interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-116">Run JavaScript code, which normally runs without displaying any UI.</span></span>

<span data-ttu-id="c9f6c-117">This article describes how to edit your manifest to define add-in commands.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-117">This article describes how to edit your manifest to define add-in commands.</span></span> <span data-ttu-id="c9f6c-118">The following diagram shows the hierarchy of elements used to define add-in commands.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-118">The following diagram shows the hierarchy of elements used to define add-in commands.</span></span> <span data-ttu-id="c9f6c-119">These elements are described in more detail in this article.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-119">These elements are described in more detail in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="c9f6c-120">Os comandos de suplemento também são compatíveis com o Outlook.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-120">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="c9f6c-121">Para obter mais informações, consulte [comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md)</span><span class="sxs-lookup"><span data-stu-id="c9f6c-121">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md)</span></span>

<span data-ttu-id="c9f6c-122">The following image is an overview of add-in commands elements in the manifest.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-122">The following image is an overview of add-in commands elements in the manifest.</span></span>
<span data-ttu-id="c9f6c-123">![Overview of add-in commands elements in the manifest](../images/version-overrides.png)</span><span class="sxs-lookup"><span data-stu-id="c9f6c-123">![Overview of add-in commands elements in the manifest](../images/version-overrides.png)</span></span>

## <a name="step-1-start-from-a-sample"></a><span data-ttu-id="c9f6c-124">Etapa 1: iniciar usando uma amostra</span><span class="sxs-lookup"><span data-stu-id="c9f6c-124">Step 1: Start from a sample</span></span>

<span data-ttu-id="c9f6c-125">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-125">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample).</span></span> <span data-ttu-id="c9f6c-126">Optionally, you can create your own manifest by following the steps in this guide.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-126">Optionally, you can create your own manifest by following the steps in this guide.</span></span> <span data-ttu-id="c9f6c-127">You can validate your manifest using the XSD file in the Office Add-in Commands Samples site.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-127">You can validate your manifest using the XSD file in the Office Add-in Commands Samples site.</span></span> <span data-ttu-id="c9f6c-128">Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-128">Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span></span>

## <a name="step-2-create-a-task-pane-add-in"></a><span data-ttu-id="c9f6c-129">Etapa 2: criar um suplemento de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="c9f6c-129">Step 2: Create a task pane add-in</span></span>

<span data-ttu-id="c9f6c-130">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-130">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article.</span></span> <span data-ttu-id="c9f6c-131">You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-131">You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span></span>

<span data-ttu-id="c9f6c-132">The following example shows an Office 2013 add-in's manifest.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-132">The following example shows an Office 2013 add-in's manifest.</span></span> <span data-ttu-id="c9f6c-133">There are no add-in commands in this manifest because there is no **VersionOverrides** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-133">There are no add-in commands in this manifest because there is no **VersionOverrides** element.</span></span> <span data-ttu-id="c9f6c-134">Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-134">Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016.</span></span> <span data-ttu-id="c9f6c-135">In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-135">In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in.</span></span> <span data-ttu-id="c9f6c-136">In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-136">In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in.</span></span> <span data-ttu-id="c9f6c-137">If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-137">If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span></span>
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## <a name="step-3-add-versionoverrides-element"></a><span data-ttu-id="c9f6c-138">Etapa 3: adicionar o elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c9f6c-138">Step 3: Add VersionOverrides element</span></span>

<span data-ttu-id="c9f6c-139">The **VersionOverrides** element is the root element that contains the definition of your add-in command.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-139">The **VersionOverrides** element is the root element that contains the definition of your add-in command.</span></span> <span data-ttu-id="c9f6c-140">**VersionOverrides** is a child element of the **OfficeApp** element in the manifest.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-140">**VersionOverrides** is a child element of the **OfficeApp** element in the manifest.</span></span> <span data-ttu-id="c9f6c-141">The following table lists the attributes of the **VersionOverrides** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-141">The following table lists the attributes of the **VersionOverrides** element.</span></span>

|<span data-ttu-id="c9f6c-142">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-142">**Attribute**</span></span>|<span data-ttu-id="c9f6c-143">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-143">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-144">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-144">**xmlns**</span></span> <br/> | <span data-ttu-id="c9f6c-145">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-145">Required.</span></span> <span data-ttu-id="c9f6c-146">O local do esquema, que deve ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-146">The schema location, which must be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span> <br/> |
|<span data-ttu-id="c9f6c-147">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-147">**xsi:type**</span></span> <br/> |<span data-ttu-id="c9f6c-148">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-148">Required.</span></span> <span data-ttu-id="c9f6c-149">The schema version.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-149">The schema version.</span></span> <span data-ttu-id="c9f6c-150">The version described in this article is "VersionOverridesV1_0".</span><span class="sxs-lookup"><span data-stu-id="c9f6c-150">The version described in this article is "VersionOverridesV1_0".</span></span>  <br/> |

<span data-ttu-id="c9f6c-151">A tabela a seguir identifica os elementos filho de **VersionOverrides**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-151">The following table identifies the child elements of **VersionOverrides**.</span></span>
  
|<span data-ttu-id="c9f6c-152">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-152">**Element**</span></span>|<span data-ttu-id="c9f6c-153">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-153">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-154">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-154">**Description**</span></span> <br/> |<span data-ttu-id="c9f6c-155">Optional.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-155">Optional.</span></span> <span data-ttu-id="c9f6c-156">Describes the add-in.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-156">Describes the add-in.</span></span> <span data-ttu-id="c9f6c-157">This child **Description** element overrides a previous **Description** element in the parent portion of the manifest.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-157">This child **Description** element overrides a previous **Description** element in the parent portion of the manifest.</span></span> <span data-ttu-id="c9f6c-158">The **resid** attribute for this **Description** element is set to the **id** of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-158">The **resid** attribute for this **Description** element is set to the **id** of a **String** element.</span></span> <span data-ttu-id="c9f6c-159">The **String** element contains the text for **Description**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-159">The **String** element contains the text for **Description**.</span></span> <br/> |
|<span data-ttu-id="c9f6c-160">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-160">**Requirements**</span></span> <br/> |<span data-ttu-id="c9f6c-161">Optional.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-161">Optional.</span></span> <span data-ttu-id="c9f6c-162">Specifies the minimum requirement set and version of Office.js that the add-in requires.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-162">Specifies the minimum requirement set and version of Office.js that the add-in requires.</span></span> <span data-ttu-id="c9f6c-163">This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-163">This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest.</span></span> <span data-ttu-id="c9f6c-164">For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-164">For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>  <br/> |
|<span data-ttu-id="c9f6c-165">**Hosts**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-165">**Hosts**</span></span> <br/> |<span data-ttu-id="c9f6c-166">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-166">Required.</span></span> <span data-ttu-id="c9f6c-167">Specifies a collection of Office hosts.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-167">Specifies a collection of Office hosts.</span></span> <span data-ttu-id="c9f6c-168">The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-168">The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest.</span></span> <span data-ttu-id="c9f6c-169">You must include a **xsi:type** attribute set to "Workbook" or "Document".</span><span class="sxs-lookup"><span data-stu-id="c9f6c-169">You must include a **xsi:type** attribute set to "Workbook" or "Document".</span></span> <br/> |
|<span data-ttu-id="c9f6c-170">**Resources**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-170">**Resources**</span></span> <br/> |<span data-ttu-id="c9f6c-171">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-171">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span> <span data-ttu-id="c9f6c-172">For example, the **Description** element's value refers to a child element in **Resources**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-172">For example, the **Description** element's value refers to a child element in **Resources**.</span></span> <span data-ttu-id="c9f6c-173">The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-173">The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article.</span></span> <br/> |

<span data-ttu-id="c9f6c-174">O exemplo a seguir mostra como usar o elemento **VersionOverrides** e seus elementos filho.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-174">The following example shows how to use the **VersionOverrides** element and its child elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a><span data-ttu-id="c9f6c-175">Etapa 4: adicionar os elementos Hosts, Host e DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="c9f6c-175">Step 4: Add Hosts, Host, and DesktopFormFactor elements</span></span>

<span data-ttu-id="c9f6c-176">The **Hosts** element contains one or more **Host** elements.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-176">The **Hosts** element contains one or more **Host** elements.</span></span> <span data-ttu-id="c9f6c-177">A **Host** element specifies a particular Office host.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-177">A **Host** element specifies a particular Office host.</span></span> <span data-ttu-id="c9f6c-178">The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-178">The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host.</span></span> <span data-ttu-id="c9f6c-179">To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-179">To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.</span></span>

<span data-ttu-id="c9f6c-180">O elemento **DesktopFormFactor** especifica as configurações para um suplemento que é executado no Office Online (em um navegador) e no Windows.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-180">The **DesktopFormFactor** element specifies the settings for an add-in that runs in Office on the web (in a browser) and Windows.</span></span>

<span data-ttu-id="c9f6c-181">Veja a seguir um exemplo dos elementos **Hosts**, **Host** e **DesktopFormFactor**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-181">The following is an example of **Hosts**, **Host**, and **DesktopFormFactor** elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-5-add-the-functionfile-element"></a><span data-ttu-id="c9f6c-182">Etapa 5: adicionar o elemento FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c9f6c-182">Step 5: Add the FunctionFile element</span></span>

<span data-ttu-id="c9f6c-183">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](../reference/manifest/control.md#button-control) for a description).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-183">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](../reference/manifest/control.md#button-control) for a description).</span></span> <span data-ttu-id="c9f6c-184">The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-184">The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require.</span></span> <span data-ttu-id="c9f6c-185">You can't link directly to a JavaScript file.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-185">You can't link directly to a JavaScript file.</span></span> <span data-ttu-id="c9f6c-186">You can only link to an HTML file.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-186">You can only link to an HTML file.</span></span> <span data-ttu-id="c9f6c-187">The file name is specified as a **Url** element in the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-187">The file name is specified as a **Url** element in the **Resources** element.</span></span>

<span data-ttu-id="c9f6c-188">Veja a seguir um exemplo do elemento **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-188">The following is an example of the **FunctionFile** element.</span></span>
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint>

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> <span data-ttu-id="c9f6c-189">Verifique se seu código JavaScript chama `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-189">Make sure your JavaScript code calls  `Office.initialize`.</span></span>

<span data-ttu-id="c9f6c-190">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-190">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`.</span></span> <span data-ttu-id="c9f6c-191">The **FunctionName** element (see [Button controls](../reference/manifest/control.md#button-control) for a description) uses the functions in **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-191">The **FunctionName** element (see [Button controls](../reference/manifest/control.md#button-control) for a description) uses the functions in **FunctionFile**.</span></span>

<span data-ttu-id="c9f6c-192">O código a seguir mostra como implementar a função usada por **FunctionName**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-192">The following code shows how to implement the function used by **FunctionName**.</span></span>

```js
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Your function must be in the global namespace.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // Show error message.
                }
                else {
                    // Show success message.
                }
            });

        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
    }
</script>
```

> [!IMPORTANT]
> The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. <span data-ttu-id="c9f6c-197">You must implement **event.completed**, otherwise your function will not run.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-197">You must implement **event.completed**, otherwise your function will not run.</span></span>

## <a name="step-6-add-extensionpoint-elements"></a><span data-ttu-id="c9f6c-198">Etapa 6: adicionar elementos do ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c9f6c-198">Step 6: Add ExtensionPoint elements</span></span>

<span data-ttu-id="c9f6c-199">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-199">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI.</span></span> <span data-ttu-id="c9f6c-200">You can define **ExtensionPoint** elements with these **xsi:type** values:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-200">You can define **ExtensionPoint** elements with these **xsi:type** values:</span></span>

- <span data-ttu-id="c9f6c-201">**PrimaryCommandSurface**, que se refere à faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-201">**PrimaryCommandSurface**, which refers to the ribbon in Office.</span></span>

- <span data-ttu-id="c9f6c-202">**ContextMenu**, que é o menu de atalho exibido quando você clica com o botão direito na interface do usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-202">**ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="c9f6c-203">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-203">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c9f6c-204">For elements that contain an ID attribute, make sure you provide a unique ID.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-204">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="c9f6c-205">We recommend that you use your company's name along with your ID.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-205">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="c9f6c-206">For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-206">For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span></span> 
  
```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso Tab">
  <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
  <!-- <OfficeTab id="TabData"> -->
    <Label resid="residLabel4" />
    <Group id="Group1Id12">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Button1Id1">

        <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

|<span data-ttu-id="c9f6c-207">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-207">**Element**</span></span>|<span data-ttu-id="c9f6c-208">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-208">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-209">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-209">**CustomTab**</span></span> <br/> |<span data-ttu-id="c9f6c-210">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-210">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="c9f6c-211">If you use the **CustomTab** element, you can't use the **OfficeTab** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-211">If you use the **CustomTab** element, you can't use the **OfficeTab** element.</span></span> <span data-ttu-id="c9f6c-212">The **id** attribute is required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-212">The **id** attribute is required.</span></span> <br/> |
|<span data-ttu-id="c9f6c-213">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-213">**OfficeTab**</span></span> <br/> |<span data-ttu-id="c9f6c-214">Obrigatório se você deseja estender uma guia padrão da faixa de opções do aplicativo do Office (usando **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-214">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="c9f6c-215">Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-215">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <br/> <span data-ttu-id="c9f6c-216">Para obter mais valores de tabulação a serem usados com o atributo **ID** , consulte [Tab Values for the Office app default Ribbon Tabs](../reference/manifest/officetab.md).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-216">For more tab values to use with the **id** attribute, see [Tab values for default Office app ribbon tabs](../reference/manifest/officetab.md).</span></span>  <br/> |
|<span data-ttu-id="c9f6c-217">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-217">**OfficeMenu**</span></span> <br/> | <span data-ttu-id="c9f6c-218">Required if you're adding add-in commands to a default context menu (using **ContextMenu**).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-218">Required if you're adding add-in commands to a default context menu (using **ContextMenu**).</span></span> <span data-ttu-id="c9f6c-219">The **id** attribute must be set to:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-219">The **id** attribute must be set to:</span></span> <br/> <span data-ttu-id="c9f6c-220">**ContextMenuText** for Excel or Word.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-220">**ContextMenuText** for Excel or Word.</span></span> <span data-ttu-id="c9f6c-221">Displays the item on the context menu when text is selected and then the user right-clicks on the selected text.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-221">Displays the item on the context menu when text is selected and then the user right-clicks on the selected text.</span></span> <br/> <span data-ttu-id="c9f6c-222">**ContextMenuCell** for Excel.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-222">**ContextMenuCell** for Excel.</span></span> <span data-ttu-id="c9f6c-223">Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-223">Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span> <br/> |
|<span data-ttu-id="c9f6c-224">**Group**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-224">**Group**</span></span> <br/> |<span data-ttu-id="c9f6c-225">A group of user interface extension points on a tab. A group can have up to six controls.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-225">A group of user interface extension points on a tab. A group can have up to six controls.</span></span> <span data-ttu-id="c9f6c-226">The **id** attribute is required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-226">The **id** attribute is required.</span></span> <span data-ttu-id="c9f6c-227">It's a string with a maximum of 125 characters.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-227">It's a string with a maximum of 125 characters.</span></span> <br/> |
|<span data-ttu-id="c9f6c-228">**Label**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-228">**Label**</span></span> <br/> |<span data-ttu-id="c9f6c-229">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-229">Required.</span></span> <span data-ttu-id="c9f6c-230">The label of the group.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-230">The label of the group.</span></span> <span data-ttu-id="c9f6c-231">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-231">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-232">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-232">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-233">**Icon**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-233">**Icon**</span></span> <br/> |<span data-ttu-id="c9f6c-234">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-234">Required.</span></span> <span data-ttu-id="c9f6c-235">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-235">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="c9f6c-236">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-236">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="c9f6c-237">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-237">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="c9f6c-238">The **size** attribute gives the size, in pixels, of the image.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-238">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="c9f6c-239">Three image sizes are required: 16, 32, and 80.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-239">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="c9f6c-240">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-240">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span> <br/> |
|<span data-ttu-id="c9f6c-241">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-241">**Tooltip**</span></span> <br/> |<span data-ttu-id="c9f6c-242">Optional.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-242">Optional.</span></span> <span data-ttu-id="c9f6c-243">The tooltip of the group.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-243">The tooltip of the group.</span></span> <span data-ttu-id="c9f6c-244">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-244">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-245">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-245">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-246">**Control**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-246">**Control**</span></span> <br/> |<span data-ttu-id="c9f6c-247">Each group requires at least one control.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-247">Each group requires at least one control.</span></span> <span data-ttu-id="c9f6c-248">A **Control** element can be either a **Button** or a **Menu**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-248">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="c9f6c-249">Use **Menu** to specify a drop-down list of button controls.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-249">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="c9f6c-250">Currently, only buttons and menus are supported.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-250">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="c9f6c-251">See the  [Button controls](../reference/manifest/control.md#button-control) and [Menu controls](../reference/manifest/control.md#menu-dropdown-button-controls) sections for more information.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-251">See the  [Button controls](../reference/manifest/control.md#button-control) and [Menu controls](../reference/manifest/control.md#menu-dropdown-button-controls) sections for more information.</span></span> <br/><span data-ttu-id="c9f6c-252">**Observação:** para facilitar a solução de problemas, recomendamos adicionar um elemento **Control** e os elementos filho **Resources** relacionados, um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-252">**Note:** To make troubleshooting easier, we recommend that you add a **Control** element and the related **Resources** child elements one at a time.</span></span>          |

### <a name="button-controls"></a><span data-ttu-id="c9f6c-253">Controles de botão</span><span class="sxs-lookup"><span data-stu-id="c9f6c-253">Button controls</span></span>

<span data-ttu-id="c9f6c-254">A button performs a single action when the user selects it.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-254">A button performs a single action when the user selects it.</span></span> <span data-ttu-id="c9f6c-255">It can either execute a JavaScript function or show a task pane.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-255">It can either execute a JavaScript function or show a task pane.</span></span> <span data-ttu-id="c9f6c-256">The following example shows how to define two buttons.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-256">The following example shows how to define two buttons.</span></span> <span data-ttu-id="c9f6c-257">The first button runs a JavaScript function without showing a UI, and the second button shows a task pane.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-257">The first button runs a JavaScript function without showing a UI, and the second button shows a task pane.</span></span> <span data-ttu-id="c9f6c-258">In the **Control** element:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-258">In the **Control** element:</span></span>

- <span data-ttu-id="c9f6c-259">O atributo **type** é obrigatório e deve ser definido como **Button**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-259">The **type** attribute is required, and must be set to **Button**.</span></span>

- <span data-ttu-id="c9f6c-260">O atributo **id** do elemento **Control** é uma cadeia de caracteres com, no máximo, 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-260">The **id** attribute of the **Control** element is a string with a maximum of 125 characters.</span></span>

```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|<span data-ttu-id="c9f6c-261">**Elementos**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-261">**Elements**</span></span>|<span data-ttu-id="c9f6c-262">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-262">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-263">**Label**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-263">**Label**</span></span> <br/> |<span data-ttu-id="c9f6c-264">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-264">Required.</span></span> <span data-ttu-id="c9f6c-265">The text for the button.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-265">The text for the button.</span></span> <span data-ttu-id="c9f6c-266">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-266">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-267">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-267">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-268">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-268">**Tooltip**</span></span> <br/> |<span data-ttu-id="c9f6c-269">Optional.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-269">Optional.</span></span> <span data-ttu-id="c9f6c-270">The tooltip for the button.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-270">The tooltip for the button.</span></span> <span data-ttu-id="c9f6c-271">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-271">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-272">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-272">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-273">**Supertip**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-273">**Supertip**</span></span> <br/> | <span data-ttu-id="c9f6c-274">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-274">Required.</span></span> <span data-ttu-id="c9f6c-275">The supertip for this button, which is defined by the following:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-275">The supertip for this button, which is defined by the following:</span></span> <br/> <span data-ttu-id="c9f6c-276">**Título**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-276">**Title**</span></span> <br/>  <span data-ttu-id="c9f6c-277">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-277">Required.</span></span> <span data-ttu-id="c9f6c-278">The text for the supertip.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-278">The text for the supertip.</span></span> <span data-ttu-id="c9f6c-279">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-279">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-280">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-280">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span> <br/> <span data-ttu-id="c9f6c-281">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-281">**Description**</span></span> <br/>  <span data-ttu-id="c9f6c-282">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-282">Required.</span></span> <span data-ttu-id="c9f6c-283">The description for the supertip.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-283">The description for the supertip.</span></span> <span data-ttu-id="c9f6c-284">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-284">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-285">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-285">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-286">**Icon**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-286">**Icon**</span></span> <br/> | <span data-ttu-id="c9f6c-287">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-287">Required.</span></span> <span data-ttu-id="c9f6c-288">Contains the **Image** elements for the button.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-288">Contains the **Image** elements for the button.</span></span> <span data-ttu-id="c9f6c-289">Image files must be .png format.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-289">Image files must be .png format.</span></span> <br/> <span data-ttu-id="c9f6c-290">**Imagem**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-290">**Image**</span></span> <br/>  <span data-ttu-id="c9f6c-291">Defines an image to display on the button.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-291">Defines an image to display on the button.</span></span> <span data-ttu-id="c9f6c-292">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-292">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="c9f6c-293">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-293">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="c9f6c-294">The **size** attribute indicates the size, in pixels, of the image.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-294">The **size** attribute indicates the size, in pixels, of the image.</span></span> <span data-ttu-id="c9f6c-295">Three image sizes are required: 16, 32, and 80.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-295">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="c9f6c-296">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-296">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span> <br/> |
|<span data-ttu-id="c9f6c-297">**Action**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-297">**Action**</span></span> <br/> | <span data-ttu-id="c9f6c-298">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-298">Required.</span></span> <span data-ttu-id="c9f6c-299">Specifies the action to perform when the user selects the button.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-299">Specifies the action to perform when the user selects the button.</span></span> <span data-ttu-id="c9f6c-300">You can specify one of the following values for the **xsi:type** attribute:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-300">You can specify one of the following values for the **xsi:type** attribute:</span></span> <br/> <span data-ttu-id="c9f6c-301">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-301">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**.</span></span> <span data-ttu-id="c9f6c-302">**ExecuteFunction** does not display a UI.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-302">**ExecuteFunction** does not display a UI.</span></span> <span data-ttu-id="c9f6c-303">The **FunctionName** child element specifies the name of the function to execute.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-303">The **FunctionName** child element specifies the name of the function to execute.</span></span> <br/> <span data-ttu-id="c9f6c-304">**ShowTaskPane**, which shows a task pane add-in.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-304">**ShowTaskPane**, which shows a task pane add-in.</span></span> <span data-ttu-id="c9f6c-305">The **SourceLocation** child element specifies the source file location of the task pane add-in to display.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-305">The **SourceLocation** child element specifies the source file location of the task pane add-in to display.</span></span> <span data-ttu-id="c9f6c-306">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-306">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element.</span></span> <br/> |

### <a name="menu-controls"></a><span data-ttu-id="c9f6c-307">Controles de menu</span><span class="sxs-lookup"><span data-stu-id="c9f6c-307">Menu controls</span></span>

<span data-ttu-id="c9f6c-308">Um controle **Menu** pode ser usado com **PrimaryCommandSurface** ou **ContextMenu** e define:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-308">A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:</span></span>
  
- <span data-ttu-id="c9f6c-309">Um item de menu no nível raiz.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-309">A root-level menu item.</span></span>
- <span data-ttu-id="c9f6c-310">Uma lista de itens de submenu.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-310">A list of submenu items.</span></span>

<span data-ttu-id="c9f6c-311">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-311">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon.</span></span> <span data-ttu-id="c9f6c-312">When the button is selected, the submenu displays as a drop-down list.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-312">When the button is selected, the submenu displays as a drop-down list.</span></span> <span data-ttu-id="c9f6c-313">When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-313">When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu.</span></span> <span data-ttu-id="c9f6c-314">In both cases, individual submenu items can either execute a JavaScript function or show a task pane.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-314">In both cases, individual submenu items can either execute a JavaScript function or show a task pane.</span></span> <span data-ttu-id="c9f6c-315">Only one level of submenus is supported at this time.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-315">Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="c9f6c-316">The following example shows how to define a menu item with two submenu items.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-316">The following example shows how to define a menu item with two submenu items.</span></span> <span data-ttu-id="c9f6c-317">The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-317">The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span> <span data-ttu-id="c9f6c-318">In the **Control** element:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-318">In the **Control** element:</span></span>

- <span data-ttu-id="c9f6c-319">O atributo **xsi:type** é obrigatório e deve ser definido como **Menu**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-319">The **xsi:type** attribute is required, and must be set to **Menu**.</span></span>
- <span data-ttu-id="c9f6c-320">O atributo **id** é uma cadeia de caracteres com, no máximo, 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-320">The **id** attribute is a string with a maximum of 125 characters.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|<span data-ttu-id="c9f6c-321">**Elementos**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-321">**Elements**</span></span>|<span data-ttu-id="c9f6c-322">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-322">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-323">**Label**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-323">**Label**</span></span> <br/> |<span data-ttu-id="c9f6c-324">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-324">Required.</span></span> <span data-ttu-id="c9f6c-325">The text of the root menu item.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-325">The text of the root menu item.</span></span> <span data-ttu-id="c9f6c-326">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-326">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-327">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-327">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-328">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-328">**Tooltip**</span></span> <br/> |<span data-ttu-id="c9f6c-329">Optional.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-329">Optional.</span></span> <span data-ttu-id="c9f6c-330">The tooltip for the menu.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-330">The tooltip for the menu.</span></span> <span data-ttu-id="c9f6c-331">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-331">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-332">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-332">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-333">**SuperTip**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-333">**SuperTip**</span></span> <br/> | <span data-ttu-id="c9f6c-334">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-334">Required.</span></span> <span data-ttu-id="c9f6c-335">The supertip for the menu, which is defined by the following:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-335">The supertip for the menu, which is defined by the following:</span></span> <br/> <span data-ttu-id="c9f6c-336">**Título**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-336">**Title**</span></span> <br/>  <span data-ttu-id="c9f6c-337">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-337">Required.</span></span> <span data-ttu-id="c9f6c-338">The text of the supertip.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-338">The text of the supertip.</span></span> <span data-ttu-id="c9f6c-339">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-339">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-340">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-340">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span> <br/> <span data-ttu-id="c9f6c-341">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-341">**Description**</span></span> <br/>  <span data-ttu-id="c9f6c-342">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-342">Required.</span></span> <span data-ttu-id="c9f6c-343">The description for the supertip.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-343">The description for the supertip.</span></span> <span data-ttu-id="c9f6c-344">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-344">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c9f6c-345">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-345">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span> <br/> |
|<span data-ttu-id="c9f6c-346">**Icon**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-346">**Icon**</span></span> <br/> | <span data-ttu-id="c9f6c-347">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-347">Required.</span></span> <span data-ttu-id="c9f6c-348">Contains the **Image** elements for the menu.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-348">Contains the **Image** elements for the menu.</span></span> <span data-ttu-id="c9f6c-349">Image files must be .png format.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-349">Image files must be .png format.</span></span> <br/> <span data-ttu-id="c9f6c-350">**Image**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-350">**Image**</span></span> <br/>  <span data-ttu-id="c9f6c-351">An image for the menu.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-351">An image for the menu.</span></span> <span data-ttu-id="c9f6c-352">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-352">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="c9f6c-353">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-353">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="c9f6c-354">The **size** attribute indicates the size in pixels of the image.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-354">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="c9f6c-355">Three image sizes, in pixels, are required: 16, 32, and 80.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-355">Three image sizes, in pixels, are required: 16, 32, and 80.</span></span> <span data-ttu-id="c9f6c-356">Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-356">Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64.</span></span> <br/> |
|<span data-ttu-id="c9f6c-357">**Items**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-357">**Items**</span></span> <br/> |<span data-ttu-id="c9f6c-358">Required.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-358">Required.</span></span> <span data-ttu-id="c9f6c-359">Contains the **Item** elements for each submenu item.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-359">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="c9f6c-360">Each **Item** element contains the same child elements as [Button controls](../reference/manifest/control.md#button-control).</span><span class="sxs-lookup"><span data-stu-id="c9f6c-360">Each **Item** element contains the same child elements as [Button controls](../reference/manifest/control.md#button-control).</span></span>  <br/> |

## <a name="step-7-add-the-resources-element"></a><span data-ttu-id="c9f6c-361">Etapa 7: adicionar o elemento Resources</span><span class="sxs-lookup"><span data-stu-id="c9f6c-361">Step 7: Add the Resources element</span></span>

<span data-ttu-id="c9f6c-362">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-362">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element.</span></span> <span data-ttu-id="c9f6c-363">Resources include icons, strings, and URLs.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-363">Resources include icons, strings, and URLs.</span></span> <span data-ttu-id="c9f6c-364">An element in the manifest can use a resource by referencing the **id** of the resource.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-364">An element in the manifest can use a resource by referencing the **id** of the resource.</span></span> <span data-ttu-id="c9f6c-365">Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-365">Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales.</span></span> <span data-ttu-id="c9f6c-366">An **id** has a maximum of 32 characters.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-366">An **id** has a maximum of 32 characters.</span></span>
  
<span data-ttu-id="c9f6c-367">The following shows an example of how to use the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-367">The following shows an example of how to use the **Resources** element.</span></span> <span data-ttu-id="c9f6c-368">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-368">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|<span data-ttu-id="c9f6c-369">**Recurso**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-369">**Resource**</span></span>|<span data-ttu-id="c9f6c-370">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-370">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-371">**Images**/ **Image**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-371">**Images**/ **Image**</span></span> <br/> | <span data-ttu-id="c9f6c-372">Provides the HTTPS URL to an image file.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-372">Provides the HTTPS URL to an image file.</span></span> <span data-ttu-id="c9f6c-373">Each image must define the three required image sizes:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-373">Each image must define the three required image sizes:</span></span> <br/>  <span data-ttu-id="c9f6c-374">16×16</span><span class="sxs-lookup"><span data-stu-id="c9f6c-374">16×16</span></span> <br/>  <span data-ttu-id="c9f6c-375">32×32</span><span class="sxs-lookup"><span data-stu-id="c9f6c-375">32×32</span></span> <br/>  <span data-ttu-id="c9f6c-376">80×80</span><span class="sxs-lookup"><span data-stu-id="c9f6c-376">80×80</span></span> <br/>  <span data-ttu-id="c9f6c-377">Os seguintes tamanhos de imagem também têm suporte, mas não são obrigatórios:</span><span class="sxs-lookup"><span data-stu-id="c9f6c-377">The following image sizes are also supported, but not required:</span></span> <br/>  <span data-ttu-id="c9f6c-378">20×20</span><span class="sxs-lookup"><span data-stu-id="c9f6c-378">20×20</span></span> <br/>  <span data-ttu-id="c9f6c-379">24×24</span><span class="sxs-lookup"><span data-stu-id="c9f6c-379">24×24</span></span> <br/>  <span data-ttu-id="c9f6c-380">40×40</span><span class="sxs-lookup"><span data-stu-id="c9f6c-380">40×40</span></span> <br/>  <span data-ttu-id="c9f6c-381">48×48</span><span class="sxs-lookup"><span data-stu-id="c9f6c-381">48×48</span></span> <br/>  <span data-ttu-id="c9f6c-382">64×64</span><span class="sxs-lookup"><span data-stu-id="c9f6c-382">64×64</span></span> <br/> |
|<span data-ttu-id="c9f6c-383">**Urls**/ **Url**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-383">**Urls**/ **Url**</span></span> <br/> |<span data-ttu-id="c9f6c-384">Provides an HTTPS URL location.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-384">Provides an HTTPS URL location.</span></span> <span data-ttu-id="c9f6c-385">A URL can be a maximum of 2048 characters.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-385">A URL can be a maximum of 2048 characters.</span></span>  <br/> |
|<span data-ttu-id="c9f6c-386">**ShortStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-386">**ShortStrings**/ **String**</span></span> <br/> |<span data-ttu-id="c9f6c-387">The text for **Label** and **Title** elements.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-387">The text for **Label** and **Title** elements.</span></span> <span data-ttu-id="c9f6c-388">Each **String** contains a maximum of 125 characters.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-388">Each **String** contains a maximum of 125 characters.</span></span> <br/> |
|<span data-ttu-id="c9f6c-389">**LongStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-389">**LongStrings**/ **String**</span></span> <br/> |<span data-ttu-id="c9f6c-390">The text for **Tooltip** and **Description** elements.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-390">The text for **Tooltip** and **Description** elements.</span></span> <span data-ttu-id="c9f6c-391">Each **String** contains a maximum of 250 characters.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-391">Each **String** contains a maximum of 250 characters.</span></span> <br/> |

> [!NOTE]
> <span data-ttu-id="c9f6c-392">Use o protocolo SSL (Secure Sockets Layer) para todas as URLs nos elementos **Image** e **Url**.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-392">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a><span data-ttu-id="c9f6c-393">Valores de guia para as guias da faixa de opções do aplicativo do Office padrão</span><span class="sxs-lookup"><span data-stu-id="c9f6c-393">Tab values for default Office app ribbon tabs</span></span>

<span data-ttu-id="c9f6c-394">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-394">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs.</span></span> <span data-ttu-id="c9f6c-395">The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-395">The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element.</span></span> <span data-ttu-id="c9f6c-396">The tab values are case sensitive.</span><span class="sxs-lookup"><span data-stu-id="c9f6c-396">The tab values are case sensitive.</span></span>

|<span data-ttu-id="c9f6c-397">**Aplicativo host do Office**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-397">**Office host application**</span></span>|<span data-ttu-id="c9f6c-398">**Valores de guia**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-398">**Tab values**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c9f6c-399">Excel</span><span class="sxs-lookup"><span data-stu-id="c9f6c-399">Excel</span></span>  <br/> |<span data-ttu-id="c9f6c-400">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-400">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span></span> <br/> |
|<span data-ttu-id="c9f6c-401">Word</span><span class="sxs-lookup"><span data-stu-id="c9f6c-401">Word</span></span>  <br/> |<span data-ttu-id="c9f6c-402">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-402">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span></span> <br/> |
|<span data-ttu-id="c9f6c-403">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c9f6c-403">PowerPoint</span></span>  <br/> |<span data-ttu-id="c9f6c-404">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span><span class="sxs-lookup"><span data-stu-id="c9f6c-404">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span></span>          <br/> |

## <a name="see-also"></a><span data-ttu-id="c9f6c-405">Confira também</span><span class="sxs-lookup"><span data-stu-id="c9f6c-405">See also</span></span>

- [<span data-ttu-id="c9f6c-406">Comandos de suplemento para Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c9f6c-406">Add-in commands for Excel, PowerPoint, and Word</span></span>](../design/add-in-commands.md)
