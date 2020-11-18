---
title: Integrar botões internos do Office em guias e grupos de controles personalizados
description: Saiba como incluir botões internos do Office em seus grupos de comandos personalizados e guias na faixa de opções do Office.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: e04107893b3c0dd453c84d38fdd5623e308b70e3
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088162"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a><span data-ttu-id="f84ca-103">Integrar botões internos do Office em guias e grupos de controles personalizados (visualização)</span><span class="sxs-lookup"><span data-stu-id="f84ca-103">Integrate built-in Office buttons into custom control groups and tabs (preview)</span></span>

<span data-ttu-id="f84ca-104">Você pode inserir botões internos do Office em seus grupos de controle personalizados na faixa de opções do Office usando a marcação no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f84ca-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="f84ca-105">(Você não pode inserir seus comandos de suplemento personalizados em um grupo interno do Office.) Você também pode inserir todos os grupos de controle internos do Office em suas guias da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="f84ca-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="f84ca-106">Este artigo pressupõe que você esteja familiarizado com os [conceitos básicos do artigo para comandos de suplemento](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="f84ca-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="f84ca-107">Verifique se você ainda não fez isso recentemente.</span><span class="sxs-lookup"><span data-stu-id="f84ca-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="f84ca-108">O recurso de suplemento e a marcação descritos neste artigo estão em visualização e só estão *disponíveis no PowerPoint na Web*.</span><span class="sxs-lookup"><span data-stu-id="f84ca-108">The add-in feature and markup described in this article is in preview and is *only available in PowerPoint on the web*.</span></span> <span data-ttu-id="f84ca-109">Recomendamos que você experimente a marcação apenas em ambientes de teste e desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="f84ca-109">We recommend that you try out the markup in test and development environments only.</span></span> <span data-ttu-id="f84ca-110">Não use a marcação de visualização em um ambiente de produção ou em documentos de negócios críticos.</span><span class="sxs-lookup"><span data-stu-id="f84ca-110">Do not use preview markup in a production environment or within business-critical documents.</span></span>
> - <span data-ttu-id="f84ca-111">A marcação descrita neste artigo funciona somente em plataformas que dão suporte ao conjunto de requisitos **AddinCommands 1,3**.</span><span class="sxs-lookup"><span data-stu-id="f84ca-111">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="f84ca-112">Consulte o comportamento da seção posterior [em plataformas sem suporte](#behavior-on-unsupported-platforms).</span><span class="sxs-lookup"><span data-stu-id="f84ca-112">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="f84ca-113">Inserir um grupo de controle interno em uma guia personalizada</span><span class="sxs-lookup"><span data-stu-id="f84ca-113">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="f84ca-114">Para inserir um grupo de controle interno do Office em uma guia, adicione um elemento de grupo [do Office](../reference/manifest/customtab.md#officegroup) como um elemento filho no `<CustomTab>` elemento pai.</span><span class="sxs-lookup"><span data-stu-id="f84ca-114">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="f84ca-115">O `id` atributo do do `<OfficeGroup>` elemento é definido como a ID do grupo interno.</span><span class="sxs-lookup"><span data-stu-id="f84ca-115">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="f84ca-116">Consulte [localizar as IDs de controles e grupos de controle](#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="f84ca-116">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="f84ca-117">O exemplo a seguir de marcação adiciona o grupo de controle de parágrafo do Office a uma guia personalizada e o posiciona para aparecer imediatamente após um grupo personalizado.</span><span class="sxs-lookup"><span data-stu-id="f84ca-117">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="f84ca-118">Inserir um controle interno em um grupo personalizado</span><span class="sxs-lookup"><span data-stu-id="f84ca-118">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="f84ca-119">Para inserir um controle interno do Office em um grupo personalizado, adicione um elemento [OfficeControl](../reference/manifest/group.md#officecontrol) como um elemento filho no `<Group>` elemento pai.</span><span class="sxs-lookup"><span data-stu-id="f84ca-119">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="f84ca-120">O `id` atributo do `<OfficeControl>` elemento é definido como a ID do controle interno.</span><span class="sxs-lookup"><span data-stu-id="f84ca-120">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="f84ca-121">Consulte [localizar as IDs de controles e grupos de controle](#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="f84ca-121">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="f84ca-122">O exemplo a seguir de marcação adiciona o controle de Superscript do Office a um grupo personalizado e o posiciona para aparecer imediatamente após um botão personalizado.</span><span class="sxs-lookup"><span data-stu-id="f84ca-122">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> <span data-ttu-id="f84ca-123">Os usuários podem personalizar a faixa de opções no aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="f84ca-123">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="f84ca-124">Qualquer personalização do usuário substituirá as configurações do manifesto.</span><span class="sxs-lookup"><span data-stu-id="f84ca-124">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="f84ca-125">Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.</span><span class="sxs-lookup"><span data-stu-id="f84ca-125">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="f84ca-126">Localizar as IDs de controles e grupos de controle</span><span class="sxs-lookup"><span data-stu-id="f84ca-126">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="f84ca-127">As IDs para controles e grupos de controle suportados estão em arquivos nas [IDs de controle do Office](https://github.com/OfficeDev/office-control-ids)do repositório.</span><span class="sxs-lookup"><span data-stu-id="f84ca-127">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="f84ca-128">Siga as instruções no arquivo Leiame desse repositório.</span><span class="sxs-lookup"><span data-stu-id="f84ca-128">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="f84ca-129">Comportamento de plataformas sem suporte</span><span class="sxs-lookup"><span data-stu-id="f84ca-129">Behavior on unsupported platforms</span></span>

<span data-ttu-id="f84ca-130">Se seu suplemento estiver instalado em uma plataforma que não ofereça suporte ao [conjunto de requisitos AddinCommands 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e os controles/grupos internos do Office não aparecerão em seus grupos/guias personalizados.</span><span class="sxs-lookup"><span data-stu-id="f84ca-130">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="f84ca-131">Para impedir que o suplemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na `<Requirements>` seção do manifesto.</span><span class="sxs-lookup"><span data-stu-id="f84ca-131">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="f84ca-132">Para obter instruções, consulte [definir o elemento requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="f84ca-132">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="f84ca-133">Como alternativa, você pode criar seu suplemento para ter uma experiência alternativa quando o **AddinCommands 1,3** não é suportado, conforme descrito em [usar verificações de tempo de execução no código JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="f84ca-133">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="f84ca-134">Por exemplo, se o suplemento contiver instruções que presumim que os botões internos estão em seus grupos personalizados, você poderia ter uma versão alternativa que supõe que os botões internos estão apenas nos seus lugares usuais.</span><span class="sxs-lookup"><span data-stu-id="f84ca-134">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
