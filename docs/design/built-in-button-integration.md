---
title: Integrar botões integrados do Office a grupos de controle personalizados e guias
description: Saiba como incluir botões do Office integrados em seus grupos de comandos personalizados e guias na faixa de opções do Office.
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505252"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a><span data-ttu-id="3c15d-103">Integrar botões integrados do Office a grupos de controle personalizados e guias</span><span class="sxs-lookup"><span data-stu-id="3c15d-103">Integrate built-in Office buttons into custom control groups and tabs</span></span>

<span data-ttu-id="3c15d-104">Você pode inserir botões do Office integrados em seus grupos de controle personalizados na faixa de opções do Office usando marcação no manifesto do complemento.</span><span class="sxs-lookup"><span data-stu-id="3c15d-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="3c15d-105">(Você não pode inserir seus comandos de complemento personalizados em um grupo do Office integrado.) Você também pode inserir grupos de controle do Office inteiros em suas guias de faixa de opções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="3c15d-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="3c15d-106">Este artigo supõe que você está familiarizado com o artigo [Conceitos básicos para comandos de complemento.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="3c15d-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="3c15d-107">Revise-o se não tiver feito isso recentemente.</span><span class="sxs-lookup"><span data-stu-id="3c15d-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="3c15d-108">O recurso de add-in e a marcação descritos neste artigo só estão disponíveis no *PowerPoint na Web*.</span><span class="sxs-lookup"><span data-stu-id="3c15d-108">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="3c15d-109">A marcação descrita neste artigo só funciona em plataformas que suportam o conjunto de **requisitos AddinCommands 1.3**.</span><span class="sxs-lookup"><span data-stu-id="3c15d-109">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="3c15d-110">Consulte a seção Comportamento [posterior em plataformas sem suporte.](#behavior-on-unsupported-platforms)</span><span class="sxs-lookup"><span data-stu-id="3c15d-110">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="3c15d-111">Inserir um grupo de controle integrado em uma guia personalizada</span><span class="sxs-lookup"><span data-stu-id="3c15d-111">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="3c15d-112">Para inserir um grupo de controle do Office integrado em uma guia, adicione um elemento [OfficeGroup](../reference/manifest/customtab.md#officegroup) como um elemento filho no elemento `<CustomTab>` pai.</span><span class="sxs-lookup"><span data-stu-id="3c15d-112">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="3c15d-113">O `id` atributo do elemento é definido como a `<OfficeGroup>` ID do grupo integrado.</span><span class="sxs-lookup"><span data-stu-id="3c15d-113">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="3c15d-114">Consulte [Encontre as IDs de controles e grupos de controle.](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="3c15d-114">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="3c15d-115">O exemplo de marcação a seguir adiciona o grupo de controle Parágrafo do Office a uma guia personalizada e posiciona-o para aparecer logo após um grupo personalizado.</span><span class="sxs-lookup"><span data-stu-id="3c15d-115">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="3c15d-116">Inserir um controle integrado em um grupo personalizado</span><span class="sxs-lookup"><span data-stu-id="3c15d-116">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="3c15d-117">Para inserir um controle do Office integrado em um grupo personalizado, adicione um elemento [OfficeControl](../reference/manifest/group.md#officecontrol) como um elemento filho no elemento `<Group>` pai.</span><span class="sxs-lookup"><span data-stu-id="3c15d-117">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="3c15d-118">O `id` atributo do elemento é definido como a `<OfficeControl>` ID do controle integrado.</span><span class="sxs-lookup"><span data-stu-id="3c15d-118">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="3c15d-119">Consulte [Encontre as IDs de controles e grupos de controle.](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="3c15d-119">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="3c15d-120">O exemplo de marcação a seguir adiciona o controle Sobrescrito do Office a um grupo personalizado e posiciona-o para aparecer logo após um botão personalizado.</span><span class="sxs-lookup"><span data-stu-id="3c15d-120">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

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
> <span data-ttu-id="3c15d-121">Os usuários podem personalizar a faixa de opções no aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="3c15d-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="3c15d-122">Qualquer personalização do usuário substituirá suas configurações de manifesto.</span><span class="sxs-lookup"><span data-stu-id="3c15d-122">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="3c15d-123">Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.</span><span class="sxs-lookup"><span data-stu-id="3c15d-123">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="3c15d-124">Encontre as IDs de controles e grupos de controle</span><span class="sxs-lookup"><span data-stu-id="3c15d-124">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="3c15d-125">As IDs para controles suportados e grupos de controle estão em arquivos nas [IDs](https://github.com/OfficeDev/office-control-ids)de controle do Office de repo .</span><span class="sxs-lookup"><span data-stu-id="3c15d-125">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="3c15d-126">Siga as instruções no arquivo ReadMe desse repo.</span><span class="sxs-lookup"><span data-stu-id="3c15d-126">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="3c15d-127">Comportamento em plataformas sem suporte</span><span class="sxs-lookup"><span data-stu-id="3c15d-127">Behavior on unsupported platforms</span></span>

<span data-ttu-id="3c15d-128">Se o seu add-in estiver instalado em uma plataforma que não oferece suporte ao conjunto de [requisitos AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e os controles/grupos internos do Office não aparecerão em seus grupos/guias personalizados.</span><span class="sxs-lookup"><span data-stu-id="3c15d-128">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="3c15d-129">Para impedir que o seu complemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na seção `<Requirements>` do manifesto.</span><span class="sxs-lookup"><span data-stu-id="3c15d-129">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="3c15d-130">Para obter instruções, consulte [Definir o elemento Requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="3c15d-130">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="3c15d-131">Como alternativa, você pode projetar seu complemento para ter uma experiência alternativa quando **AddinCommands 1.3** não tiver suporte, conforme descrito em Usar verificações de tempo de execução em seu código [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="3c15d-131">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="3c15d-132">Por exemplo, se o seu add-in contiver instruções que pressuem que os botões integrados estão em seus grupos personalizados, você pode ter uma versão alternativa que presume que os botões integrados estão apenas em seus locais usuais.</span><span class="sxs-lookup"><span data-stu-id="3c15d-132">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
