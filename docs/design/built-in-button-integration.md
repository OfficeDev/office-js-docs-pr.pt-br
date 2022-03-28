---
title: Integrar botões de Office integrados a grupos e guias de controle personalizados
description: Saiba como incluir botões de Office em seus grupos de comandos personalizados e guias na faixa de Office de opções.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 98a40b7c455cf56457595ae55f8d7d2799b270b4
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483923"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Integrar botões de Office integrados a grupos e guias de controle personalizados

Você pode inserir botões de Office em seus grupos de controle personalizados Office faixa de opções usando a marcação no manifesto do complemento. (Você não pode inserir seus comandos de complemento personalizados em um grupo de Office integrado.) Você também pode inserir grupos de controle Office inteiros em suas guias de faixa de opções personalizadas.

> [!NOTE]
> Este artigo supõe que você está familiarizado com o artigo [Conceitos básicos para comandos de complemento](add-in-commands.md). Revise-o se não tiver feito isso recentemente.

> [!IMPORTANT]
>
> - O recurso de complemento e a marcação descritos neste artigo só *estão disponíveis PowerPoint na Web*.
> - A marcação descrita neste artigo só funciona em plataformas que suportam o conjunto de **requisitos AddinCommands 1.3**. Consulte a seção Comportamento [posterior em plataformas sem suporte](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Inserir um grupo de controle integrado em uma guia personalizada

Para inserir um grupo de controle Office em uma guia, adicione um elemento [OfficeGroup](/javascript/api/manifest/customtab#officegroup) como um elemento filho no elemento **CustomTab** pai. O `id` atributo do elemento **OfficeGroup** é definido como a ID do grupo integrado. Consulte [Encontre as IDs de controles e grupos de controle](#find-the-ids-of-controls-and-control-groups).

O exemplo de marcação a seguir adiciona o grupo de controle Office Paragraph a uma guia personalizada e posiciona-o para aparecer logo após um grupo personalizado.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Inserir um controle integrado em um grupo personalizado

Para inserir um controle Office em um grupo personalizado, adicione um elemento [OfficeControl](/javascript/api/manifest/group#officecontrol) como um elemento filho no elemento **Group** pai. O `id` atributo do **elemento OfficeControl** é definido como a ID do controle integrado. Consulte [Encontre as IDs de controles e grupos de controle](#find-the-ids-of-controls-and-control-groups).

O exemplo de marcação a seguir adiciona o Office sobrescrito a um grupo personalizado e posiciona-o para aparecer logo após um botão personalizado.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button1">
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
> Os usuários podem personalizar a faixa de opções no Office aplicativo. Qualquer personalização do usuário substituirá suas configurações de manifesto. Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Encontre as IDs de controles e grupos de controle

As IDs para controles e grupos de controle com suporte estão em arquivos no repo [Office IDs de controle](https://github.com/OfficeDev/office-control-ids). Siga as instruções no arquivo ReadMe desse repo.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento em plataformas sem suporte

Se o seu add-in estiver instalado em uma plataforma que não oferece suporte ao conjunto de [requisitos AddinCommands 1.3](/javascript/api/requirement-sets/add-in-commands-requirement-sets), a marcação descrita neste artigo será ignorada e os controles/grupos de Office internos não aparecerão em seus grupos/guias personalizados. Para impedir que o seu complemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na seção Requisitos  do manifesto. Para obter instruções, [consulte Especificar quais Office e plataformas podem hospedar seu complemento](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Como alternativa, projete seu complemento para ter uma experiência quando **AddinCommands 1.3** não tiver suporte, conforme descrito em [Design para experiências alternativas](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Por exemplo, se o seu add-in contiver instruções que pressuem que os botões integrados estão em seus grupos personalizados, você pode projetar uma versão que presume que os botões integrados estão apenas em seus locais usuais.
