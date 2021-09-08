---
title: Integrar botões de Office integrados a grupos e guias de controle personalizados
description: Saiba como incluir botões de Office em seus grupos de comandos personalizados e guias na faixa de Office de opções.
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937833"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Integrar botões de Office integrados a grupos e guias de controle personalizados

Você pode inserir botões de Office em seus grupos de controle personalizados na faixa Office faixa de opções usando marcação no manifesto do complemento. (Você não pode inserir seus comandos de complemento personalizados em um grupo de Office integrado.) Você também pode inserir grupos de controle Office inteiros em suas guias de faixa de opções personalizadas.

> [!NOTE]
> Este artigo supõe que você está familiarizado com o artigo [Conceitos básicos para comandos de complemento.](add-in-commands.md) Revise-o se não tiver feito isso recentemente.

> [!IMPORTANT]
>
> - O recurso de complemento e a marcação descritos neste artigo só estão disponíveis *em PowerPoint na Web*.
> - A marcação descrita neste artigo só funciona em plataformas que suportam o conjunto de **requisitos AddinCommands 1.3**. Consulte a seção Comportamento [posterior em plataformas sem suporte.](#behavior-on-unsupported-platforms)

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Inserir um grupo de controle integrado em uma guia personalizada

Para inserir um grupo de controle Office em uma guia, adicione um elemento [OfficeGroup](../reference/manifest/customtab.md#officegroup) como um elemento filho no elemento `<CustomTab>` pai. O `id` atributo do elemento é definido como a `<OfficeGroup>` ID do grupo integrado. Consulte [Encontre as IDs de controles e grupos de controle.](#find-the-ids-of-controls-and-control-groups)

O exemplo de marcação a seguir adiciona o grupo Office controle Paragraph a uma guia personalizada e posiciona-a para aparecer logo após um grupo personalizado.

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Inserir um controle integrado em um grupo personalizado

Para inserir um controle de Office em um grupo personalizado, adicione um elemento [OfficeControl](../reference/manifest/group.md#officecontrol) como um elemento filho no elemento `<Group>` pai. O `id` atributo do elemento é definido como a `<OfficeControl>` ID do controle integrado. Consulte [Encontre as IDs de controles e grupos de controle.](#find-the-ids-of-controls-and-control-groups)

O exemplo de marcação a seguir adiciona o Office sobrescrito a um grupo personalizado e posiciona-o para aparecer logo após um botão personalizado.

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
> Os usuários podem personalizar a faixa de opções no Office aplicativo. Qualquer personalização do usuário substituirá suas configurações de manifesto. Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Encontre as IDs de controles e grupos de controle

As IDs para controles suportados e grupos de controle estão em arquivos no repo [Office IDs de Controle.](https://github.com/OfficeDev/office-control-ids) Siga as instruções no arquivo ReadMe desse repo.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento em plataformas sem suporte

Se o seu add-in estiver instalado em uma plataforma que não oferece suporte ao conjunto de [requisitos AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e os controles/grupos de Office internos não aparecerão em seus grupos/guias personalizados. Para impedir que o seu complemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na seção `<Requirements>` do manifesto. Para obter instruções, consulte [Definir o elemento Requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Como alternativa, você pode projetar seu complemento para ter uma experiência alternativa quando **AddinCommands 1.3** não tiver suporte, conforme descrito em Usar verificações de tempo de execução em seu código [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). Por exemplo, se o seu add-in contiver instruções que pressuem que os botões integrados estão em seus grupos personalizados, você pode ter uma versão alternativa que presume que os botões integrados estão apenas em seus locais usuais.
