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
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a>Integrar botões internos do Office em guias e grupos de controles personalizados (visualização)

Você pode inserir botões internos do Office em seus grupos de controle personalizados na faixa de opções do Office usando a marcação no manifesto do suplemento. (Você não pode inserir seus comandos de suplemento personalizados em um grupo interno do Office.) Você também pode inserir todos os grupos de controle internos do Office em suas guias da faixa de opções personalizada.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com os [conceitos básicos do artigo para comandos de suplemento](add-in-commands.md). Verifique se você ainda não fez isso recentemente.

> [!IMPORTANT]
>
> - O recurso de suplemento e a marcação descritos neste artigo estão em visualização e só estão *disponíveis no PowerPoint na Web*. Recomendamos que você experimente a marcação apenas em ambientes de teste e desenvolvimento. Não use a marcação de visualização em um ambiente de produção ou em documentos de negócios críticos.
> - A marcação descrita neste artigo funciona somente em plataformas que dão suporte ao conjunto de requisitos **AddinCommands 1,3**. Consulte o comportamento da seção posterior [em plataformas sem suporte](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Inserir um grupo de controle interno em uma guia personalizada

Para inserir um grupo de controle interno do Office em uma guia, adicione um elemento de grupo [do Office](../reference/manifest/customtab.md#officegroup) como um elemento filho no `<CustomTab>` elemento pai. O `id` atributo do do `<OfficeGroup>` elemento é definido como a ID do grupo interno. Consulte [localizar as IDs de controles e grupos de controle](#find-the-ids-of-controls-and-control-groups).

O exemplo a seguir de marcação adiciona o grupo de controle de parágrafo do Office a uma guia personalizada e o posiciona para aparecer imediatamente após um grupo personalizado.

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Inserir um controle interno em um grupo personalizado

Para inserir um controle interno do Office em um grupo personalizado, adicione um elemento [OfficeControl](../reference/manifest/group.md#officecontrol) como um elemento filho no `<Group>` elemento pai. O `id` atributo do `<OfficeControl>` elemento é definido como a ID do controle interno. Consulte [localizar as IDs de controles e grupos de controle](#find-the-ids-of-controls-and-control-groups).

O exemplo a seguir de marcação adiciona o controle de Superscript do Office a um grupo personalizado e o posiciona para aparecer imediatamente após um botão personalizado.

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
> Os usuários podem personalizar a faixa de opções no aplicativo do Office. Qualquer personalização do usuário substituirá as configurações do manifesto. Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Localizar as IDs de controles e grupos de controle

As IDs para controles e grupos de controle suportados estão em arquivos nas [IDs de controle do Office](https://github.com/OfficeDev/office-control-ids)do repositório. Siga as instruções no arquivo Leiame desse repositório.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento de plataformas sem suporte

Se seu suplemento estiver instalado em uma plataforma que não ofereça suporte ao [conjunto de requisitos AddinCommands 1,3](../reference/requirement-sets/add-in-commands-requirement-sets.md), a marcação descrita neste artigo será ignorada e os controles/grupos internos do Office não aparecerão em seus grupos/guias personalizados. Para impedir que o suplemento seja instalado em plataformas que não suportam a marcação, adicione uma referência ao conjunto de requisitos na `<Requirements>` seção do manifesto. Para obter instruções, consulte [definir o elemento requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest). Como alternativa, você pode criar seu suplemento para ter uma experiência alternativa quando o **AddinCommands 1,3** não é suportado, conforme descrito em [usar verificações de tempo de execução no código JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). Por exemplo, se o suplemento contiver instruções que presumim que os botões internos estão em seus grupos personalizados, você poderia ter uma versão alternativa que supõe que os botões internos estão apenas nos seus lugares usuais.
