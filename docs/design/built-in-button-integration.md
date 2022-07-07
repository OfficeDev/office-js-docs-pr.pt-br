---
title: Integrar botões internos do Office em guias e grupos de controle personalizados
description: Saiba como incluir botões internos do Office em seus grupos de comandos personalizados e guias na faixa de opções do Office.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc706fcd0b049647847a73f7c40144dba9df0e2
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659784"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Integrar botões internos do Office em guias e grupos de controle personalizados

Você pode inserir botões internos do Office em seus grupos de controle personalizados na faixa de opções do Office usando marcação no manifesto do suplemento. (Você não pode inserir seus comandos de suplemento personalizados em um grupo interno do Office.) Você também pode inserir grupos de controles internos do Office inteiros nas guias da faixa de opções personalizadas.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com o artigo [Conceitos básicos para comandos de suplemento](add-in-commands.md). Examine-o se você ainda não fez isso recentemente.

> [!IMPORTANT]
>
> - O recurso de suplemento e a marcação descritos neste artigo só *estão disponíveis PowerPoint na Web*.
> - A marcação descrita neste artigo só funciona em plataformas que dão suporte ao conjunto de requisitos **AddinCommands 1.3**. Consulte a seção Posterior [Comportamento em plataformas sem suporte](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Inserir um grupo de controle interno em uma guia personalizada

Para inserir um grupo de controle interno do Office em uma guia, adicione um elemento [OfficeGroup](/javascript/api/manifest/customtab#officegroup) como um elemento filho no elemento **\<CustomTab\>** pai. O `id` atributo do elemento é **\<OfficeGroup\>** definido como a ID do grupo interno. Consulte [Localizar as IDs de controles e grupos de controles](#find-the-ids-of-controls-and-control-groups).

O exemplo de marcação a seguir adiciona o grupo de controle Parágrafo do Office a uma guia personalizada e o posiciona para aparecer logo após um grupo personalizado.

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Inserir um controle interno em um grupo personalizado

Para inserir um controle interno do Office em um grupo personalizado, adicione um elemento [OfficeControl](/javascript/api/manifest/group#officecontrol) como um elemento filho no elemento **\<Group\>** pai. O `id` atributo do **\<OfficeControl\>** elemento é definido como a ID do controle interno. Consulte [Localizar as IDs de controles e grupos de controles](#find-the-ids-of-controls-and-control-groups).

O exemplo de marcação a seguir adiciona o controle Sobrescrito do Office a um grupo personalizado e o posiciona para aparecer logo após um botão personalizado.

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
> Os usuários podem personalizar a faixa de opções no aplicativo do Office. As personalizações do usuário substituirão as configurações de manifesto. Por exemplo, um usuário pode remover um botão de qualquer grupo e remover qualquer grupo de uma guia.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Localizar as IDs de controles e grupos de controles

As IDs para controles com suporte e grupos de controle estão em arquivos nas [IDs de controle do Office do repositório](https://github.com/OfficeDev/office-control-ids). Siga as instruções no arquivo LeiaMe desse repositório.

## <a name="behavior-on-unsupported-platforms"></a>Comportamento em plataformas sem suporte

Se o suplemento estiver instalado em uma plataforma que não dá suporte ao conjunto de [requisitos AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), a marcação descrita neste artigo será ignorada e os controles/grupos internos do Office não aparecerão em seus grupos/guias personalizados. Para impedir que o suplemento seja instalado em plataformas que não dão suporte à marcação, **\<Requirements\>** adicione uma referência ao conjunto de requisitos na seção do manifesto. Para obter instruções, [consulte Especificar quais versões e plataformas do Office podem hospedar seu suplemento](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Como alternativa, projete seu suplemento para ter uma experiência quando **o AddinCommands 1.3** não tiver suporte, conforme descrito em Design para experiências [alternativas](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Por exemplo, se o suplemento contiver instruções que pressupõem que os botões internos estejam em seus grupos personalizados, você poderá criar uma versão que pressuponha que os botões internos estejam apenas em seus locais comuns.
