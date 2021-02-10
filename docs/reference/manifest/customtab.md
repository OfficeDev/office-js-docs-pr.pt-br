---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173924"
---
# <a name="customtab-element"></a>Elemento CustomTab

Na faixa de opções, especifique a guia e o grupo para os comandos do seu complemento. Isso pode estar na guia padrão (Página **Início,** Mensagem ou **Reunião)** ou em uma guia personalizada definida pelo complemento.

Em guias personalizadas, o complemento pode ter grupos personalizados ou integrados. Os suplementos estão limitados a uma guia personalizada.

O **atributo id** deve ser exclusivo dentro do manifesto.

> [!IMPORTANT]
> No Outlook no Mac, `CustomTab` o elemento não está disponível, portanto, você terá que usar o [OfficeTab.](officetab.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Não |  Define um grupo de comandos  |
|  [OfficeGroup](#officegroup)      | Não |  Representa um grupo de controles integrado do Office. **Importante:** não disponível no Outlook. |
|  [Label](#label-tab)      | Sim |  O rótulo para CustomTab ou Group.  |
|  [InsertAfter](#insertafter)      | Não |  Especifica que a guia personalizada deve estar imediatamente após uma guia do Office. **Importante:** não disponível no Outlook. |
|  [InsertBefore](#insertbefore)      | Não |  Especifica que a guia personalizada deve estar imediatamente antes de uma guia do Office. **Importante:** não disponível no Outlook. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Não |  Especifica se a guia personalizada deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas. **Importante:** não disponível no Outlook. |

### <a name="group"></a>Grupo

Opcional, mas se não estiver presente, deve haver pelo menos um **elemento OfficeGroup.** Confira [Elemento Group](group.md) A ordem do **Grupo** e **do OfficeGroup** no manifesto deve ser a ordem em que você deseja que apareçam na guia personalizada. Eles podem ser intercalados se houver vários elementos, mas todos devem estar acima do **elemento Label.**

### <a name="officegroup"></a>OfficeGroup

Opcional, mas se não estiver presente, deve haver pelo menos um **elemento Group.** Representa um grupo de controles integrado do Office. O **atributo id** especifica a ID do grupo do Office integrado. Para encontrar a ID de um grupo integrado, consulte [Encontrar as IDs de controles e grupos de controles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) A ordem do **Grupo** e **do OfficeGroup** no manifesto deve ser a ordem em que você deseja que apareçam na guia personalizada. Eles podem ser intercalados se houver vários elementos, mas todos devem estar acima do **elemento Label.**

> [!IMPORTANT]
> O `OfficeGroup` elemento não está disponível no Outlook.

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizada. O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)

### <a name="insertafter"></a>InsertAfter

Opcional. Especifica que a guia personalizada deve ser imediatamente após uma guia do Office. O valor do elemento é a ID da guia integrado, como "TabHome" ou "TabReview". (Consulte [Encontrar as IDs de controles e grupos de controles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Se presente, deve ser após o **elemento Label.** You cannot have both **InsertAfter** and **InsertBefore**.

> [!IMPORTANT]
> O `InsertAfter` elemento não está disponível no Outlook.

### <a name="insertbefore"></a>InsertBefore

Opcional. Especifica que a guia personalizada deve estar imediatamente antes de uma guia do Office. O valor do elemento é a ID da guia integrado, como "TabHome" ou "TabReview". (Consulte [Encontrar as IDs de controles e grupos de controles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Se presente, deve ser após o **elemento Label.** You cannot have both **InsertAfter** and **InsertBefore**.

> [!IMPORTANT]
> O `InsertBefore` elemento não está disponível no Outlook.

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Opcional (booliana). Especifica se a **CustomTab** ficará oculta em combinações de aplicativos e plataformas que suportam uma API que instala uma guia contextual personalizada na faixa de opções em tempo de execução. O valor padrão, se não estiver presente, é `false` . Se usado, **OverriddenByRibbonApi** deve ser o *primeiro* filho de **CustomTab**. Para obter mais informações, [consulte OverriddenByRibbonApi](overriddenbyribbonapi.md).

> [!IMPORTANT]
> O `OverriddenByRibbonApi` elemento não está disponível no Outlook.

## <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
