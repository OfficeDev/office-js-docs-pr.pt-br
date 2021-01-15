---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771323"
---
# <a name="customtab-element"></a>Elemento CustomTab

Na faixa de opções, especifique a guia e o grupo para os comandos de suplemento. Isso pode ser a guia padrão ( **página inicial**, de **mensagem** ou **reunião**) ou em uma guia personalizada definida pelo suplemento.

Nas guias personalizadas, o suplemento pode ter grupos internos ou personalizados. Os suplementos estão limitados a uma guia personalizada.

O atributo **ID** deve ser exclusivo dentro do manifesto.

> [!IMPORTANT]
> No Outlook no Mac, o `CustomTab` elemento não está disponível, portanto, você terá que usar o [OfficeTab](officetab.md) .

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Não |  Define um grupo de comandos  |
|  [O Microsoft Office](#officegroup)      | Não |  Representa um grupo de controle interno do Office.  |
|  [Label](#label-tab)      | Sim |  O rótulo para CustomTab ou Group.  |
|  [InsertAfter](#insertafter)      | Não |  Especifica que a guia personalizada deve ser imediatamente após uma guia interna especificada do Office.  |
|  [InsertBefore](#insertbefore)      | Não |  Especifica que a guia personalizada deve ser imediatamente anterior à guia interna especificada do Office.  |

### <a name="group"></a>Grupo

Opcional, mas, se não estiver presente, deve haver pelo **menos um elemento** de um. Confira [Elemento Group](group.md) A ordem do **grupo** e do grupo do **Office** no manifesto deve ser a ordem que você deseja que eles apareçam na guia Personalizar. Eles podem ser mesclados se houver vários elementos, mas todos devem estar acima do elemento **rótulo** .

### <a name="officegroup"></a>O Microsoft Office

Opcional, mas se não houver, deve haver pelo menos um elemento de **grupo** . Representa um grupo de controle interno do Office. O atributo **ID** especifica a ID do grupo interno do Office. Para localizar a ID de um grupo interno, confira [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). A ordem do **grupo** e do grupo do **Office** no manifesto deve ser a ordem que você deseja que eles apareçam na guia Personalizar. Eles podem ser mesclados se houver vários elementos, mas todos devem estar acima do elemento **rótulo** .

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizado. O atributo **Resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .

### <a name="insertafter"></a>InsertAfter

Opcional. Especifica que a guia personalizada deve ser imediatamente após uma guia interna especificada do Office. O valor do elemento é a ID da guia interna, como "TabHome" ou "TabReview". (Consulte [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) Se presente, deve ser após o elemento **Label** . Você não pode ter **InsertAfter** e **InsertBefore**.

### <a name="insertbefore"></a>InsertBefore

Opcional. Especifica que a guia personalizada deve ser imediatamente anterior à guia interna especificada do Office. O valor do elemento é a ID da guia interna, como "TabHome" ou "TabReview". (Consulte [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  Se presente, deve ser após o elemento **Label** . Você não pode ter **InsertAfter** e **InsertBefore**.

## <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
