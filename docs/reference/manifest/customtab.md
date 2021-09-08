---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 09/02/2021
localization_priority: Normal
ms.openlocfilehash: 642b6eabaa9885041dd122b179ee2baa3e772977
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939080"
---
# <a name="customtab-element"></a>Elemento CustomTab

Na faixa de opções, especifique a guia e o grupo para os comandos do seu complemento. Isso pode estar na guia padrão **(Home,** **Message** ou **Meeting**) ou em uma guia personalizada definida pelo add-in.

Em guias personalizadas, o complemento pode ter grupos personalizados ou integrados. Os suplementos estão limitados a uma guia personalizada.

O **atributo id** deve ser exclusivo no manifesto.

> [!IMPORTANT]
> No Outlook no Mac, o elemento não está disponível, portanto, você `CustomTab` terá que usar o [OfficeTab](officetab.md) em vez disso.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Não |  Define um grupo de comandos  |
|  [OfficeGroup](#officegroup)      | Não |  Representa um grupo de controle Office integrado. **Importante**: não disponível no Outlook. |
|  [Label](#label-tab)      | Sim |  O rótulo para CustomTab ou Group.  |
|  [InsertAfter](#insertafter)      | Não |  Especifica que a guia personalizada deve ser imediatamente após uma guia Office de Office especificada. **Importante**: disponível somente no PowerPoint. |
|  [InsertBefore](#insertbefore)      | Não |  Especifica que a guia personalizada deve ser imediatamente antes de uma guia de Office de Office. **Importante**: disponível somente no PowerPoint. |

### <a name="group"></a>Grupo

Opcional, mas se não estiver presente, deve haver pelo menos um **elemento OfficeGroup.** Confira [Elemento Group](group.md) A ordem de **Group** e **OfficeGroup** no manifesto deve ser a ordem que você deseja que eles apareçam na guia personalizada. Eles podem ser intermendados se houver vários elementos, mas todos devem estar acima do **elemento Label.**

### <a name="officegroup"></a>OfficeGroup

Opcional, mas se não estiver presente, deve haver pelo menos um **elemento Group.** Representa um grupo de controle Office integrado. O **atributo id** especifica a ID do grupo de Office integrado. Para encontrar a ID de um grupo integrado, consulte [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). A ordem de **Group** e **OfficeGroup** no manifesto deve ser a ordem que você deseja que eles apareçam na guia personalizada. Eles podem ser intermendados se houver vários elementos, mas todos devem estar acima do **elemento Label.**

> [!IMPORTANT]
> O `OfficeGroup` elemento não está disponível no Outlook.

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizada. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources.](resources.md)

### <a name="insertafter"></a>InsertAfter

Opcional. Especifica que a guia personalizada deve ser imediatamente após uma guia Office. O valor do elemento é a ID da guia integrado, como "TabHome" ou "TabReview". (Consulte [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) Se presente, deve estar após o **elemento Label.** Não é possível ter **InsertAfter e** **InsertBefore**.

> [!IMPORTANT]
> O `InsertAfter` elemento só está disponível em PowerPoint.

### <a name="insertbefore"></a>InsertBefore

Opcional. Especifica que a guia personalizada deve ser imediatamente antes de uma guia Office. O valor do elemento é a ID da guia integrado, como "TabHome" ou "TabReview". (Consulte [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  Se presente, deve estar após o **elemento Label.** Não é possível ter **InsertAfter e** **InsertBefore**.

> [!IMPORTANT]
> O `InsertBefore` elemento só está disponível em PowerPoint.
