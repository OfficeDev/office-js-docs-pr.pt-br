---
title: Elemento OfficeTab no arquivo de manifesto
description: O elemento OfficeTab define a guia faixa de opções onde o comando do seu complemento é exibido.
ms.date: 06/20/2019
ms.localizationpriority: medium
ms.openlocfilehash: 94af0e8744c496538a506fdc4626cd3f515129bb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151787"
---
# <a name="officetab-element"></a>Elemento OfficeTab

Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Essa pode ser a guia padrão **(Home**, **Message** ou **Meeting**) ou uma guia personalizada definida pelo complemento. Este elemento é obrigatório.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  Group      | Sim |  Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.  |

A seguir estão os valores de `id` tabulação válidos por aplicativo. Os valores **em negrito** são suportados na área de trabalho e online (por exemplo, Word 2016 ou posterior em Windows e Word na Web).

### <a name="outlook"></a>Outlook

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>Group

Um grupo de pontos de extensão da interface do usuário em uma guia. Um grupo pode ter até seis controles. O **atributo id** é necessário e cada **id** deve ser exclusivo no manifesto. A **id é** uma cadeia de caracteres com no máximo 125 caracteres. Confira [Elemento Group](group.md)

## <a name="officetab-example"></a>Exemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
