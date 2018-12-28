---
title: Elemento OfficeTab no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 721064687c3c892b565a94e418815726cc0817f5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432869"
---
# <a name="officetab-element"></a>Elemento OfficeTab

Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Pode ser a guia padrão (**Início**, **Mensagem** ou **Reunião**) ou uma guia personalizada definida pelo suplemento. Este elemento é obrigatório.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  Grupo      | Sim |  Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.  |

A seguir estão os valores válidos de `id` por host. Os valores em **negrito** têm suporte na área de trabalho e online (por exemplo, Word 2016 ou posterior para Windows e Word Online).

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

Um grupo de pontos de extensão de interface do usuário em uma guia. O grupo pode ter até seis controles. O atributo **id** é obrigatório, e cada **id** deve ser exclusiva dentro do manifesto. A **id** é uma cadeia de caracteres com, no máximo, 125 caracteres. Confira [Elemento Group](group.md)

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
