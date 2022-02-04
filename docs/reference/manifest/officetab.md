---
title: Elemento OfficeTab no arquivo de manifesto
description: O elemento OfficeTab define a guia faixa de opções onde o comando do seu complemento é exibido.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="officetab-element"></a>Elemento OfficeTab

Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Pode ser a guia padrão ( **Home**, **Message** ou **Meeting**) ou uma guia personalizada definida pelo add-in. Este elemento é obrigatório.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  Group      | Sim |  Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.  |

A seguir estão os valores de tabulação `id` válidos por aplicativo. Os valores **em negrito** são suportados na área de trabalho e online (por exemplo, Word 2016 ou posterior em Windows e Word na Web).

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

Um grupo de pontos de extensão da interface do usuário em uma guia. Um grupo pode ter até seis controles. O **atributo id** é necessário e cada **id** deve ser exclusivo entre todos os grupos no manifesto. A **id é** uma cadeia de caracteres com no máximo 125 caracteres. Confira [Elemento Group](group.md)

## <a name="officetab-example"></a>Exemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="Contoso.msgreadTabMessage.group1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
