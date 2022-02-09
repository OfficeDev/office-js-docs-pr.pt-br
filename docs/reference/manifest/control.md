---
title: Elemento Control no arquivo de manifesto
description: Define um controle que executa uma ação ou inicia um painel de tarefas.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa7ff9b0162070b378352ce187de15a34323b998
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467833"
---
# <a name="control-element"></a>Elemento Control

Define um controle que executa uma ação ou inicia um painel de tarefas. Um elemento **Control** pode ser um botão ou um menu. Pelo menos um **Control** deve ser incluído em um elemento [Group](group.md).

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (Para um complemento do painel de tarefas.)
- Alguns elementos filho podem estar associados a conjuntos de requisitos adicionais.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|**xsi:type**|Sim|O tipo de controle que está sendo definido. Pode ser `Button`, `Menu`ou `MobileButton`. |
|**id**|Sim|A ID do elemento Control. Pode ter no máximo 125 caracteres. Deve ser exclusivo em todos os **elementos Control** no manifesto.|

> [!NOTE]
> O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1. Ele só se aplica aos elementos **Control** contidos em um elemento [MobileFormFactor](mobileformfactor.md).

## <a name="child-elements"></a>Elementos filho

Os elementos filho válidos dependem do valor do **atributo xsi:type** .

- [Tipo de botão do elemento Control](control-button.md)
- [Tipo de menu do elemento Control](control-menu.md)
- [Tipo mobileButton do elemento Control](control-mobilebutton.md)
