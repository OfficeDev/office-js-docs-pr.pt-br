---
title: Elemento Control do tipo MobileButton no arquivo de manifesto
description: Define um botão em um dispositivo móvel que executa uma ação ou inicia um painel de tarefas.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d498b728bf7f19cf239ffc6178f19cdf9a62de58
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467888"
---
# <a name="control-element-of-type-mobilebutton"></a>Elemento Control do tipo MobileButton

Define um botão que executa uma ação ou inicia um painel de tarefas e que aparece apenas em plataformas móveis.

> [!NOTE]
> Este artigo supõe familiaridade com o artigo de referência de [Controle](control.md) básico que contém informações importantes sobre os atributos do elemento.

Um botão móvel executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um `id` valor de atributo exclusivo entre todos os elementos **Control** no manifesto.

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Label](#label)     | Sim |  O texto do botão. |
|  [Icon](icon.md)      | Sim |  Uma imagem para o botão.         |
|  [Action](action.md)    | Sim |  Especifica a ação a ser executada. Pode haver apenas um **filho action** de um **elemento Control** . |

### <a name="label"></a>Rótulo

Especifica o texto do botão por meio de seu único atributo, **resid**, que não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no **filho ShortStrings** do elemento [Resources](resources.md) .

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="examples"></a>Exemplos

No exemplo a seguir, o botão executa uma função.

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

No exemplo a seguir, o botão exibe um painel de tarefas.

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
