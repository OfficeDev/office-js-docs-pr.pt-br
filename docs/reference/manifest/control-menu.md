---
title: Elemento Control do tipo Menu no arquivo de manifesto
description: Define um menu cujos itens podem executar ações ou iniciar painéis de tarefas.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7287b8e2cdf2378140ef50a41306820a0fd4002f
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467880"
---
# <a name="control-element-of-type-menu"></a>Elemento Control do tipo Menu

Um menu define uma lista de opções. Cada item de menu executa uma função ou mostra um painel de tarefas.

> [!NOTE]
> Este artigo supõe familiaridade com o artigo de referência de [Controle](control.md) básico que contém informações importantes sobre os atributos do elemento.

O controle de menu define:

- Um controle de menu no nível raiz.
- Uma lista de itens de menu.

Quando usado com o **ponto de extensão PrimaryCommandSurface**[, o](extensionpoint.md) item de menu raiz é exibido como um botão na faixa de opções. Quando o botão é selecionado, o menu é exibido como uma lista lista listada. Não há suporte para submenus.

Quando usado com o **ponto de extensão ContextMenu**[, um](extensionpoint.md) item de menu raiz é exibido no menu de contexto. Quando o item raiz é selecionado, os itens de menu são exibidos como um submenu. Nenhum dos itens pode ser um submenu porque apenas um nível de submenus é suportado.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Label](#label)     | Sim |  O texto do menu. |
|  **Dica de Ferramenta**    |Não|A dica de ferramenta do menu. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um **elemento String** . O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).|
|  [Supertip](supertip.md)  | Sim |  A super dica para este menu.    |
|  [Icon](icon.md)      | Sim |  Uma imagem para o menu.         |
|  **Items**     | Sim |  Uma coleção de itens a ser exibido no menu. Contém o **elemento Item** para cada item. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Não |  Especifica se o menu deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas. Se usado, ele deve ser o *primeiro* elemento filho. |

### <a name="label"></a>Rótulo

Especifica o texto do nome do menu por meio de seu único atributo, **resid**, que pode não ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no filho **ShortStrings** do elemento [Resources](resources.md) .

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="examples"></a>Exemplos

No exemplo a seguir, o menu tem dois itens. O primeiro exibe um painel de tarefas. O segundo executa uma função. O menu foi configurado *para não ficar* visível quando o complemento está sendo executado em uma plataforma que dá suporte a guias contextuais. Para obter mais informações, leia [Implemente uma experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

```xml
<Control xsi:type="Menu" id="Contoso.TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="GetData">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getData</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

No exemplo a seguir, o segundo item do menu é configurado para não ficar  visível quando o complemento está sendo executado em uma plataforma que dá suporte a guias contextuais. Para obter mais informações, leia [Implemente uma experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

```xml
<Control xsi:type="Menu" id="Contoso.msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```
