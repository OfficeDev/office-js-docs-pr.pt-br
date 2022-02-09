---
title: Elemento Control do tipo Button no arquivo de manifesto
description: Define um botão que executa uma ação ou inicia um painel de tarefas.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: adc58424fe9898bffcbd9e16bed8f3b13b9df4a2
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467877"
---
# <a name="control-element-of-type-button"></a>Elemento Control do tipo Button

Define um botão que executa uma ação ou inicia um painel de tarefas.

> [!NOTE]
> Este artigo supõe familiaridade com o artigo de referência de [Controle](control.md) básico que contém informações importantes sobre os atributos do elemento.

Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um `id` valor de atributo exclusivo entre todos os elementos **Control** no manifesto.

> [!IMPORTANT]
> Os controles de tipo "Button" são ignorados em plataformas móveis. Para dar suporte a plataformas móveis, você também deve ter um controle do tipo "MobileButton" para cada controle do tipo "Button".

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Label](#label)     | Sim |  O texto do botão. |
|  **Dica de Ferramenta**    |Não|A dica de ferramenta do botão. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um **elemento String** . O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).|
|  [Supertip](supertip.md)  | Sim |  A dica detalhada do botão.    |
|  [Icon](icon.md)      | Sim |  Uma imagem para o botão.         |
|  [Action](action.md)    | Sim |  Especifica a ação a ser executada. Pode haver apenas um **filho action** de um **elemento Control** . |
|  [Enabled](enabled.md)    | Não |  Especifica se o controle está habilitado quando o complemento é lançado.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Não |  Especifica se o botão deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas. Se usado, ele deve ser o *primeiro* elemento filho. |

### <a name="label"></a>Rótulo

Especifica o texto do botão por meio de seu único atributo, **resid**, que não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no **filho ShortStrings** do elemento [Resources](resources.md) .

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

No exemplo a seguir, o botão executa uma função. Ele também é configurado para ser desabilitado quando o complemento é lançado. Ele pode ser habilitado programaticamente. Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).

```xml
<Control xsi:type="Button" id="Contoso.msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
  <Enabled>false</Enabled>
</Control>
```

No exemplo a seguir, o botão exibe um painel de tarefas.

```xml
<Control xsi:type="Button" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
