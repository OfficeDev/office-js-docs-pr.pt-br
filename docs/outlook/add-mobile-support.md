---
title: Adicionar suporte móvel a um suplemento do Outlook
description: A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: a4fb02fee8bb429d0193903ba03fcee17b7ede48
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607614"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Adicionar suporte para comandos de suplementos para Outlook Mobile

O uso de comandos de suplemento no Outlook Mobile permite que os usuários acessem a mesma funcionalidade (com algumas [limitações](#code-considerations)) já existentes no Outlook na Web, no Windows e no Mac. A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.

## <a name="updating-the-manifest"></a>Atualização do manifesto

A primeira etapa para habilitar os comandos de suplemento no Outlook Mobile é defini-los no manifesto do suplemento. O esquema [VersionOverrides](../reference/manifest/versionoverrides.md) versão 1.1 define um novo fator forma para dispositivos móveis, o [MobileFormFactor](../reference/manifest/mobileformfactor.md).

Esse elemento contém todas as informações para carregar o suplemento em clientes móveis. Isso permite que você defina elementos de interface completamente diferentes e arquivos JavaScript para a experiência móvel.

O exemplo a seguir mostra um único botão de painel de tarefas em um `MobileFormFactor` elemento.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

Isso é muito semelhante aos elementos que aparecem em um elemento [DesktopFormFactor](../reference/manifest/desktopformfactor.md), com algumas diferenças importantes.

- O elemento [OfficeTab](../reference/manifest/officetab.md) não é usado.
- O elemento [ExtensionPoint](../reference/manifest/extensionpoint.md) deve ter apenas um elemento filho. Se o suplemento apenas adiciona um botão, o elemento filho deve ser um elemento [Control](../reference/manifest/control.md). Se o suplemento adiciona mais de um botão, o elemento filho deve ser um elemento [Group](../reference/manifest/group.md) que contém vários elementos `Control`.
- Não há nenhum tipo `Menu` equivalente ao elemento `Control`.
- O elemento [Supertip](../reference/manifest/supertip.md) não é usado.
- Os tamanhos de ícone obrigatórios são diferentes. Suplementos móveis devem, no mínimo, dar suporte a ícones de 25 x 25, 32 x 32 e 48 x 48 pixels.

## <a name="code-considerations"></a>Considerações sobre código

Criar um suplemento para o Mobile traz algumas considerações adicionais.

### <a name="use-rest-instead-of-exchange-web-services"></a>Usar REST em vez de Serviços Web do Exchange

O método [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) não é suportado no Outlook Mobile. Os suplementos devem preferir obter as informações da API Office.js sempre que possível. Se os suplementos exigem informações que não são expostas pela API Office.js devem usar as [APIs REST do Outlook](/outlook/rest/) para acessar as caixas de correio do usuário.

O conjunto de requisitos de caixa de correio 1,5 introduziu uma nova versão do [Office. Context. Mailbox. getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) que pode solicitar um token de acesso compatível com as APIs REST e uma nova propriedade [Office. Context. Mailbox. restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) que pode ser usada para localizar o ponto de extremidade da API REST para o usuário.

### <a name="pinch-zoom"></a>Pinçar e zoom

Por padrão, os usuários podem usar o gesto de “pinçar/zoom” para aplicar zoom aos painéis de tarefas. Se isso não fizer sentido em seu cenário, desative esse recurso em seu HTML.

### <a name="close-task-panes"></a>Fechar painéis de tarefas

Nos Outlook Mobile, os painéis de tarefa ocupam a tela inteira e, por padrão, exigem que o usuário os feche para retornar à mensagem. Considere o uso do método [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) para fechar o painel de tarefas quando seu cenário estiver concluído.

### <a name="compose-mode-and-appointments"></a>Modo de redação e compromissos

Atualmente, os suplementos do Outlook Mobile dão suporte à ativação apenas durante a leitura de mensagens. Os suplementos não são ativados ao redigir mensagens ou ao exibir ou redigir compromissos. No entanto, os suplementos integrados do provedor de reunião online podem ser ativados no modo organizador de compromisso. Confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](online-meeting.md) para saber mais sobre essa exceção.

### <a name="unsupported-apis"></a>APIs sem suporte

As APIs introduzidas no conjunto de requisitos 1,6 ou posterior não são suportadas pelo Outlook Mobile. As seguintes APIs de conjuntos de requisitos anteriores também não são suportadas.

  - [Office.context.officeTheme](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [Office.context.mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [Office.context.mailbox.convertToEwsId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [Office.context.mailbox.item.displayReplyAllForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.displayReplyForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a>Confira também

[Suporte ao conjunto de requisitos](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)