---
title: Adicionar suporte móvel a um suplemento do Outlook
description: Saiba como adicionar suporte para o Outlook Mobile, incluindo como atualizar o manifesto do suplemento e alterar seu código para cenários móveis, se necessário.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c84b4aeb04cd2c8b3c2f0a7afa9fd1631c22afc5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607538"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Adicionar suporte para comandos de suplementos para Outlook Mobile

O uso de comandos de suplemento no Outlook Mobile permite que os usuários acessem a mesma funcionalidade (com algumas [limitações) que](#code-considerations) eles já têm no Outlook na Web, Windows e Mac. A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.

## <a name="updating-the-manifest"></a>Atualização do manifesto

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

O exemplo a seguir mostra um único botão do painel de tarefas em um `MobileFormFactor` elemento.

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

Isso é muito semelhante aos elementos que aparecem em um elemento [DesktopFormFactor](/javascript/api/manifest/desktopformfactor), com algumas diferenças importantes.

- O elemento [OfficeTab](/javascript/api/manifest/officetab) não é usado.
- The [ExtensionPoint](/javascript/api/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](/javascript/api/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](/javascript/api/manifest/group) element that contains multiple `Control` elements.
- Não há nenhum tipo `Menu` equivalente ao elemento `Control`.
- O elemento [Supertip](/javascript/api/manifest/supertip) não é usado.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## <a name="code-considerations"></a>Considerações sobre código

Criar um suplemento para o Mobile traz algumas considerações adicionais.

### <a name="use-rest-instead-of-exchange-web-services"></a>Usar REST em vez de Serviços Web do Exchange

The [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.

O conjunto de requisitos de caixa de correio 1.5 introduziu uma nova versão do [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) que pode solicitar um token de acesso compatível com as APIs REST e uma nova propriedade [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) que pode ser usada para localizar o ponto de extremidade da API REST para o usuário.

### <a name="pinch-zoom"></a>Pinçar e zoom

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### <a name="close-task-panes"></a>Fechar painéis de tarefas

In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) method to close the task pane when your scenario is complete.

### <a name="compose-mode-and-appointments"></a>Modo de redação e compromissos

Atualmente, os suplementos no Outlook Mobile só dão suporte à ativação durante a leitura de mensagens. Os suplementos não são ativados ao redigir mensagens ou ao exibir ou redigir compromissos. No entanto, há duas exceções:

1. Os suplementos integrados do provedor de reunião online podem ser ativados no modo Organizador de Compromissos. Para obter mais informações sobre essa exceção (incluindo APIs disponíveis), consulte Criar um suplemento [móvel do Outlook para um provedor de reunião online](online-meeting.md#available-apis).
1. Os suplementos que registram anotações de compromisso e outros detalhes do CRM (gerenciamento de relacionamento com o cliente) ou serviços de anotações podem ser ativados no modo Participante do Compromisso. Para obter mais informações sobre essa exceção (incluindo APIs disponíveis), consulte as anotações de compromisso de log para um aplicativo externo nos [suplementos móveis do Outlook](mobile-log-appointments.md#available-apis).

### <a name="unsupported-apis"></a>APIs sem suporte

As APIs introduzidas no conjunto de requisitos 1.6 ou posterior não são compatíveis com o Outlook Mobile. As APIs a seguir de conjuntos de requisitos anteriores também não têm suporte.

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## <a name="see-also"></a>Confira também

[Conjuntos de requisitos suportados pelos Exchange Servers e clientes do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)