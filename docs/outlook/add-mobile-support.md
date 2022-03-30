---
title: Adicionar suporte móvel a um suplemento do Outlook
description: A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6e99c862d4cd63590a86c757bf2b720c096826a9
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496968"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Adicionar suporte para comandos de suplementos para Outlook Mobile

O uso de comandos de Outlook Mobile permite que os usuários acessem a mesma funcionalidade (com algumas limitações[) que](#code-considerations) eles já têm no Outlook na Web, Windows e Mac. A adição de suporte para o Outlook Mobile requer atualização do manifesto do suplemento e, possivelmente, a alteração do código para cenários móveis.

## <a name="updating-the-manifest"></a>Atualização do manifesto

A primeira etapa para habilitar os comandos de suplemento no Outlook Mobile é defini-los no manifesto do suplemento. O esquema [VersionOverrides](/javascript/api/manifest/versionoverrides) versão 1.1 define um novo fator forma para dispositivos móveis, o [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

Esse elemento contém todas as informações para carregar o suplemento em clientes móveis. Isso permite que você defina elementos de interface completamente diferentes e arquivos JavaScript para a experiência móvel.

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
- O elemento [ExtensionPoint](/javascript/api/manifest/extensionpoint) deve ter apenas um elemento filho. Se o suplemento apenas adiciona um botão, o elemento filho deve ser um elemento [Control](/javascript/api/manifest/control). Se o suplemento adiciona mais de um botão, o elemento filho deve ser um elemento [Group](/javascript/api/manifest/group) que contém vários elementos `Control`.
- Não há nenhum tipo `Menu` equivalente ao elemento `Control`.
- O elemento [Supertip](/javascript/api/manifest/supertip) não é usado.
- Os tamanhos de ícone obrigatórios são diferentes. Suplementos móveis devem, no mínimo, dar suporte a ícones de 25 x 25, 32 x 32 e 48 x 48 pixels.

## <a name="code-considerations"></a>Considerações sobre código

Criar um suplemento para o Mobile traz algumas considerações adicionais.

### <a name="use-rest-instead-of-exchange-web-services"></a>Usar REST em vez de Serviços Web do Exchange

O método [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) não é suportado no Outlook Mobile. Os suplementos devem preferir obter as informações da API Office.js sempre que possível. Se os suplementos exigem informações que não são expostas pela API Office.js devem usar as [APIs REST do Outlook](/outlook/rest/) para acessar as caixas de correio do usuário.

O conjunto de requisitos de caixa de correio 1.5 introduziu uma nova versão de [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) que pode solicitar um token de acesso compatível com as APIs REST e uma nova [propriedade Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) que pode ser usada para encontrar o ponto de extremidade da API REST para o usuário.

### <a name="pinch-zoom"></a>Pinçar e zoom

Por padrão, os usuários podem usar o gesto de “pinçar/zoom” para aplicar zoom aos painéis de tarefas. Se isso não fizer sentido em seu cenário, desative esse recurso em seu HTML.

### <a name="close-task-panes"></a>Fechar painéis de tarefas

Nos Outlook Mobile, os painéis de tarefa ocupam a tela inteira e, por padrão, exigem que o usuário os feche para retornar à mensagem. Considere o uso do método [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) para fechar o painel de tarefas quando seu cenário estiver concluído.

### <a name="compose-mode-and-appointments"></a>Modo de redação e compromissos

Atualmente, os complementos no Outlook Mobile só suportam a ativação ao ler mensagens. Os suplementos não são ativados ao redigir mensagens ou ao exibir ou redigir compromissos. No entanto, os complementos integrados do provedor de reunião online podem ser ativados no modo Organizador de Compromissos. Para saber mais sobre essa exceção (incluindo APIs disponíveis), consulte [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md#available-apis).

### <a name="unsupported-apis"></a>APIs sem suporte

AS APIs introduzidas no conjunto de requisitos 1.6 ou posterior não são suportadas pelo Outlook Mobile. As APIs a seguir de conjuntos de requisitos anteriores também não são suportadas.

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