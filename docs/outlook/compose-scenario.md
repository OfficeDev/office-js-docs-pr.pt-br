---
title: Crie suplementos do Outlook para formulários de redação
description: Saiba mais sobre os cenários e recursos dos suplementos do Outlook nos formulários de redação.
ms.date: 10/03/2022
ms.localizationpriority: high
ms.openlocfilehash: ef81b21eaa0bc63a5bf38757cb188e8850ade443
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467248"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a>Criar suplementos do Outlook para formulários de redação

Você pode criar suplementos de composição, que são suplementos do Outlook ativados em formulários de composição. Ao contrário dos suplementos de leitura (suplementos do Outlook que são ativados no modo de leitura quando um usuário está exibindo uma mensagem ou compromisso), os suplementos de composição estão disponíveis nos seguintes cenários de usuário.

- Redação de nova mensagem, solicitação de reunião ou compromisso em um formulário de redação.

- Exibição ou edição de compromisso existente, ou item de reunião no qual o usuário seja o organizador.

   > [!NOTE]
   > If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available. Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.

- Redação de uma mensagem de resposta embutida ou resposta a uma mensagem em um formulário de redação separado.

- Edição de uma resposta (**Aceitar**, **Provisório** ou **Recusar**) a uma solicitação de reunião ou a um item de reunião.

- Proposição de novo horário para um item de reunião.

- Encaminhamento ou resposta a uma solicitação de reunião ou a um item de reunião.

In each of these compose scenarios, any add-in command buttons defined by the add-in are shown. For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in. The following figure shows add-in commands in a compose form.

![Mostra um fomulário de criação do Outlook com comandos de suplementos.](../images/compose-form-commands.png)

A figura a seguir mostra o painel de seleção do suplemento composto por dois suplementos de redação que não implementam comandos de suplemento, ativado quando o usuário está compondo uma resposta embutida no Outlook.

![Aplicativo de email modelos ativado para item redigido.](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a>Tipos de suplementos disponíveis no modo de redação

Os suplementos de redação são implementados como [Comandos de suplemento para Outlook](add-in-commands-for-outlook.md). Para ativar suplementos para redação de emails ou respostas de reunião, os suplementos devem incluir um [elemento de ponto de extensão MessageComposeCommandSurface](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface) no manifesto. Para ativar suplementos para redação ou edição de compromissos ou reuniões em que o usuário é o organizador, os suplementos devem incluir um [elemento de ponto de extensão AppointmentOrganizerCommandSurface](/javascript/api/manifest/extensionpoint#appointmentorganizercommandsurface).

> [!NOTE]
> Os suplementos desenvolvidos para servidores ou clientes sem suporte para comandos de suplemento usam [regras de ativação](activation-rules.md) em um elemento [Rule](/javascript/api/manifest/rule) contido no elemento [OfficeApp](/javascript/api/manifest/officeapp). Os novos suplementos devem usar comandos de suplemento, exceto quando o suplemento for desenvolvido para servidores e clientes mais antigos.

## <a name="api-features-available-to-compose-add-ins"></a>Recursos de API disponíveis para suplementos de redação

- [Adicionar e remover anexos de um item em um formulário de redação no Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [Obter e definir dados de item em um formulário de redação no Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook](get-set-or-add-recipients.md)
- [Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook](get-or-set-the-subject.md)
- [Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook](insert-data-in-the-body.md)
- [Obter ou definir o local ao criar um compromisso no Outlook](get-or-set-the-location-of-an-appointment.md)
- [Obter ou definir a hora ao criar um compromisso no Outlook](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a>Confira também

- [Começar com os suplementos do Outlook para Office](../quickstarts/outlook-quickstart.md)
