---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.3
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 83d15038aedeeae2ce873207f17fe61dd1d3b3a7
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064694"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.3

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente. 

## <a name="whats-new-in-13"></a>Novidades na versão 1.3?

O conjunto de requisitos 1.3 inclui todos os recursos do [Conjunto de requisitos 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Ele adicionou os seguintes recursos.

- Foi adicionado o suporte para [comandos de suplemento](/outlook/add-ins/add-in-commands-for-outlook).
- Foi adicionada a capacidade para salvar ou fechar um item que está sendo composto.
- O objeto [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3) foi aprimorado para permitir que os suplementos obtenham ou definam todo o corpo.
- Foram adicionados os métodos de conversão para converter IDs entre os formatos EWS e REST.
- Mais capacidade de adicionar mensagens de notificação à barra de informações nos itens.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-): Retorna o corpo atual em um formato especificado.
- Foi adicionado o [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-): Substitui todo o corpo com o texto especificado.
- Foi adicionado o objeto [Event](/javascript/api/office/office.addincommands.event): Passado como um parâmetro para funções de comando sem interface de usuário em um suplemento do Outlook. Usado para sinalizar a conclusão do processamento.
- Foi adicionado o [Office.context.mailbox.item.close](office.context.mailbox.item.md#close): Fecha o item atual que está sendo composto.
- Foi adicionado o [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback): Salva um item de forma assíncrona.
- Foi adicionado o [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessages): Obtém as mensagens de notificação de um item.
- Foi adicionado o [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string): Converte uma ID de item formatada para REST em formato EWS.
- Foi adicionado o [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string): Converte uma ID de item formatada para EWS em formato REST.
- Foi adicionado o [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3): Especifica o tipo de mensagem de notificação para um compromisso ou uma mensagem.
- Foi adicionado o [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3): Especifica a versão da API REST que corresponde a uma ID de item formatado para REST.
- Foi adicionado o objeto [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3): Fornece métodos para acessar as mensagens de notificação em um suplemento do Outlook.
- Foi adicionado o tipo [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3): Retornado pelo método `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/outlook/add-ins/quick-start)
