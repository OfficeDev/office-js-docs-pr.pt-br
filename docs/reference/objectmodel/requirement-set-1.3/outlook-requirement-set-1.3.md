---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.3
description: Recursos e APIs que foram introduzidos para Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.3.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 74ffc618f0f3555eef47abb38bb5118ac7177b9a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152124"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.3

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-13"></a>Novidades na versão 1.3?

O conjunto de requisitos 1.3 inclui todos os recursos do conjunto [de requisitos 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Ele adicionou os seguintes recursos.

- Foi adicionado o suporte para [comandos de suplemento](../../../outlook/add-in-commands-for-outlook.md).
- Foi adicionada a capacidade para salvar ou fechar um item que está sendo composto.
- Objeto [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) aprimorado para permitir que os complementos recebam ou deem o corpo inteiro.
- Foram adicionados os métodos de conversão para converter IDs entre os formatos EWS e REST.
- Mais capacidade de adicionar mensagens de notificação à barra de informações nos itens.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#getAsync_coercionType__options__callback_): Retorna o corpo atual em um formato especificado.
- Foi adicionado o [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#setAsync_data__options__callback_): Substitui todo o corpo com o texto especificado.
- Foi adicionado o objeto [Event](/javascript/api/office/office.addincommands.event): Passado como um parâmetro para funções de comando sem interface de usuário em um suplemento do Outlook. Usado para sinalizar a conclusão do processamento.
- Foi adicionado o [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods): Fecha o item atual que está sendo composto.
- Foi adicionado o [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods): Salva um item de forma assíncrona.
- Foi adicionado o [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties): Obtém as mensagens de notificação de um item.
- Foi adicionado o [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods): Converte uma ID de item formatada para REST em formato EWS.
- Foi adicionado o [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods): Converte uma ID de item formatada para EWS em formato REST.
- Foi adicionado o [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3&preserve-view=true): Especifica o tipo de mensagem de notificação para um compromisso ou uma mensagem.
- Foi adicionado o [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3&preserve-view=true): Especifica a versão da API REST que corresponde a uma ID de item formatado para REST.
- Foi adicionado o objeto [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true): Fornece métodos para acessar as mensagens de notificação em um suplemento do Outlook.
- Foi adicionado o tipo [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3&preserve-view=true): Retornado pelo método `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
