---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 12/17/2019
localization_priority: Priority
ms.openlocfilehash: a3cc49562add2f6fe54cf83d2f2ed64ebb61d8c7
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815043"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Conjunto de requisitos do modo de visualização de API para suplementos do Outlook

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!IMPORTANT]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdmethods"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)

<br>

---

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)

<br>

---

### <a name="sso"></a>SSO

#### <a name="officeruntimeauthgetaccesstokenofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[OfficeRuntime.auth.getAccessToken](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.

**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](/outlook/add-ins/quick-start)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
