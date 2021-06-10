---
title: Outlook conjunto de requisitos de visualização de API de complemento
description: Recursos e APIs que estão atualmente em visualização para Outlook de complementos.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: c7ca92e6a30f3109baff5721ae4e9930ef23dc56
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/09/2021
ms.locfileid: "52854008"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook conjunto de requisitos de visualização de API de complemento

O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.

> [!IMPORTANT]
> Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md). Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele. Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Você pode ser capaz de visualizar recursos em Outlook na Web configurando a versão direcionada em [seu locatário Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) "Configure preview access" é notado nesta página para recursos aplicáveis.
>
> Para outros recursos, você pode solicitar acesso aos bits de visualização para Outlook na Web usando sua conta Microsoft 365, concluindo e enviando [esse formulário](https://aka.ms/OWAPreview). "Solicitar acesso de visualização" é notado nesses recursos.

O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).

## <a name="features-in-preview"></a>Recursos no modo de visualização

Os seguintes recursos estão no modo de visualização.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Ativação do complemento em itens protegidos pelo IRM (Gerenciamento de Direitos de Informação)

Os complementos agora podem ser ativados em itens protegidos por IRM. Para ativar esse recurso, um administrador de locatário precisa habilitar o direito de uso definindo a opção Permitir política personalizada de acesso `OBJMODEL` programático em  Office. Confira [Direitos de uso e descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.

**Disponível em**: Outlook no Windows, começando com a com build 13229.10000 (conectada a uma assinatura Microsoft 365 de terceiros)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriedades de calendário adicionais

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a propriedade de evento de todos os dias de um compromisso no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Adicionado um novo número `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Ativação baseada em evento

Esse recurso foi lançado no [conjunto de requisitos 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). No entanto, eventos adicionais agora estão disponíveis na visualização. Para saber mais, confira [Eventos com suporte.](../../../outlook/autolaunch.md#supported-events)

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Integração à mensagens acionáveis

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)

<br>

---

---

### <a name="office-theme"></a>Tema do Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Capacidade adicional para obter o tema do Office.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Adicionado `OfficeThemeChanged` evento `Mailbox`.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)

<br>

---

---

### <a name="session-data"></a>Os dados da sessão

#### <a name="officesessiondata"></a>[Office. SessionData](/javascript/api/outlook/office.sessiondata)

Adicionado um novo objeto que representa os dados de sessão de um item.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo Redação.

**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
