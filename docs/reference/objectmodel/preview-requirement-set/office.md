---
title: Namespace do Office – conjunto de requisitos de pré-visualização
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: bbec602680da7914666daf33ed36c45751ae69c6
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068319"
---
# <a name="office"></a>Office

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Membro |
| [CoercionType](#coerciontype-string) | Membro |
| [EventType](#eventtype-string) | Membro |
| [SourceProperty](#sourceproperty-string) | Membro |

### <a name="namespaces"></a>Namespaces

[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.

### <a name="members"></a>Membros

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Especifica o resultado de uma chamada assíncrona.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Succeeded`| String|A chamada foi bem-sucedida.|
|`Failed`| String|Falha na chamada.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Especifica como forçar os dados retornados ou definir de acordo com o método chamado.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Html`| String|Solicita que os dados sejam retornados no formato HTML.|
|`Text`| String|Solicita que os dados sejam retornados no formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

---

####  <a name="eventtype-string"></a>EventType :String

Especifica o evento associado a um manipulador de eventos.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

| Nome | Tipo | Descrição | Conjunto de requisitos mínimos |
|---|---|---|---|
|`AppointmentTimeChanged`| String | A data ou hora da série ou do compromisso selecionado foi alterada. | 1.7 |
|`AttachmentsChanged`| String | Um anexo foi adicionado a ou removido do item. | Visualização |
|`ItemChanged`| Cadeia de caracteres | Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado. | 1,5 |
|`OfficeThemeChanged`| Cadeia de caracteres | O tema do Office na caixa de correio foi alterado. | Visualização |
|`RecipientsChanged`| Cadeia de caracteres | A lista de destinatários do item selecionado ou o local do compromisso foi alterado. | 1.7 |
|`RecurrenceChanged`| String | O padrão de recorrência da série selecionada foi alterado. | 1.7 |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Especifica a origem dos dados retornados pelo método chamado.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Body`| String|A origem dos dados é o corpo de uma mensagem.|
|`Subject`| String|A origem dos dados é o assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|
