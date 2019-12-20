---
title: Namespace do Office – conjunto de requisitos de visualização
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ef9634058fcdc633e9ad3a0adb74c4abebf8038b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815057"
---
# <a name="office"></a>Office

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="properties"></a>Propriedades

| Propriedade | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| [context](office.context.md) | Escrever<br>Leitura | [Context](/javascript/api/office/office.context?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Enumerações

| Enumeração | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Escrever<br>Leitura | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Namespaces

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.

## <a name="enumeration-details"></a>Detalhes da enumeração

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: cadeia de caracteres

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType: cadeia de caracteres

Especifica como forçar dados retornados ou definidos pelo método invocado.

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="eventtype-string"></a>EventType: cadeia de caracteres

Especifica o evento associado a um manipulador de eventos.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

| Nome | Tipo | Descrição | Conjunto de requisitos mínimo |
|---|---|---|:---:|
|`AppointmentTimeChanged`| String | A data ou hora do compromisso ou série selecionado foi alterada. | 1.7 |
|`AttachmentsChanged`| String | Um anexo foi adicionado ou removido do item. | 1,8 |
|`EnhancedLocationsChanged`| String | O local do compromisso selecionado foi alterado. | 1,8 |
|`ItemChanged`| String | Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado. | 1,5 |
|`OfficeThemeChanged`| String | O tema do Office na caixa de correio foi alterado. | Visualização |
|`RecipientsChanged`| String | A lista de destinatários do item selecionado ou local do compromisso foi alterada. | 1.7 |
|`RecurrenceChanged`| String | O padrão de recorrência da série selecionada foi alterado. | 1.7 |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5 |
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler |

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty: cadeia de caracteres

Especifica a origem dos dados retornados pelo método chamado.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

|Nome| Tipo| Descrição|
|---|---|---|
|`Body`| String|A origem dos dados é o corpo de uma mensagem.|
|`Subject`| Cadeia de caracteres|A origem dos dados é o assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|
