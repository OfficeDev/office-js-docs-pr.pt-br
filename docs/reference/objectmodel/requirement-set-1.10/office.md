---
title: Office namespace - conjunto de requisitos 1.10
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.10.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e7b7ab9127ebf8ce9b7394d348144fe63b47de6c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592023"
---
# <a name="office-mailbox-requirement-set-110"></a>Office (conjunto de requisitos de caixa de correio 1.10)

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [context](office.context.md) | Escrever<br>Ler | [Context](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>Enumerações

| Enumeração | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Escrever<br>Ler | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Escrever<br>Ler | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Escrever<br>Ler | Cadeia de caracteres | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Escrever<br>Ler | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>Namespaces

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .

## <a name="enumeration-details"></a>Detalhes da enumeração

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: String

Especifica o resultado de uma chamada assíncrona.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades

|Nome| Tipo| Descrição|
|---|---|---|
|`Succeeded`| Cadeia de caracteres|A chamada foi bem-sucedida.|
|`Failed`| String|Falha na chamada.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType: String

Especifica como forçar dados retornados ou definidos pelo método invocado.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades

|Nome| Tipo| Descrição|
|---|---|---|
|`Html`| Cadeia de caracteres|Solicita que os dados sejam retornados no formato HTML.|
|`Text`| String|Solicita que os dados sejam retornados no formato de texto.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="eventtype-string"></a>EventType: String

Especifica o evento associado a um manipulador de eventos.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades

| Nome | Tipo | Descrição | Conjunto de requisitos mínimo |
|---|---|---|:---:|
|`AppointmentTimeChanged`| Cadeia de caracteres | A data ou hora do compromisso ou série selecionado foi alterada. | 1.7 |
|`AttachmentsChanged`| Cadeia de caracteres | Um anexo foi adicionado ou removido do item. | 1,8 |
|`EnhancedLocationsChanged`| Cadeia de caracteres | O local do compromisso selecionado foi alterado. | 1,8 |
|`ItemChanged`| Cadeia de caracteres | Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado. | 1,5 |
|`OfficeThemeChanged`| Cadeia de caracteres | O Office tema na caixa de correio foi alterado. | 1.10 |
|`RecipientsChanged`| Cadeia de caracteres | A lista de destinatários do item ou local do compromisso selecionado foi alterada. | 1.7 |
|`RecurrenceChanged`| Cadeia de caracteres | O padrão de recorrência da série selecionada foi alterado. | 1.7 |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5 |
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty: String

Especifica a origem dos dados retornados pelo método chamado.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades

|Nome| Tipo| Descrição|
|---|---|---|
|`Body`| Cadeia de caracteres|A origem dos dados é o corpo de uma mensagem.|
|`Subject`| String|A origem dos dados é o assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|