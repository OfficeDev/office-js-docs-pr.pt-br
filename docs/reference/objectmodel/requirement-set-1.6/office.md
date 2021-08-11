---
title: Office namespace - conjunto de requisitos 1.6
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.6.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 3fef16b322cd3ded8c3df9ef051f897a6d2fe29de928ca420f476cf3ecb6552a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097414"
---
# <a name="office-mailbox-requirement-set-16"></a>Office (conjunto de requisitos de caixa de correio 1.6)

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [context](office.context.md) | Escrever<br>Ler | [Context](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>Enumerações

| Enumeração | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Escrever<br>Ler | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Escrever<br>Ler | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Escrever<br>Ler | Cadeia de Caracteres | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Escrever<br>Ler | Cadeia de Caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>Namespaces

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .

## <a name="enumeration-details"></a>Detalhes da enumeração

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: String

Especifica o resultado de uma chamada assíncrona.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades

|Nome| Tipo| Descrição|
|---|---|---|
|`Succeeded`| String|A chamada foi bem-sucedida.|
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
|`Html`| String|Solicita que os dados sejam retornados no formato HTML.|
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
|`ItemChanged`| Cadeia de Caracteres | Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado. | 1,5 |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5 |
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler |

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
|`Body`| String|A origem dos dados é o corpo de uma mensagem.|
|`Subject`| String|A origem dos dados é o assunto de uma mensagem.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|
