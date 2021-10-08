---
title: Office namespace - conjunto de requisitos 1.4
description: Office namespace disponíveis para Outlook que usam conjunto de requisitos da API de Caixa de Correio 1.4.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: d4f351dc22dd22ac92dc7cc6bec53369471ed634
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237592"
---
# <a name="office-mailbox-requirement-set-14"></a>Office (conjunto de requisitos de caixa de correio 1.4)

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office?view=outlook-js-1.4&preserve-view=true).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Modo do Outlook aplicável](../../../outlook/outlook-add-ins-overview.md#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [context](office.context.md) | Escrever<br>Leitura | [Context](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>Enumerações

| Enumeração | Modos | Tipo de retorno | Minimum<br>conjunto de requisitos |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Escrever<br>Leitura | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Escrever<br>Leitura | Cadeia de caracteres | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [SourceProperty](#sourceproperty-string) | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>Namespaces

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .

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
