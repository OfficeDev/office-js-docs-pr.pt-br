---
title: Namespace do Office – conjunto de requisitos 1,4
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 2617ba3c80f44b1cddab568f94044a95a3061065
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064674"
---
# <a name="office"></a>Office

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

### <a name="namespaces"></a>Namespaces

[context](Office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.

### <a name="members"></a>Membros

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|
