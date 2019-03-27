---
title: Namespace do Office – conjunto de requisitos 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d8c51646818681629fa0c184962776beffe22a55
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871413"
---
# <a name="office"></a>Office

O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Member |
| [CoercionType](#coerciontype-string) | Member |
| [EventType](#eventtype-string) | Member |
| [SourceProperty](#sourceproperty-string) | Membro |

### <a name="namespaces"></a>Namespaces

[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.

[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.

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
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

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
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

---

####  <a name="eventtype-string"></a>EventType :String

Especifica o evento associado a um manipulador de eventos.

##### <a name="type"></a>Tipo

*   String

##### <a name="properties"></a>Propriedades:

| Nome | Tipo | Descrição |
|---|---|---|
|`ItemChanged`| String | Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado. |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

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
