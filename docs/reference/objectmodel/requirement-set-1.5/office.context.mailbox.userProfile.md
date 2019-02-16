---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e98e88cde184db121e69fdd267dff4e39d887b1f
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067823"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [displayName](#displayname-string) | Membro |
| [emailAddress](#emailaddress-string) | Membro |
| [timeZone](#timezone-string) | Membro |

### <a name="members"></a>Membros

####  <a name="displayname-string"></a>displayName :String

Obtém o nome de exibição do usuário.

##### <a name="type"></a>Tipo

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

##### <a name="example"></a>Exemplo

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

Obtém o endereço de email SMTP do usuário.

##### <a name="type"></a>Tipo

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

##### <a name="example"></a>Exemplo

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

Obtém o fuso horário padrão do usuário.

##### <a name="type"></a>Tipo

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|

##### <a name="example"></a>Exemplo

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
