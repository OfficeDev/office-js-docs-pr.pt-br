---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 8e33d40bec9b561c642ad6e0da73ae13a18378b6
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695900"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [accountType](#accounttype-string) | Membro |
| [displayName](#displayname-string) | Membro |
| [emailAddress](#emailaddress-string) | Membro |
| [timeZone](#timezone-string) | Membro |

### <a name="members"></a>Membros

#### <a name="accounttype-string"></a>AccountType: cadeia de caracteres

> [!NOTE]
> Atualmente, esse membro só tem suporte no Outlook 2016 ou posterior no Mac (Build 16.9.1212 ou posterior).

Obtém o tipo de conta do usuário associado à caixa de correio. Os valores possíveis estão listados na tabela a seguir.

| Valor | Descrição |
|-------|-------------|
| `enterprise` | A caixa de correio está em um servidor Exchange local. |
| `gmail` | A caixa de correio está associada a uma conta do gmail. |
| `office365` | A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365. |
| `outlookCom` | A caixa de correio está associada a uma conta pessoal do Outlook.com. |

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a>displayName: cadeia de caracteres

Obtém o nome de exibição do usuário.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a>emailAddress: cadeia de caracteres

Obtém o endereço de email SMTP do usuário.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a>timeZone: cadeia de caracteres

Obtém o fuso horário padrão do usuário.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
