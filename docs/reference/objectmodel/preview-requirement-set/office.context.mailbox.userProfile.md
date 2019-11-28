---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos de visualização
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 4afc64f247155576ab3f0024d1929a29a0f7dc0c
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629255"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="properties"></a>Propriedades

| Propriedade | Mínimo<br>nível de permissão | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|---|
| [accountType](#accounttype-string) | ReadItem | Escrever<br>Ler | String | 1.6 |
| [displayName](#displayname-string) | ReadItem | Escrever<br>Ler | String | 1.0 |
| [emailAddress](#emailaddress-string) | ReadItem | Escrever<br>Ler | String | 1.0 |
| [timeZone](#timezone-string) | ReadItem | Escrever<br>Ler | String | 1.0 |

## <a name="property-details"></a>Detalhes da propriedade

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
