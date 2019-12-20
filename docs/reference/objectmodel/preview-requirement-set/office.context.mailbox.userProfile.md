---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos de visualização
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: be95b38764c49084950b3bcb0aeb2394f1a4b940
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815050"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile

Fornece informações sobre o usuário em um suplemento do Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Mínimo<br>nível de permissão | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#accounttype) | ReadItem | Escrever<br>Leitura | String | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayName](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#displayname) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [emailAddress](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#emailaddress) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
