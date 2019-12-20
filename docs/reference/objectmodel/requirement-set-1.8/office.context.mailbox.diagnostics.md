---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2c5ea33ffd8bc560288935f7ee65ebb93aadf1aa
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814247"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics

Fornece informações de diagnóstico para um suplemento do Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

## <a name="properties"></a>Propriedades

| Propriedade | Mínimo<br>nível de permissão | Modelos | Tipo de retorno | Mínimo<br>conjunto de requisitos |
|---|---|---|---|:---:|
| [hostName](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#hostname) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [hostVersion](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#hostversion) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [OWAView](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#owaview) | ReadItem | Escrever<br>Leitura | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
