---
title: 'Office.context.mailbox.diagnostics: conjunto de requisitos da visualização'
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: ae9fca6ea2f7de99b275989bb2129948a60bc86f
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432995"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Fornece informações de diagnóstico para um suplemento do Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [hostName](#hostname-string) | Membro |
| [hostVersion](#hostversion-string) | Membro |
| [OWAView](#owaview-string) | Membro |

### <a name="members"></a>Membros

####  <a name="hostname-string"></a>hostName :String

Obtém uma cadeia de caracteres que representa o nome do aplicativo host.

Uma cadeia de caracteres que pode ser um dos valores a seguir: `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

####  <a name="hostversion-string"></a>hostVersion :String

Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.

Se o suplemento de email estiver em execução no cliente do Outlook para área de trabalho ou Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a cadeia de caracteres `15.0.468.0`.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composição ou leitura|

####  <a name="owaview-string"></a>OWAView :String

Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook Web App.

A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.

Se o aplicativo host não for Outlook Web App, acessar essa propriedade resultará em `undefined`.

O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e à quantidade de colunas que pode ser exibida:

*   `OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.
*   `TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.
*   `ThreeColumns`, que é exibido quando a tela é ainda mais larga. Por exemplo, o Outlook Web App usa esse modo de exibição em um modo de tela cheia em um computador de mesa.

##### <a name="type"></a>Tipo:

*   Cadeia de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redação ou leitura|