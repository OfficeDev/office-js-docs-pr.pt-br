---
title: Office. Context. Mailbox. Diagnostics – conjunto de requisitos 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: a2d54aee0545fb43aea798d799f190519226d185
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268506"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Fornece informações de diagnóstico para um suplemento do Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [hostName](#hostname-string) | Membro |
| [hostVersion](#hostversion-string) | Membro |
| [OWAView](#owaview-string) | Membro |

### <a name="members"></a>Membros

#### <a name="hostname-string"></a>Nome do host: cadeia de caracteres

Obtém uma cadeia de caracteres que representa o nome do aplicativo host.

Uma cadeia de caracteres que pode ser um dos seguintes valores `Outlook`: `OutlookIOS`, ou `OutlookWebApp`.

> [!NOTE]
> O `Outlook` valor é retornado para o Outlook em clientes de área de trabalho (ou seja, Windows e Mac).

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

#### <a name="hostversion-string"></a>hostVersion: cadeia de caracteres

Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do servidor Exchange (por exemplo, "15.0.468.0").

Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou `hostVersion` Ios, a propriedade retornará a versão do aplicativo host, Outlook. No Outlook na Web, a propriedade retorna a versão do servidor Exchange.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

#### <a name="owaview-string"></a>OWAView: cadeia de caracteres

Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook na Web.

A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.

Se o aplicativo host não for o Outlook na Web, então acessar essa propriedade resultará `undefined`em.

O Outlook na Web tem três exibições que correspondem à largura da tela e à janela e ao número de colunas que podem ser exibidas:

*   `OneColumn`, que é exibido quando a tela é estreita. O Outlook na Web usa esse layout de coluna única em toda a tela de um smartphone.
*   `TwoColumns`, que é exibido quando a tela é mais larga. O Outlook na Web usa esse modo de exibição na maioria dos Tablets.
*   `ThreeColumns`, que é exibido quando a tela é ainda mais larga. Por exemplo, o Outlook na Web usa esse modo de exibição em uma janela de tela inteira em um computador desktop.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|
