---
title: Elemento RequestedHeight no arquivo de manifesto
description: O elemento RequestedHeight especifica a altura inicial (em pixels) de um suplemento de conteúdo ou email.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 853d12baf290167f3e6a635201e8b5d1d0e35a51
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720453"
---
# <a name="requestedheight-element"></a>Elemento RequestedHeight

Especifica a altura inicial (em pixels) de um suplemento de conteúdo ou de email. 

**Tipo de suplemento:** Conteúdo, Email

## <a name="syntax"></a>Sintaxe

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Contido em

- [DefaultSettings](defaultsettings.md) (suplementos de conteúdo) com um valor entre 32 e 1000
- [DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (suplementos de email ) com um valor entre 32 e 450
- [ExtensionPoint](extensionpoint.md) (suplementos contextuais de email) com um valor entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão **CustomPane**
