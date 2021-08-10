---
title: Elemento RequestedHeight no arquivo de manifesto
description: O elemento RequestedHeight especifica a altura inicial (em pixels) de um conteúdo ou de um complemento de email.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 0e5f9de909d32622ac244ff4118c8a3192abf2ff0fe89ed81a6188ddcb265549
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092980"
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
- [ExtensionPoint](extensionpoint.md) (complementos de email contextuais) com um valor que pode estar entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão [ **CustomPane** (preterido)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)
