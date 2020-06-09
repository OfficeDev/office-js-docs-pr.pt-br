---
title: Elemento RequestedHeight no arquivo de manifesto
description: O elemento RequestedHeight especifica a altura inicial (em pixels) de um suplemento de conteúdo ou email.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 44675918a4208683f442fe8a6e8f4f906f484571
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611726"
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
- [ExtensionPoint](extensionpoint.md) (suplementos de email contextuais) com um valor que pode ser entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o [ponto de extensão **CustomPane** (preterido)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)
