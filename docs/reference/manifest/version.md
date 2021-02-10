---
title: Elemento Version no arquivo de manifesto
description: O elemento Version especifica a versão do seu Complemento do Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173931"
---
# <a name="version-element"></a>Elemento Version

Especifica a versão de seu Complemento do Office. O número da versão pode ser 1, 2, 3 ou 4 partes (ou seja, n, n.n, n.n.n ou n.n.n.n).

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

Cada parte do número da versão pode ter no máximo 5 dígitos.
