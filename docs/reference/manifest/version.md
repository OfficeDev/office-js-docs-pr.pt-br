---
title: Elemento Version no arquivo de manifesto
description: O elemento Version especifica seu Office versão do add-in.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 9641153cbe6fa0284986b8dd286ba2114b32a82894bd5f8d33516e2a56c90be9
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096324"
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
