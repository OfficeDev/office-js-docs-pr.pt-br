---
title: Elemento Version no arquivo de manifesto
description: O elemento Version especifica seu Office versão do add-in.
ms.date: 02/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 34cefa22123ed4ee723d51a669e01e042efc2934
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152075"
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
