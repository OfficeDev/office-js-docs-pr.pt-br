---
title: Elemento Requirements no arquivo de manifesto
description: O elemento Requirements especifica o conjunto de requisitos mínimo e os métodos que seu Office de complemento precisa para ativar.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 3a5a393485094b5cc830b5120c3abd8c211eff1e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151838"
---
# <a name="requirements-element"></a>Elemento Requirements

Especifica o conjunto mínimo de Office da API JavaScript[(](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) conjuntos de requisitos e/ou métodos) que seu Office Add-in precisa ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Métodos](methods.md)|x||x|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).
