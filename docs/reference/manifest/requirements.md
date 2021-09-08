---
title: Elemento Requirements no arquivo de manifesto
description: O elemento Requirements especifica o conjunto de requisitos mínimo e os métodos que seu Office de complemento precisa para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937328"
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

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Métodos](methods.md)|x||x|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).
