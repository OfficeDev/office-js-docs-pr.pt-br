---
title: Elemento Methods no arquivo de manifesto
description: O elemento Methods especifica Office lista de métodos de API JavaScript que seu Office Add-in exige para ser ativado pelo Office ou para substituir as configurações de manifesto base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c39c6363cd33e103cf40c0f7f047fa694db1411
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222273"
---
# <a name="methods-element"></a>Elemento Methods

O significado desse elemento depende de onde ele é usado no manifesto.

## <a name="in-the-base-manifest"></a>No manifesto base

Quando usado no manifesto base (ou seja, o elemento **Requirements** pai é um filho direto do [OfficeApp](officeapp.md)), o elemento **Methods** especifica a lista de métodos da API JavaScript do Office que seu Office Add-in precisa para ser ativado por Office.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>Como um neto de um elemento VersionOverrides

Especifica o conjunto mínimo de métodos de API JavaScript Office que devem ser suportados pela versão e plataforma do Office (como Windows, Mac, Web e iOS ou iPad) para que [o VersionOverrides](versionoverrides.md) entre em vigor.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

- O mesmo que o elemento [Pai Requirements.](requirements.md)

**Associado a esses conjuntos de requisitos:**

- O mesmo que o elemento [Pai Requirements.](requirements.md)

## <a name="syntax"></a>Sintaxe

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a>Contido em

[Requisitos](requirements.md)

## <a name="can-contain"></a>Pode conter

[Method](method.md)

## <a name="remarks"></a>Comentários

Os **elementos Métodos** e **Métodos** não são suportados em complementos de email quando usados no manifesto base. Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).
