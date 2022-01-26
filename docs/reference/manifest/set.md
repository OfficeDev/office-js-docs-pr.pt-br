---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto de requisitos da API JavaScript Office que seu Office Add-in requer para ser ativado por Office ou para substituir as configurações de manifesto base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55e1b25765bfbe53108bc9201c0c851c6ef9161d
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222231"
---
# <a name="set-element"></a>Elemento Set

O significado desse elemento depende de onde ele é usado no manifesto.

## <a name="in-the-base-manifest"></a>No manifesto base

Quando usado no manifesto base (ou seja, o elemento Requisitos do vôrent [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) é um filho direto do [OfficeApp](officeapp.md)), o elemento **Set** especifica um conjunto de requisitos da API JavaScript do Office que seu add-in do Office precisa para ser ativado por Office. 

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Como bisneto de um elemento VersionOverrides

Especifica um [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) conjunto de requisitos da API JavaScript do Office que deve ser suportado pela versão e plataforma do Office (como Windows, Mac, Web e iOS ou iPad) para que [o VersionOverrides](versionoverrides.md) entre em vigor.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

- O mesmo que o elemento [Requisitos](requirements.md) do vôver.

**Associado a esses conjuntos de requisitos:**

- O mesmo que o elemento [Requisitos](requirements.md) do vôver.

## <a name="syntax"></a>Sintaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contido em

[Sets](sets.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Nome|cadeia de caracteres|obrigatório|O nome de um [conjunto de requisitos](../../develop/office-versions-and-requirement-sets.md).|
|MinVersion|cadeia de caracteres|opcional|Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se for especificado no elemento [Sets](sets.md) pai.|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Para obter mais informações sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets,** consulte Specify which Office versions and platforms can host [your add-in](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in).

