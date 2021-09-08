---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto Office de requisitos da API JavaScript que seu Office de complemento requer para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936619"
---
# <a name="set-element"></a>Elemento Set

Especifica um conjunto de requisitos da API JavaScript Office que o seu Office Descrição requer para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

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

Para obter mais informações sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets,** consulte Definir o elemento [Requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT]
> Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível. Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas). Além disso, você não pode declarar suporte para métodos específicos nos suplementos de email.
