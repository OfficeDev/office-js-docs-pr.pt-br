---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641414"
---
# <a name="set-element"></a>Elemento Set

Especifica um conjunto de requisitos da API JavaScript do Office que o suplemento do Office exige para ativar.

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
|MinVersion|cadeia de caracteres|opcional|Especifica a versão mínima do conjunto de APIs exigido pelo seu suplemento. Substitui o valor de **DefaultMinVersion**, se estiver especificado no elemento [sets](sets.md) pai.|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT]
> Para suplementos de email, há apenas um conjunto de requisitos `"Mailbox"` disponível. Esse conjunto de requisitos contém o subconjunto completo da API compatível com os suplementos de email do Outlook. Você deve especificar o conjunto de requisitos de `"Mailbox"` no manifesto de seu suplemento de email (não é opcional como no caso de suplementos de conteúdo e do painel de tarefas). Além disso, você não pode declarar suporte para métodos específicos nos suplementos de email.
