---
title: Elemento Set no arquivo de manifesto
description: O elemento Set especifica um conjunto Office de requisitos da API JavaScript que seu Office de complemento requer para ativar.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 93524d64fd915d6f42f4e4a0cd0ab6cc3335f4ce
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152002"
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
