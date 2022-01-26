---
title: Elemento Method no arquivo de manifesto
description: O elemento Method especifica um método individual da API JavaScript Office que seus Office Add-ins exigem para serem ativados pelo Office ou substituir as configurações de manifesto base.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 052fb41a7077781843ea7e63d9601a819058dfa6
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222266"
---
# <a name="method-element"></a>Elemento Method

O significado desse elemento depende de onde ele é usado no manifesto.

## <a name="in-the-base-manifest"></a>No manifesto base

Quando usado no manifesto base (ou seja, o elemento Requisitos do vôrent  é um filho direto do [OfficeApp](officeapp.md)), o elemento **Method** especifica um método individual da API JavaScript do Office que seu add-in do Office precisa para ser ativado por Office.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Como bisneto de um elemento VersionOverrides

Especifica um método individual da API javaScript Office que deve ser suportada pela versão e plataforma do Office (como Windows, Mac, Web e iOS ou iPad) para que [o VersionOverrides](versionoverrides.md) entre em vigor.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

- O mesmo que o elemento [Requisitos](requirements.md) do vôver.

**Associado a esses conjuntos de requisitos:**

- O mesmo que o elemento [Requisitos](requirements.md) do vôver.

## <a name="syntax"></a>Sintaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contido em

[Methods](methods.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Nome|cadeia de caracteres|obrigatório|Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o `getSelectedDataAsync` método, você deve especificar `"Document.getSelectedDataAsync"` .|

## <a name="remarks"></a>Comentários

Os **elementos Métodos** e **Métodos** não são suportados por complementos de email quando usados no manifesto base. Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

> [!IMPORTANT]
> Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento. Para obter mais informações sobre como fazer isso, consulte [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
