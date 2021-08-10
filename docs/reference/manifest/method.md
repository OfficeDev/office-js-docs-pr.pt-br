---
title: Elemento Method no arquivo de manifesto
description: O elemento Method especifica um método individual da API javaScript Office que os seus Office Desempresos exigem para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 811cd84e1ad2aade8b7042eefa822eee6b2ab200a8fa1b71c9fe5fc34874ec66
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089723"
---
# <a name="method-element"></a>Elemento Method

Especifica um método individual da API javaScript Office que seu Office Descrição requer para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

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

Os `Methods` elementos e não são `Method` suportados por complementos de email. Para obter mais informações sobre conjuntos de requisitos, [consulte Office versões e conjuntos de requisitos.](../../develop/office-versions-and-requirement-sets.md)

> [!IMPORTANT]
> Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento. Para obter mais informações sobre como fazer isso, consulte [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
