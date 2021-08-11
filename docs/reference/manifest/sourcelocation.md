---
title: Elemento SourceLocation no arquivo de manifesto
description: O elemento SourceLocation especifica os locais do arquivo de origem para o Office Do-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 6830a26cf192802c97c486511695b4ace35ac8263cfcd30ceaf71398f0d83a07
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095002"
---
# <a name="sourcelocation-element"></a>Elemento SourceLocation

Especifica os locais de arquivo de origem do seu Office como uma URL entre 1 e 2018 caracteres. O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contido em

- [DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)
- [FormSettings](formsettings.md) (suplementos de email)
- [ExtensionPoint](extensionpoint.md) (Contextual e LaunchEvent mail add-ins)

## <a name="can-contain"></a>Pode conter

[Override](override.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obrigatório|Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|
