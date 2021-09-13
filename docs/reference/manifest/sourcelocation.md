---
title: Elemento SourceLocation no arquivo de manifesto
description: O elemento SourceLocation especifica os locais do arquivo de origem para o Office Do-in.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1b30227beee7deceb019b5970f2bb7b6f233dd56
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151994"
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
