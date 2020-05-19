---
title: Elemento SourceLocation no arquivo de manifesto
description: O elemento SourceLocation especifica os locais do arquivo de origem para o suplemento do Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 642780c3231523ea579ca548b3f3f984b2856666
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278396"
---
# <a name="sourcelocation-element"></a>Elemento SourceLocation

Especifica os locais do arquivo de origem para o suplemento do Office como uma URL entre 1 e 2018 caracteres de comprimento. O local de origem deve ser um endereço HTTPS, não um caminho de arquivo.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contido em

- [DefaultSettings](defaultsettings.md) (suplementos de conteúdo e de painel de tarefas)
- [FormSettings](formsettings.md) (suplementos de email)
- [ExtensionPoint](extensionpoint.md) (contextuais e LaunchEvent (Visualizar) suplementos de email)

## <a name="can-contain"></a>Pode conter

[Override](override.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obrigatório|Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|
