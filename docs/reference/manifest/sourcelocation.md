---
title: Elemento SourceLocation no arquivo de manifesto
description: O elemento SourceLocation especifica os locais do arquivo de origem para o suplemento do Office.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: fcca051b0d85c98cb011d5b886981c543ef8e3b0
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717898"
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
- [ExtensionPoint](extensionpoint.md) (suplementos contextuais de email)

## <a name="can-contain"></a>Pode conter

[Override](override.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obrigatório|Especifica o valor padrão para essa configuração para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).|
