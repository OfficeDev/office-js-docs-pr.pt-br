---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596785"
---
# <a name="appdomain-element"></a>Elemento AppDomain

Especifica domínios adicionais que carregam páginas na janela do suplemento. Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).
> 2. *Não* Coloque uma barra de fechamento, "/", no valor.

## <a name="contained-in"></a>Contido em

[AppDomains](appdomains.md)

## <a name="remarks"></a>Comentários

Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md). Confira mais informações em [Manifesto XML de Suplementos do Office](../../develop/add-in-manifests.md).
