---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337192"
---
# <a name="appdomain-element"></a>Elemento AppDomain

Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.

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

Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md). Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).
