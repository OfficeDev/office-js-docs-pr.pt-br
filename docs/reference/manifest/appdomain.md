---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433065"
---
# <a name="appdomain-element"></a>Elemento AppDomain

Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Contido em

[AppDomains](appdomains.md)

## <a name="remarks"></a>Comentários

Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md). Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).
