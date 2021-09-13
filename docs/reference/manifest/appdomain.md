---
title: Elemento AppDomain no arquivo de manifesto
description: Especifica domínios adicionais que são usados pelo seu complemento e devem ser confiáveis por Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: c17195e6d9d3f4f22465c8aa1fc626afd3eb06c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151677"
---
# <a name="appdomain-element"></a>Elemento AppDomain

Especifica um domínio adicional que Office deve confiar, além do especificado no [elemento SourceLocation](sourcelocation.md). A especificação de um domínio tem esses efeitos:

- Ele permite que páginas, rotas ou outros recursos no domínio sejam abertos diretamente no painel de tarefas raiz do add-in em plataformas Office desktop. (Especificar um domínio em um **AppDomain** não é necessário para Office na Web ou para abrir um recurso em um IFrame, nem é necessário para abrir um recurso em uma caixa de diálogo aberta com a [API](../../develop/dialog-api-in-office-add-ins.md)de Diálogo .)
- Ele permite que as páginas no domínio façam chamadas Office.js API de IFrames dentro do complemento.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. Se houver uma porta explícita para o domínio, inclua-a (por exemplo, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).
> 3. Se um subdomínio precisar ser confiável, inclua-o (por exemplo, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ). O subdomínio `mysubdomain.mydomain.com` e `mydomain.com` são domínios diferentes. Se ambos precisam ser confiáveis, ambos precisam estar em elementos **AppDomain** separados.
> 4. Listar o mesmo domínio especificado no [elemento SourceLocation](sourcelocation.md) não tem efeito e pode ser enganoso. Em particular, quando você está desenvolvendo em , você não precisa criar um `localhost` **elemento AppDomain** para `localhost` .
> 5. Não inclua segmentos de uma URL além do domínio. Por exemplo, não inclua a URL completa de uma página.
> 6. Não *coloque* uma barra de fechamento , "/", no valor.

## <a name="contained-in"></a>Contido em

[AppDomains](appdomains.md)

## <a name="remarks"></a>Comentários

Para saber mais, confira o [manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md).
