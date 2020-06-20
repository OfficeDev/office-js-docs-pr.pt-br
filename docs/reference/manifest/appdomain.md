---
title: Elemento AppDomain no arquivo de manifesto
description: Especifica domínios adicionais que são usados pelo seu suplemento e que deve ser confiável para o Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778645"
---
# <a name="appdomain-element"></a>Elemento AppDomain

Especifica um domínio adicional no qual o Office deve confiar, além do especificado no [elemento SourceLocation](sourcelocation.md). A especificação de um domínio tem estes efeitos:

- Ele permite que páginas, rotas ou outros recursos no domínio sejam abertos diretamente no painel de tarefas raiz do suplemento em plataformas do Office. (Especificar um domínio em um **AppDomain** não é necessário para o Office na Web ou para abrir um recurso em um iframe, nem é necessário para abrir um recurso em uma caixa de diálogo aberta com a [API da caixa de diálogo](../../develop/dialog-api-in-office-add-ins.md).)
- Ele permite que as páginas no domínio façam chamadas de API Office.js de IFrames no suplemento.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. Se houver uma porta explícita para o domínio, inclua-a (por exemplo, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).
> 3. Se um subdomínio precisar ser confiável, inclua-o (por exemplo, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ). O subdomínio `mysubdomain.mydomain.com` e os `mydomain.com` domínios são diferentes. Se ambos precisam ser confiáveis, então ambos precisam estar em elementos **AppDomain** separados.
> 4. A listagem do mesmo domínio que o especificado no [elemento SourceLocation](sourcelocation.md) não tem efeito e pode ser enganosa. Em particular, quando você está desenvolvendo `localhost` , não é necessário criar um elemento **AppDomain** para `localhost` .
> 5. Não inclua nenhum segmento de uma URL além do domínio. Por exemplo, não inclua a URL completa de uma página.
> 6. *Não* Coloque uma barra de fechamento, "/", no valor.

## <a name="contained-in"></a>Contido em

[AppDomains](appdomains.md)

## <a name="remarks"></a>Comentários

Para saber mais, confira o [manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md).
