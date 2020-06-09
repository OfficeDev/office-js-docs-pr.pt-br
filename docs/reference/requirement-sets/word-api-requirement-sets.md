---
title: Conjuntos de requisitos da API JavaScript do Word
description: Informações do conjunto de requisitos do Suplemento do Office para builds do Word.
ms.date: 04/16/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: bffd78455cd6d87a1323c4133ce16f9723e37a4c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611279"
---
# <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Word

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Disponibilidade do conjunto de requisitos

Os suplementos do Word são executados em várias versões do Office, incluindo o Office 2016 ou posterior no Windows, Office na Web, iPad e Mac. A tabela a seguir lista os conjuntos de requisitos do Word, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build desses aplicativos.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados, faça referência à biblioteca **production** no CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Para obter informações sobre o uso de APIs de visualização, confira o artigo [APIs de visualização do JavaScript para Excel](word-preview-apis.md).

|  Conjunto de requisitos  |   Office no Windows\*<br>(conectado à assinatura do Office 365)  |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  |
|:-----|-----|:-----|:-----|:-----|
| [Visualização](word-preview-apis.md) | Use a versão mais recente do Office para testar as APIs de visualização (talvez seja exigido ser membro do [programa Office Insider](https://insider.office.com)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Versão 1612 (Compilação 7668.1000) ou posterior| Março de 2017, 2.22 ou posterior | Março de 2017, 15.32 ou posterior| Março de 2017 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Atualização de dezembro de 2015, versão 1601 (build 6568.1000) ou posterior | Janeiro de 2016, 1.18 ou posterior | Janeiro de 2016, 15.19 ou posterior| Setembro de 2016 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Versão 1509 (build 4266.1001) ou posterior| Janeiro de 2016, 1.18 ou posterior | Janeiro de 2016, 15.19 ou posterior| Setembro de 2016 |

> [!NOTE]
> Versões permanentes dos conjuntos de requisitos de suporte do Office como a seguir:
>
> - O Office 2019 é compatível com o WordApi 1.3 e versões anteriores.
> - O Office 2016 é compatível somente com o conjunto de requisitos do WordApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>Também consulte

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
