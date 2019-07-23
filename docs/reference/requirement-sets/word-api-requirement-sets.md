---
title: Conjuntos de requisitos da API JavaScript do Word
description: O requisito do suplemento do Office define informações para os builds do Word
ms.date: 07/17/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 4af7a9c14489d148ffdc06a68ad6c26bf326abc5
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804622"
---
# <a name="word-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Word

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="requirement-set-availability"></a>Disponibilidade do conjunto de requisitos

Os suplementos do Word são executados em várias versões do Office, incluindo o Office 2016 ou posterior no Windows, Office na Web, iPad e Mac. A tabela a seguir lista os conjuntos de requisitos do Word, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build desses aplicativos.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados, faça referência à biblioteca **production** no CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Para obter informações sobre o uso de APIs de visualização, confira o artigo [APIs de visualização do JavaScript para Excel](word-preview-apis.md).

|  Conjunto de requisitos  |   Office no Windows\*<br>(conectado à assinatura do Office 365)  |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  |
|:-----|-----|:-----|:-----|:-----|
| [Visualização](word-preview-apis.md) | Use a versão mais recente do Office para testar as APIs de visualização (talvez seja exigido ser membro do [programa Office Insider](https://products.office.com/office-insider)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Versão 1612 (Compilação 7668.1000) ou posterior| Março de 2017, 2.22 ou posterior | Março de 2017, 15.32 ou posterior| Março de 2017 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Atualização de dezembro de 2015, versão 1601 (build 6568.1000) ou posterior | Janeiro de 2016, 1.18 ou posterior | Janeiro de 2016, 15.19 ou posterior| Setembro de 2016 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Versão 1509 (build 4266.1001) ou posterior| Janeiro de 2016, 1.18 ou posterior | Janeiro de 2016, 15.19 ou posterior| Setembro de 2016 |

> [!NOTE]
> O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém o conjunto de requisitos 1.1 de WordApi.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do Word](/javascript/api/word)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
