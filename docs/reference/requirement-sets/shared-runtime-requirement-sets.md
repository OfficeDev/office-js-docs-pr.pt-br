---
title: Conjuntos de requisitos de tempo de execução compartilhado
description: Especifica as plataformas e hosts do Office que dão suporte às APIs SharedRuntime.
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 9750dd2e20a9f2426c7faf3864a2fccac5c11a80
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600673"
---
# <a name="shared-runtime-requirement-sets"></a>Conjuntos de requisitos de tempo de execução compartilhado

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Partes de um suplemento do Office que executam código JavaScript, como painéis de tarefas, arquivos de função iniciados a partir de comandos de suplemento e funções personalizadas do Excel, podem compartilhar um único tempo de execução do JavaScript. Isso permite que todas as partes compartilhem um conjunto de variáveis globais, compartilhem um conjunto de bibliotecas carregadas e se comuniquem entre si sem precisar passar mensagens por meio de um armazenamento persistente.

A tabela a seguir lista o conjunto de requisitos SharedRuntime 1,1, os aplicativos host do Office que dão suporte a esse conjunto de requisitos e os números de versão ou de compilação do aplicativo do Office.

|  Conjunto de requisitos  |  Office 2013 (ou posterior) no Windows<br>(compra avulsa) | Office no Windows<br>(conectado à assinatura do Office 365)   |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  | Servidor do Office Online |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1,1  | Não disponível | Versão 2002 (Build 12527,20092) ou posterior | Não disponível | 16,35 ou posterior | Fevereiro de 2020 | Não disponível |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
