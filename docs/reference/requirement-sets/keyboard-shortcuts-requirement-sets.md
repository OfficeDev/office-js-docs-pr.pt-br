---
title: Conjuntos de requisitos de Atalhos de Teclado
description: Informações de conjunto de requisitos de atalhos de teclado para Office de complementos.
ms.date: 11/22/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 209cc46c37ac004422796e267a8c350e33ffc615
ms.sourcegitcommit: b3ddc1ddf7ee810e6470a1ea3a71efd1748233c9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/24/2021
ms.locfileid: "61153791"
---
# <a name="keyboard-shortcuts-requirement-sets"></a>Conjuntos de requisitos de Atalhos de Teclado

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de Atalhos de Teclado, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo Office.

|  Conjunto de requisitos  | Office 2013 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1  | N/D | Versão: 2111 (build 14701.10000) | N/D | 16.55 | Setembro de 2021 |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="keyboardshortcuts-11"></a>KeyboardShortcuts 1.1

Para obter detalhes sobre as APIs neste conjunto de requisitos, [consulte Office.actions](/javascript/api/office/office.actions).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
