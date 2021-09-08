---
title: Conjuntos de requisitos de origem da caixa de diálogo
description: Saiba mais sobre os conjuntos de requisitos de Origem da Caixa de Diálogo.
ms.date: 07/22/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 24513823eb60435359d5d7307a11a192fece2015
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938900"
---
# <a name="dialog-origin-requirement-sets"></a>Conjuntos de requisitos de origem da caixa de diálogo

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de Origem da Caixa de Diálogo, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

|  Conjunto de requisitos  | Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(assinatura) |  Office no iPad<br>(assinatura)  |  Office no Mac<br>(assinatura)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Compilar<br>15.0.5371.1000<br>ou posterior | Compilar<br>16.0.5200.1000<br>ou posterior | Compilar<br>A definir<br>ou posterior | A definir | 2,52 ou posterior | 16,52 ou posterior | Julho de 2021 | Versão 2108<br>(Build 10377.1000)<br>ou posterior |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-origin-11"></a>Origem da caixa de diálogo 1.1

A Origem da Caixa de Diálogo 1.1 é a primeira versão da API. Ele fornece suporte para mensagens entre domínios entre uma caixa de diálogo e sua página pai. Para obter detalhes sobre essas APIs, consulte o [tópico Office.ui](/javascript/api/office/office.ui) reference.

## <a name="see-also"></a>Confira também

- [Usar a API de diálogo do Office em suplementos do Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
