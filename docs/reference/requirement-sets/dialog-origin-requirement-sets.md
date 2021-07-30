---
title: Conjuntos de requisitos de origem da caixa de diálogo
description: Saiba mais sobre os conjuntos de requisitos de Origem da Caixa de Diálogo.
ms.date: 07/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 1ec5949c689021f080491a19aea4661627b2d95c
ms.sourcegitcommit: f46e4aeb9c31f674380dd804fd72957998b3a532
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/23/2021
ms.locfileid: "53536057"
---
# <a name="dialog-origin-requirement-sets"></a>Conjuntos de requisitos de origem da caixa de diálogo

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de Origem da Caixa de Diálogo, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

|  Conjunto de requisitos  | Office 2013 no Windows\*<br>(compra avulsa) | Office 2016 no Windows\*<br>(compra avulsa) | Office 2019 ou posterior no Windows\*<br>(compra avulsa) | Office no Windows<br>(assinatura) |  Office no iPad<br>(assinatura)  |  Office no Mac<br>(assinatura)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Compilar<br>15.0.5371.1000<br>ou posterior | Compilar<br>16.0.5200.1000<br>ou posterior | Compilar<br>A definir<br>ou posterior | A definir | 2,52 ou posterior | 16,52 ou posterior | Julho de 2021 | Versão 2108<br>(Build 10377.1000)<br>ou posterior |

>\*Os usuários da compra única Office podem não ter aceito todos os patches e atualizações. Em caso afirmativo, a DLL que Office usa para relatar sua versão na interface do usuário pode ser maior do que as versões listadas aqui, mesmo que as DLLs atualizadas necessárias para dar suporte a DialogApi não tenham sido instaladas no computador do usuário. Para garantir que o patch necessário seja instalado, o usuário deve ir para a lista de atualizações do Office ( lista [Office 2013](/officeupdates/msp-files-office-2013) ou [lista Office 2016](/officeupdates/msp-files-office-2016)), pesquisar **osfclient-x-none** e instalar o patch listado.

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