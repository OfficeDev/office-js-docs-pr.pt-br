---
title: Conjuntos de requisitos da Dialog API
description: Saiba mais sobre os conjuntos de requisitos de API da caixa de diálogo.
ms.date: 06/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: aa591a1b37c94a4db621d19786857303bb6ac473
ms.sourcegitcommit: 449a728118db88dea22a44f83728d21604d6ee8c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/12/2020
ms.locfileid: "44719060"
---
# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos da Dialog API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da Dialog API, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  | Office 2013 no Windows\*<br>(compra avulsa) | Office 2016 ou posterior no Windows\*<br>(compra avulsa)   | Office no Windows<br>(conectado à assinatura do Office 365) |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 ou posterior | Build 16.0.4390.1000 ou posterior | Versão 1602 (build 6741.0000) ou posterior | 1.22 ou posterior | 15.20 ou posterior| Janeiro de 2017 | Versão 1608 (build 7601.6800) ou posterior|

>\*Os usuários do Office de compra única podem não ter aceitado todos os patches e atualizações. Em caso afirmativo, a DLL que o Office usa para relatar sua versão na interface do usuário pode ser maior do que as versões listadas aqui, mesmo se as DLLs atualizadas necessárias para dar suporte ao DialogApi não estiverem instaladas no computador do usuário. Para garantir que o patch necessário está instalado, o usuário deve ir para a lista atualização do Office ([lista](/officeupdates/msp-files-office-2013) do Office 2013 ou [lista do Office 2016](/officeupdates/msp-files-office-2016)), procurar **osfclient-x-None**e instalar o patch listado.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>Dialog API 1.1

O Dialog API 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência da [API da caixa de diálogo](/javascript/api/office/office.ui) .

## <a name="see-also"></a>Confira também

- [Usar a API de diálogo do Office em suplementos do Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
