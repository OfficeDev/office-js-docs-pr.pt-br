---
title: Conjuntos de requisitos da Dialog API
description: ''
ms.date: 07/05/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: a524edf6734618a56e050d2c25eedbd23ca13973
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617014"
---
# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos da Dialog API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da Dialog API, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  | Office 2013 no Windows\*<br>(compra avulsa) | Office 2016 ou posterior no Windows\*<br>(compra avulsa)   | Office no Windows<br>(conectado à assinatura do Office 365) |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 ou posterior | Build 16.0.4390.1000 ou posterior | Versão 1602 (build 6741.0000) ou posterior | 1.22 ou posterior | 15.20 ou posterior| Janeiro de 2017 | Versão 1608 (build 7601.6800) ou posterior|

>\*Os usuários do Office de compra única podem não ter aceitado todos os patches e atualizações. Em caso afirmativo, a DLL que o Office usa para relatar sua versão na interface do usuário pode ser maior do que as versões listadas aqui, mesmo se as DLLs atualizadas necessárias para dar suporte ao DialogApi não estiverem instaladas no computador do usuário. Para garantir que o patch necessário está instalado, o usuário deve ir para a lista atualização do Office ([lista](/officeupdates/msp-files-office-2013) do Office 2013 ou [lista do Office 2016](/officeupdates/msp-files-office-2016)), procurar **osfclient-x-None**e instalar o patch listado. 

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>Dialog API 1.1

O Dialog API 1.1 é a primeira versão da API. Para saber mais sobre a API, confira o tópico de referência [Dialog API](/javascript/api/office/office.ui).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
