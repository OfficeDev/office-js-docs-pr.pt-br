---
title: Conjuntos de requisitos da Dialog API
description: Saiba mais sobre os conjuntos de requisitos de API da caixa de diálogo.
ms.date: 09/14/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 79b6960387519ac3c8b41b0b31cf6f40b5e7e067
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771358"
---
# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos da API de Caixa de diálogo

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de caixa de diálogo, os aplicativos cliente do Office que dão suporte a esse conjunto de requisitos e os números de versão ou de compilação do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 no Windows\*<br>(compra avulsa) | Office 2016 ou posterior no Windows\*<br>(compra avulsa)   | Office no Windows<br>scriçõe |  Office no iPad<br>scriçõe  |  Office no Mac<br>scriçõe  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1,2  | N/D | N/D | Consulte suporte<br>seção abaixo | 2,67 ou posterior | 16,37 ou posterior | Junho de 2020 | N/D |
| DialogApi 1.1  | Build 15.0.4855.1000 ou posterior | Build 16.0.4390.1000 ou posterior | Versão 1602 (build 6741.0000) ou posterior | 1.22 ou posterior | 15.20 ou posterior | Janeiro de 2017 | Versão 1608 (build 7601.6800) ou posterior|

>\* Os usuários do Office de compra única podem não ter aceitado todos os patches e atualizações. Em caso afirmativo, a DLL que o Office usa para relatar sua versão na interface do usuário pode ser maior do que as versões listadas aqui, mesmo se as DLLs atualizadas necessárias para dar suporte ao DialogApi não estiverem instaladas no computador do usuário. Para garantir que o patch necessário está instalado, o usuário deve ir para a lista atualização do Office ([lista](/officeupdates/msp-files-office-2013) do Office 2013 ou [lista do Office 2016](/officeupdates/msp-files-office-2016)), procurar **osfclient-x-None** e instalar o patch listado.

## <a name="office-on-windows-subscription-support"></a>Suporte do Office no Windows (assinatura)

O conjunto de requisitos DialogApi 1,2 é suportado no canal de consumidor versão 2005 (Build, 12827,20268 ou posterior). Para o Office no Windows, o recurso também é suportado no canal Semi-Annual e nas compilações de canal corporativo mensal, 9 de junho de 2020 ou mais recente. As compilações mínimas suportadas para cada canal são as seguintes:  

|Canal | Versão | Build|
|:-----|:-----|:-----|
|Canal Atual | 2005 ou maior | 12827,20160 ou maior|
|Canal Empresarial Mensal | 2004 ou maior | 12730,20430 ou maior|
|Canal Empresarial Semestral | 2002 ou maior | 12527,20720 ou maior|

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11-and-12"></a>API de diálogo 1,1 e 1,2

O Dialog API 1.1 é a primeira versão da API. O conjunto de requisitos 1,2 adiciona suporte ao envio de dados da página pai para a caixa de diálogo com o método [Office. Dialog. messageChild](/javascript/api/office/office.dialog#messageChild_message_) . Para obter detalhes sobre essas APIs, consulte o tópico de referência da [API da caixa de diálogo](/javascript/api/office/office.ui) .

## <a name="see-also"></a>Confira também

- [Usar a API de diálogo do Office em suplementos do Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
