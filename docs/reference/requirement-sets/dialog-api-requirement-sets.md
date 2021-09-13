---
title: Conjuntos de requisitos da Dialog API
description: Saiba mais sobre os conjuntos de requisitos da API de Caixa de Diálogo.
ms.date: 07/19/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 7fdef827cf47903b0b7e2872110a5a6801735bf4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151965"
---
# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos da API de Caixa de diálogo

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de Caixa de Diálogo, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo Office.

|  Conjunto de requisitos  | Office 2013 no Windows\*<br>(compra avulsa) | Office 2016 ou posterior no Windows\*<br>(compra avulsa)   | Office no Windows<br>(assinatura) |  Office no iPad<br>(assinatura)  |  Office no Mac<br>(assinatura)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | N/D | N/D | Consulte suporte<br>seção abaixo | 2.37 ou posterior | 16.37 ou posterior | Junho de 2020 | N/A |
| DialogApi 1.1  | Build 15.0.4855.1000 ou posterior | Build 16.0.4390.1000 ou posterior | Versão 1602 (build 6741.0000) ou posterior | 1.22 ou posterior | 15.20 ou posterior | Janeiro de 2017 | Versão 1608 (build 7601.6800) ou posterior|

>\*Os usuários da compra única Office podem não ter aceito todos os patches e atualizações. Em caso afirmativo, a DLL que Office usa para relatar sua versão na interface do usuário pode ser maior do que as versões listadas aqui, mesmo que as DLLs atualizadas necessárias para dar suporte a DialogApi não tenham sido instaladas no computador do usuário. Para garantir que o patch necessário seja instalado, o usuário deve ir para a lista de atualizações do Office ( lista [Office 2013](/officeupdates/msp-files-office-2013) ou [lista Office 2016](/officeupdates/msp-files-office-2016)), pesquisar **osfclient-x-none** e instalar o patch listado.

## <a name="office-on-windows-subscription-support"></a>Office suporte Windows (assinatura)

O conjunto de requisitos DialogApi 1.2 é suportado no Canal do Consumidor versão 2005 (build, 12827.20268 ou superior). Para Office no Windows, o recurso também é suportado nas builds do Canal Semi-Annual e do Canal mensal Enterprise disponíveis em 9 de junho de 2020 ou posterior. As builds mínimas com suporte para cada canal são as seguinte:  

|Canal | Versão | Build|
|:-----|:-----|:-----|
|Canal Atual | 2005 ou superior | 12827.20160 ou superior|
|Canal Empresarial Mensal | 2004 ou superior | 12730.20430 ou superior|
|Canal Empresarial Semestral | 2002 ou superior | 12527.20720 ou superior|

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11-and-12"></a>API de caixa de diálogo 1.1 e 1.2

O Dialog API 1.1 é a primeira versão da API. O conjunto de requisitos 1.2 adiciona suporte para o envio de dados da página pai à caixa de diálogo com o [método Office.dialog.messageChild.](/javascript/api/office/office.dialog#messageChild_message_) Para obter detalhes sobre essas APIs, consulte o [tópico de referência da API](/javascript/api/office/office.ui) de Diálogo.

## <a name="see-also"></a>Confira também

- [Usar a API de diálogo do Office em suplementos do Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
