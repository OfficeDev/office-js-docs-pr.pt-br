---
title: Conjuntos de requisitos comuns da API
description: Especifica quais plataformas Office e builds suportam as APIs dinâmicas da faixa de opções.
ms.date: 05/12/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: a608eff12fb21d7a4a6beb195749141bd473aa1c
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330182"
---
# <a name="ribbon-api-requirement-sets"></a>Conjuntos de requisitos comuns da API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

O conjunto de APIs da Faixa de Opções oferece suporte ao controle programático de quando comandos personalizados de complemento (ou seja, botões de faixa de opções personalizados e itens de menu) estão habilitados e desabilitados.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API da Faixa de Opções, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

|  Conjunto de requisitos  | Office 2013 no Windows<br>(compra avulsa) | Office 2016 ou posterior no Windows<br>(compra avulsa)   | Office no Windows\*<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac\*<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web\*  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | Não disponível | Não disponível | Consulte suporte<br>seção abaixo | Não disponível | 16.38 | Novembro de 2020 | Não disponível|
| RibbonApi 1.2  | Não disponível | Não disponível | 2102 (build 13801.20294) | Não disponível | em breve | Maio de 2021 | Não disponível|

> **&#42;** A API da Faixa de Opções é suportada somente Excel e requer uma assinatura Microsoft 365 de opções.

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>Suporte para a versão 1.1 no Office no Windows (assinatura)

A versão 1.1 do conjunto de requisitos RibbonApi é suportada no Canal do Consumidor versão 2006 (build 13001.20498 ou superior). Para Office no Windows o recurso também é suportado nas builds do Canal Semi-Annual e do Canal mensal Enterprise disponíveis em 14 de julho de 2020 ou posterior. As builds mínimas com suporte para cada canal são as seguinte:  

|Canal | Versão | Build|
|:-----|:-----|:-----|
|Canal Atual | 2006 ou superior | 20266.20266 ou superior|
|Canal Empresarial Mensal | 2005 ou superior | 12827.20538 ou superior|
|Canal Empresarial Mensal | 2004 | 12730.20602 ou superior|
|Canal Empresarial Semestral | 2002 ou superior | 12527.20880 ou superior|

## <a name="more-information"></a>Mais informações

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e com build de versões de canal de atualização para Microsoft 365 clientes](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número de com build de um aplicativo Microsoft 365 cliente](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API da Faixa de Opções 1.1

A API da Faixa de Opções 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o [tópico Office.ribbon](/javascript/api/office/office.ribbon) reference.

## <a name="ribbon-api-12"></a>API da Faixa de Opções 1.2

A API da Faixa de Opções 1.2 adiciona suporte a guias contextuais. Para obter mais informações, confira [Criar guias contextuais personalizadas em Suplementos do Office](../../design/contextual-tabs.md).

> [!NOTE]
> O conjunto de **requisitos RibbonApi 1.2** ainda não tem suporte no manifesto, portanto, você não deve especificá-lo na seção do `<Requirements>` manifesto.

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
