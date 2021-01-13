---
title: Conjuntos de requisitos comuns da API
description: Especifica quais plataformas e builds do Office suportam as APIs dinâmicas da faixa de opções.
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 91c909755779d122fba8d77dc246784f6a0dd1a3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839982"
---
# <a name="ribbon-api-requirement-sets"></a>Conjuntos de requisitos comuns da API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

O conjunto de APIs da Faixa de Opções dá suporte ao controle programático de quando os comandos personalizados do add-in (ou seja, botões da faixa de opções personalizados e itens de menu) estão habilitados e desabilitados.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API da Faixa de Opções, os aplicativos cliente do Office que suportam esse conjunto de requisitos e os números de versão ou com build do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 no Windows<br>(compra avulsa) | Office 2016 ou posterior no Windows<br>(compra avulsa)   | Office no Windows\*<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac\*<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web\*  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/D | N/D | Consulte o suporte<br>seção abaixo | N/D | 16.38 | Novembro de 2020 | N/D|

> **&#42;** A API da Faixa de Opções tem suporte apenas no Excel e requer a assinatura do Microsoft 365.

## <a name="office-on-windows-subscription-support"></a>Suporte do Office no Windows (assinatura)

O conjunto de requisitos é suportado no Canal para Consumidores versão 2006 (build, 13001.20498 ou superior). Para o Office no Windows, o recurso também tem suporte nos builds do Canal Semi-Annual e do Canal Empresarial Mensal disponíveis em 14 de julho de 2020 ou posterior. Os builds mínimos com suporte para cada canal são os seguinte:  

|Canal | Versão | Build|
|:-----|:-----|:-----|
|Canal Atual | 2006 ou superior | 20266.20266 ou superior|
|Canal Empresarial Mensal | 2005 ou superior | 12827.20538 ou superior|
|Canal Empresarial Mensal | 2004 | 12730.20602 ou superior|
|Canal Empresarial Semestral | 2002 ou superior | 12527.20880 ou superior|

## <a name="more-information"></a>Mais informações

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e build de versões de canal de atualização para clientes do Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar o número de versão e build de um aplicativo cliente do Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> O **conjunto de requisitos RibbonApi 1.1** ainda não tem suporte no manifesto, portanto, você não pode especificá-lo na seção do `<Requirements>` manifesto.


## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>RIBBON API 1.1

A API da Faixa de Opções 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência [Office.ribbon.](/javascript/api/office/office.ribbon)

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)