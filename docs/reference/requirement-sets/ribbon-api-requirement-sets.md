---
title: Conjuntos de requisitos comuns da API
description: Especifica quais plataformas e compilações do Office oferecem suporte às APIs de faixa de opções dinâmicas.
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 878670367b253fa7700434681244b43b9cfa36a7
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996512"
---
# <a name="ribbon-api-requirement-sets"></a>Conjuntos de requisitos comuns da API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

O conjunto de API da faixa de opções suporta o controle programático de quando comandos de suplemento personalizados (ou seja, botões de faixa de opções personalizados e itens de menu) estão habilitados e desabilitados.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API da faixa de opções, os aplicativos cliente do Office que dão suporte a esse conjunto de requisitos e os números de versão ou de compilação do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 no Windows<br>(compra avulsa) | Office 2016 ou posterior no Windows<br>(compra avulsa)   | Office no Windows\*<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac\*<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web\*  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1,1  | N/D | N/D | Consulte suporte<br>seção abaixo | N/A | 16,38 | Novembro de 2020 | N/A|

> **&#42;** A API da faixa de opções só tem suporte no Excel e requer assinatura do Microsoft 365.

## <a name="office-on-windows-subscription-support"></a>Suporte do Office no Windows (assinatura)

O conjunto de requisitos é suportado no canal de consumidor versão 2006 (Build, 13001,20498 ou posterior). Para o Office no Windows, o recurso também é suportado no canal de Semi-Annual e nas versões de canal corporativos mensais disponíveis, 14 de julho de 2020 ou mais recente. As compilações mínimas suportadas para cada canal são as seguintes:  

|Canal | Versão | Build|
|:-----|:-----|:-----|
|Canal Atual | 2006 ou maior | 20266,20266 ou maior|
|Canal Empresarial Mensal | 2005 ou maior | 12827,20538 ou maior|
|Canal Empresarial Mensal | 2004 | 12730,20602 ou maior|
|Canal Empresarial Semestral | 2002 ou maior | 12527,20880 ou maior|

## <a name="more-information"></a>Mais informações

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e de build de lançamentos de canais de atualização para clientes do Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número do Build para um aplicativo cliente Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> O conjunto de requisitos **RibbonApi 1,1** ainda não tem suporte no manifesto, portanto, você não pode especificá-lo na seção do manifesto `<Requirements>` .


## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API da faixa de opções 1,1

A API da faixa de opções 1,1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência da faixa de opções do [Office ](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de aplicativos do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
