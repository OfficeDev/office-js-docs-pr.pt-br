---
title: Conjuntos de requisitos comuns da API
description: Especifica quais plataformas e compilações do Office oferecem suporte às APIs de faixa de opções dinâmicas.
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 1d8f17d56b779fb1590daa82e3c391423c3bfc9a
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159189"
---
# <a name="ribbon-api-requirement-sets"></a>Conjuntos de requisitos comuns da API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

O conjunto de API da faixa de opções suporta o controle programático de quando comandos de suplemento personalizados (ou seja, botões de faixa de opções personalizados e itens de menu) estão habilitados e desabilitados.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API da faixa de opções, os aplicativos host do Office que dão suporte a esse conjunto de requisitos e os números de versão ou de compilação do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 no Windows<br>(compra avulsa) | Office 2016 ou posterior no Windows<br>(compra avulsa)   | Office no Windows\*<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac\*<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web\*  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1,1  | N/D | N/D | Versão 2002 (Build 12527,20264) ou posterior | 16,38 ou posterior | N/D | Fevereiro de 2020 | N/D|

> **&#42;** Durante a fase de visualização, a API da faixa de opções só tem suporte no Excel e requer assinatura do Microsoft 365. Você deve usar o build e a versão mensal mais recente do canal Insiders. É necessário ingressar no programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). Observe que, quando uma compilação é graduada para o canal semestral de produção, o suporte para recursos de visualização, incluindo a API da faixa de opções, é desativado para essa compilação.

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e compilação de versões de canal de atualização para clientes Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número do Build para um aplicativo cliente Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API da faixa de opções 1,1

A API da faixa de opções 1,1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência da faixa de opções do [Office](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
