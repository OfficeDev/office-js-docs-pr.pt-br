---
title: Abrir conjuntos de requisitos de janela do navegador
description: Especifica quais plataformas e compilações do Office suportam a API openBrowserWindow.
ms.date: 09/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8f6966f5bdcecd9c55a20f2d640d066906c1b6a3
ms.sourcegitcommit: 2479812e677d1a7337765fe8f1c8345061d4091a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/19/2020
ms.locfileid: "48135232"
---
# <a name="open-browser-window-api-requirement-sets"></a>Abrir conjuntos de requisitos da API da janela do navegador

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

O conjunto de APIs do OpenBrowserWindow permite que suplementos Abram um navegador para realizar tarefas que não podem ser executadas sempre no controle de WebView em modo seguro dentro do suplemento propriamente dito; por exemplo, baixar um arquivo PDF quando o controle de WebView é fornecido pelo Microsoft Edge.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos de API do OpenBrowserWindow, os aplicativos host do Office que dão suporte a esse conjunto de requisitos e os números de compilação ou versão para o aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 no Windows ou posterior<br>(compra avulsa) | Office no Windows<br>(conectado à assinatura do Office 365) |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1,1  | N/D | Versão 1810 (Build 16.0.11001.20074) ou posterior | 16.0.0.0 ou posterior | 16.0.0.0 ou posterior | N/D | N/D|

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1,1

O OpenBrowserWindowApi 1,1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência do [Office. Context. UI](/javascript/api/office/office.context.ui) .

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
