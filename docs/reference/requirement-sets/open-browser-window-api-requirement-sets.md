---
title: Abrir conjuntos de requisitos de janela do navegador
description: Especifica quais plataformas e builds do Office suportam a API openBrowserWindow.
ms.date: 04/09/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dd15136b350d42ec49187e436142aaecbfe70f40
ms.sourcegitcommit: 841bcad3c6c5139fd0953707c0be73ce890fa463
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2021
ms.locfileid: "51687430"
---
# <a name="open-browser-window-api-requirement-sets"></a>Conjuntos de requisitos da API da Janela do Navegador Aberto

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

O conjunto de API OpenBrowserWindow permite que os complementos abram um navegador para realizar tarefas que nem sempre podem ser feitas no controle de webview em áreas externas dentro do próprio add-in; por exemplo, baixando um arquivo PDF quando o controle webview é fornecido pelo Microsoft Edge.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API OpenBrowserWindow, os aplicativos host do Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 no Windows ou posterior<br>(compra avulsa) | Office no Windows<br>(Conectado à assinatura do Microsoft 365) |  Office no iPad<br>(Conectado à assinatura do Microsoft 365)  |  Office no Mac<br>(Conectado à assinatura do Microsoft 365)  | Office na Web  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | N/D | Versão 1810 (build 16.0.11001.20074) ou posterior | 16.0.0.0 ou posterior | 16.0.0.0 ou posterior | N/D | N/D|

> [!NOTE]
> O conjunto de requisitos OpenBrowserWindowApi só está disponível da seguinte maneira:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e com build de versões de canal de atualização para Aplicativos do Microsoft 365](/officeupdates/update-history-microsoft365-apps-by-date)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número de com build de um aplicativo cliente do Office](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

O OpenBrowserWindowApi 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência [office.context.ui.](/javascript/api/office/office.context#ui)

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
