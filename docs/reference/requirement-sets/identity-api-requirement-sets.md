---
title: Conjuntos de requisitos da API de Identidade
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 4ed3fde5b3cef6dd6633d9eafbd098b24cc27308
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457658"
---
# <a name="identity-api-requirement-sets"></a>Conjuntos de requisitos da API de Identidade

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de Identidade, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  | Office 2013 para Windows | Office 365 para Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com e Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/D | Versão prévia **&#42;** | Em breve | Versão prévia **&#42;**| Versão prévia | Versão prévia| Em breve | Em breve |

> **&#42;** Durante a fase de versão prévia, a API de Identidade é compatível com o Mac e Windows 2016 somente para usuários do programa Office Insider usando o Modo rápido. Para participar do programa Office Insider, confira [Faça parte do Office Insider](https://products.office.com/office-insider?tab=tab-1). Para alternar para o FastTrack, confira [Insider – Modo rápido](https://answers.microsoft.com/pt-BR/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Para saber mais sobre as versões, números de build e sobre o Servidor do Office Online, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

A IdentityAPI 1.1 de Logon Único é a primeira versão da API. Para saber mais sobre essa API, confira a seção [Referência da API de SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) em [Habilitar o SSO em um suplemento](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
