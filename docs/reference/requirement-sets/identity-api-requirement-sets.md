---
title: Conjuntos de requisitos da API de Identidade
description: ''
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 6460516c32f70995d43ae64b213c93153a015bff
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600687"
---
# <a name="identity-api-requirement-sets"></a>Conjuntos de requisitos da API de Identidade

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de Identidade, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  | Office 2013 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado à assinatura do Office 365) |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web  | SharePoint Online | OneDrive.com |Outlook.com e Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Visualização do IdentityAPI  | N/D | Versão prévia<b>*</b> | Em breve | Versão prévia<b>*</b> | Visualização<b>* &#8224;</b> | Visualização<b>* &#8224;</b>| Em breve | Em breve |

> **&#42;** Durante a fase de visualização, a API de identidade requer o Office 365 (a versão de assinatura do Office). Você deve usar o build e a versão mensal mais recente do canal Insiders. É necessário ingressar no programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.
>
> **&#8224;** Os suplementos que usam as APIs SSO nessas plataformas só funcionarão se o administrador de locatário do usuário tiver concedido o consentimento para o suplemento. O usuário não pode conceder consentimento mesmo ao seu próprio perfil do Azure AD.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>Visualização do IdentityAPI

Para obter detalhes sobre essa API, consulte a versão que usa promessas em [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) ou a versão que usa retornos de chamada em [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
