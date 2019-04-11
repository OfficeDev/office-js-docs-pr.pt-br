---
title: Conjuntos de requisitos da API de Identidade
description: ''
ms.date: 04/09/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 9229483bcf2916d35bd1fc8961c2c2a73cf9caed
ms.sourcegitcommit: fbe2a799fda71aab73ff1c5546c936edbac14e47
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/10/2019
ms.locfileid: "31764387"
---
# <a name="identity-api-requirement-sets"></a>Conjuntos de requisitos da API de Identidade

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de Identidade, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  | Office 2013 ou posterior para Windows | Office 365 para Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com e Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/D | Versão prévia<b>*</b> | Em breve | Versão prévia<b>*</b> | Versão prévia<b>*</b> | Versão prévia<b>*</b>| Em breve | Em breve |

> **& # 42;** Durante a fase de visualização, a API de identidade requer o Office 365 (a versão de assinatura do Office). Você deve usar o build e a versão mensal mais recente do canal Insiders. É necessário ingressar no programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e compilação de versões de canal de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral do Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1

A IdentityAPI 1.1 de Logon Único é a primeira versão da API. Para saber mais sobre essa API, confira a seção [Referência da API de SSO](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) em [Habilitar o SSO em um suplemento](/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar hosts do Office e requisitos de API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
