---
title: Conjuntos de requisitos da API de Identidade
description: Informações do conjunto de requisitos da API de identidade para os Complementos do Office.
ms.date: 01/26/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c662e7a5306692fd75de51acc7cadfd1df3e7406
ms.sourcegitcommit: 85b4839be743059bf155ff44e49d64968444d80a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/31/2021
ms.locfileid: "51471721"
---
# <a name="identity-api-requirement-sets"></a>Conjuntos de requisitos da API de Identidade

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de Identidade, os aplicativos cliente do Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | N/D | 2008 (build 13127.20000) ou posterior | Em breve | 16.40 ou posterior | Microsoft SharePoint Online e OneDrive\* |

\* Atualmente, o conjunto de requisitos é suportado no Office na Web apenas para documentos abertos do Microsoft SharePoint Online e do OneDrive.

> [!NOTE]
> Outlook: Para exigir o conjunto de API de identidade 1.3 no código do seu complemento, verifique se ele tem suporte chamando `isSetSupported('IdentityAPI', '1.3')` . Não há suporte para declará-lo no manifesto do complemento do Outlook. Você também pode determinar se a API tem suporte, verificando se ela não é `undefined`. Para mais detalhes, confira [Usar APIs de conjuntos de requisitos posteriores](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>IdentityAPI Preview

Para obter detalhes sobre essa API, consulte a versão que usa Promessas em [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) ou a versão que usa retornos de chamada em [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
