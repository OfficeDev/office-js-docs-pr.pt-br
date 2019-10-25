---
title: Conjuntos de requisitos da API JavaScript do Excel
description: Informações do conjunto de requisitos do Suplemento do Office para Builds do Excel
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 11fcf917b5dddb4a366465f11362c93660881597
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37681939"
---
# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="requirement-set-availability"></a>Disponibilidade do conjunto de requisitos

Os suplementos do Excel são executados em várias versões do Office, incluindo o Office 2016 ou posterior no Windows, Office na Web, Mac e iPad. A tabela a seguir lista conjuntos de requisitos do Excel, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e as versões ou números de build desses aplicativos.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados, faça referência à biblioteca **production** no CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Para obter informações sobre o uso de APIs de visualização, confira o artigo [APIs de visualização do JavaScript para Excel](./excel-preview-apis.md).

|  Conjunto de requisitos  |  Office no Windows<br>(conectado à assinatura do Office 365)  |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Visualização](excel-preview-apis.md)  | Use a versão mais recente do Office para testar as APIs de visualização (talvez seja exigido ser membro do [programa Office Insider](https://products.office.com/office-insider)) |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Versão 1907 (Build 11929.20306) ou posterior | 2.30 ou posterior. | 16.30 ou posterior | Outubro de 2019 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Versão 1903 (Build 11425.20204) ou posterior | 2.24 ou posterior | 16.24 ou posterior | Maio de 2019 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Versão 1808 (Build 10730.20102) ou posterior | 2.17 ou posterior | 16.17 ou posterior | Setembro de 2018 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Versão 1801 (Build 9001.2171) ou posterior   | 2.9 ou posterior  | 16.9 ou posterior  | Abril de 2018 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Versão 1704 (Compilação 8201.2001) ou posterior   | 2.2 ou posterior  | 15.36 ou posterior | Abril de 2017 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Versão 1703 (Compilação 8067.2070) ou posterior   | 2.2 ou posterior  | 15.36 ou posterior | Março de 2017 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Versão 1701 (build 7870.2024) ou posterior   | 2.2 ou posterior  | 15.36 ou posterior | Janeiro de 2017 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Versão 1608 (build 7369.2055) ou posterior   | 1.27 ou posterior | 15.27 ou posterior | Setembro de 2016 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Versão 1601 (build 6741.2088) ou posterior   | 1.21 ou posterior | 15.22 ou posterior | janeiro de 2016 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Versão 1509 (build 4266.1001) ou posterior   | 1.19 ou posterior | 15.20 ou posterior | janeiro de 2016 |

> [!NOTE]
> O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém o conjunto de requisitos 1.1 de ExcelApi.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

- 
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel)
- [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)
