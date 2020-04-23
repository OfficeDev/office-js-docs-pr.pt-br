---
title: Conjuntos de requisitos da API JavaScript do Excel
description: Informações do conjunto de requisitos do Suplemento do Office para builds do Excel.
ms.date: 04/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6da9e34359521157e809764907c3a6c3a62ae76c
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547224"
---
# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Disponibilidade do conjunto de requisitos

Os suplementos do Excel são executados em várias versões do Office, incluindo o Office 2016 ou posterior no Windows, Office na Web, Mac e iPad. A tabela a seguir lista conjuntos de requisitos do Excel, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e as versões ou números de build desses aplicativos.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados ou `ExcelApiOnline`, faça referência à biblioteca **production** no CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Para obter informações sobre o uso de APIs de visualização, confira o artigo [APIs de visualização do JavaScript para Excel](excel-preview-apis.md).

|  Conjunto de requisitos  |  Office no Windows<br>(conectado à assinatura do Office 365)  |  Office no iPad<br>(conectado à assinatura do Office 365)  |  Office no Mac<br>(conectado à assinatura do Office 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Visualização](excel-preview-apis.md)  | Use a versão mais recente do Office para testar as APIs de visualização (talvez seja exigido ser membro do [programa Office Insider](https://insider.office.com)) |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | Não disponível | N/D | Não disponível | Mais recente (confira a [página conjunto de requisitos](./excel-api-online-requirement-set.md)) |
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
> Versões permanentes dos conjuntos de requisitos de suporte do Office como a seguir:
>
> - O Office 2019 é compatível com o ExcelApi 1.8 e versões anteriores.
> - O Office 2016 é compatível somente com o conjunto de requisitos do ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>Também consulte

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)
