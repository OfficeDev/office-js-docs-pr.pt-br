---
title: Conjuntos de requisitos da API JavaScript do Excel
description: Informações do conjunto de requisitos do Suplemento do Office para builds do Excel.
ms.date: 10/26/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: f99e9033d4b5acbcba6c4f799bcc73b263cfaf6c
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774724"
---
# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Disponibilidade do conjunto de requisitos

Os suplementos do Excel são executados em várias versões do Office, incluindo o Office 2016 ou posterior no Windows, Office na Web, Mac e iPad. A tabela a seguir lista conjuntos de requisitos do Excel, ou seja, os aplicativos do cliente Office que oferecem suporte a esse conjunto de requisitos, e as versões ou números de compilação desses aplicativos.

> [!NOTE]
> Para usar APIs em qualquer um dos conjuntos de requisitos numerados ou `ExcelApiOnline`, faça referência à biblioteca **production** no CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Para obter informações sobre o uso de APIs de visualização, confira o artigo [APIs de visualização do JavaScript para Excel](excel-preview-apis.md).

|  Conjunto de requisitos  |  Office no Windows<br>(conectado a uma assinatura do Microsoft 365)  |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Visualização](excel-preview-apis.md)  | Use a versão mais recente do Office para experimentar APIs de visualização (pode ser necessário ingressar no [Programa Office Insider](https://insider.office.com)). |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | Não disponível | N/D | Não disponível | Mais recente (confira a [página conjunto de requisitos](excel-api-online-requirement-set.md)) |
| [ExcelApi 1.12](excel-api-1-12-requirement-set.md) | Versão 2008 (Build 13127.20408) ou posterior | 16.40 ou posterior | 16.40 ou posterior | Setembro de 2020 |
| [ExcelApi 1.11](excel-api-1-11-requirement-set.md) | Versão 2002 (Build 12527.20470) ou posterior | 16.35 ou posterior | 16.33 ou posterior | Maio de 2020 |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Versão 1907 (Build 11929.20306) ou posterior | 16.0 ou posterior | 16.30 ou posterior | Outubro de 2019 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Versão 1903 (Build 11425.20204) ou posterior | 16.0 ou posterior | 16.24 ou posterior | Maio de 2019 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Versão 1808 (Build 10730.20102) ou posterior | 16.0 ou posterior | 16.17 ou posterior | Setembro de 2018 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Versão 1801 (Build 9001.2171) ou posterior   | 16.0 ou posterior  | 16.9 ou posterior  | Abril de 2018 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Versão 1704 (Compilação 8201.2001) ou posterior   | 15.0 ou posterior  | 15.36 ou posterior | Abril de 2017 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Versão 1703 (Compilação 8067.2070) ou posterior   | 15.0 ou posterior  | 15.36 ou posterior | Março de 2017 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Versão 1701 (build 7870.2024) ou posterior   | 15.0 ou posterior  | 15.36 ou posterior | Janeiro de 2017 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Versão 1608 (build 7369.2055) ou posterior   | 15.0 ou posterior | 15.27 ou posterior | Setembro de 2016 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Versão 1601 (build 6741.2088) ou posterior   | 15.0 ou posterior | 15.22 ou posterior | janeiro de 2016 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Versão 1509 (build 4266.1001) ou posterior   | 15.0 ou posterior | 15.20 ou posterior | janeiro de 2016 |

> [!NOTE]
> Versões permanentes dos conjuntos de requisitos de suporte do Office como a seguir:
>
> - O Office 2019 é compatível com o ExcelApi 1.8 e versões anteriores.
> - O Office 2016 é compatível somente com o conjunto de requisitos do ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="how-to-use-excel-requirement-sets-at-runtime-and-in-the-manifest"></a>Como usar os conjuntos de requisitos do Excel no tempo de execução e no manifesto

> [!NOTE]
> Esta seção pressupõe que você esteja familiarizado com a visão geral dos conjuntos de requisitos em [Versões e conjuntos de requisitos do Office](../../develop/office-versions-and-requirement-sets.md) e [Especificar aplicativos do Office e requisitos de API](../../develop/specify-office-hosts-and-api-requirements.md).

Os conjuntos de requisitos são grupos nomeados de membros da API. Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um aplicativo do Office dá suporte às APIs necessárias ao suplemento.

### <a name="checking-for-requirement-set-support-at-runtime"></a>Verificando o suporte ao conjunto de requisitos no tempo de execução

O exemplo de código a seguir mostra como determinar se o aplicativo do Office, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Definindo o suporte ao conjunto de requisitos no manifesto

Você pode usar o [elemento Requirements](../manifest/requirements.md) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado. Se a plataforma ou o aplicativo do Office não for compatível com os conjuntos de requisitos ou métodos de API especificados no `Requirements` elemento do manifesto, o suplemento não será executado nesse aplicativo ou plataforma, e não será exibido na lista de suplementos mostrados no **Meus suplementos** . Se o seu suplemento exige um conjunto específico de requisitos para funcionalidade total, mas pode fornecer um valor mesmo para os usuários nas plataformas que não têm suporte para o conjunto de requisitos, recomendamos verificar o suporte a requisitos no tempo de execução conforme descrito acima, em vez de definir o suporte ao conjunto de requisitos no manifesto.

O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos cliente do Office que dão suporte ao conjunto de requisitos ExcelApi, versão 1.3 ou superior.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
