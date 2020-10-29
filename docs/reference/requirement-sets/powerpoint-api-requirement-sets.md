---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do PowerPoint.
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774723"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do PowerPoint

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos do cliente Office que oferecem suporte a esses conjuntos de requisitos e as versões de compilação ou datas de disponibilidade.

|  Conjunto de requisitos  |  Office no Windows<br>(conectado a uma assinatura do Microsoft 365)  |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Visualização](powerpoint-preview-apis.md)  | Use a versão mais recente do Office para experimentar APIs de visualização (pode ser necessário ingressar no [Programa Office Insider](https://insider.office.com)). |
| PowerPointApi 1.1 | Versão 1810 (Build 11001.20074) ou posterior | 2.17 ou posterior | 16.19 ou posterior | Outubro de 2018 |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre as versões do Office e os números de build, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>API JavaScript do PowerPoint 1.1

O PowerPoint JavaScript API 1.1 contém uma [única API para criar uma nova apresentação](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-). Para obter detalhes sobre a API, confira [Criar uma apresentação](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a>Como usar os conjuntos de requisitos do PowerPoint em tempo de execução e no manifesto

> [!NOTE]
> Esta seção pressupõe que você esteja familiarizado com a visão geral dos conjuntos de requisitos em [Versões e conjuntos de requisitos do Office](../../develop/office-versions-and-requirement-sets.md) e [Especificar aplicativos do Office e requisitos de API](../../develop/specify-office-hosts-and-api-requirements.md).

Os conjuntos de requisitos são grupos nomeados de membros da API. Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um aplicativo do Office dá suporte às APIs necessárias ao suplemento.

### <a name="checking-for-requirement-set-support-at-runtime"></a>Verificando o suporte ao conjunto de requisitos no tempo de execução

O exemplo de código a seguir mostra como determinar se o aplicativo do Office, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Definindo o suporte ao conjunto de requisitos no manifesto

Você pode usar o [elemento Requirements](../manifest/requirements.md) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado. Se a plataforma ou o aplicativo do Office não for compatível com os conjuntos de requisitos ou métodos de API especificados no `Requirements` elemento do manifesto, o suplemento não será executado nesse aplicativo ou plataforma, e não será exibido na lista de suplementos mostrados no **Meus suplementos** . Se o seu suplemento exige um conjunto específico de requisitos para funcionalidade total, mas pode fornecer um valor mesmo para os usuários nas plataformas que não têm suporte para o conjunto de requisitos, recomendamos verificar o suporte a requisitos no tempo de execução conforme descrito acima, em vez de definir o suporte ao conjunto de requisitos no manifesto.

O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica que o suplemento deve ser carregado em todos os aplicativos do cliente do Office que oferecem suporte ao conjunto de requisitos da versão 1.1 ou superior do PowerPointApi.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns. Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Documentação de Referência da API JavaScript do PowerPoint](/javascript/api/powerpoint)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
