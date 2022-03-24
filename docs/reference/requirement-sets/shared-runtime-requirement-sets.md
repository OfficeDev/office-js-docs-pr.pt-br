---
title: Conjuntos de requisitos de tempo de execução compartilhados
description: Especifica as plataformas e Office que suportam as APIs sharedRuntime.
ms.date: 03/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: ef4bb9aebea8ee6f9c316a68cac784c20c4db160
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746315"
---
# <a name="shared-runtime-requirement-sets"></a>Conjuntos de requisitos de tempo de execução compartilhados

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Partes de um Office que executem código JavaScript, como painéis de tarefas, arquivos de função lançados de comandos de Excel e funções personalizadas do Excel, podem compartilhar um único tempo de execução do JavaScript. Isso permite que todas as partes compartilhem um conjunto de variáveis globais, compartilhem um conjunto de bibliotecas carregadas e se comuniquem entre si sem precisar passar mensagens por meio de um armazenamento persistente. Para obter mais informações, [consulte Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).

A tabela a seguir lista o conjunto de requisitos SharedRuntime 1.1, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

| Conjunto de requisitos | Office 2021 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) | Office no iPad<br>(conectado a uma assinatura do Microsoft 365) | Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)  | Office na Web | Servidor do Office Online |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Build 16.0.14326.20454 ou posterior | Versão 2002 (build 12527.20092) ou posterior | N/D | 16.35 ou posterior | Fevereiro de 2020 | N/D |

> [!IMPORTANT]
> No momento, o tempo de execução de JavaScript compartilhado não é compatível com iPad ou em versões de compra avulsa do Office 2019 ou anterior. Para obter detalhes adicionais de suporte, consulte as seções a seguir.

## <a name="support-for-version-11-on-excel"></a>Suporte para a versão 1.1 no Excel

O conjunto de requisitos SharedRuntime 1.1 é lançado para Excel na Web, Windows e Mac.

## <a name="preview-support-for-version-11-on-word-and-powerpoint"></a>Visualizar suporte para a versão 1.1 no Word e PowerPoint

A tabela a seguir lista builds de aplicativo adicionais que suportam uma visualização do tempo de execução javaScript compartilhado. A versão de visualização do tempo de execução compartilhado está sujeita a alterações. Não é compatível para uso em ambientes de produção. Para obter o build mais recente, você precisa [Ingressar no Office Insider](https://insider.office.com/join). Uma boa maneira de experimentar os recursos de pré-visualização é usando uma assinatura do Microsoft 365. Se você ainda não tem uma assinatura do Microsoft 365, pode obter uma ingressando no[ programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/office/dev-program).

|Aplicativo do Office |Compilar |
|-------------------|------|
|PowerPoint no Windows |Build 16.0.13218.10000 ou posterior |
|PowerPoint no Mac |Build Build 16.46.207.0 ou posterior |
|PowerPoint Online | Fevereiro de 2022 |
|Word no Windows |Build 16.0.13218.10000 ou posterior |
|Word no Mac |Build 16.46.207.0 ou posterior |

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
