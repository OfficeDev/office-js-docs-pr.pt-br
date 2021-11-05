---
title: Conjuntos de requisitos de tempo de execução compartilhados
description: Especifica as plataformas e Office que suportam as APIs sharedRuntime.
ms.date: 11/03/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: a5f7d3c9394de047b358d7f190c5adae5b5199b1
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793600"
---
# <a name="shared-runtime-requirement-sets"></a>Conjuntos de requisitos de tempo de execução compartilhados

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Partes de um Office que executem código JavaScript, como painéis de tarefas, arquivos de função lançados a partir de comandos de Excel e funções personalizadas, podem compartilhar um único tempo de execução do JavaScript. Isso permite que todas as partes compartilhem um conjunto de variáveis globais, compartilhem um conjunto de bibliotecas carregadas e se comuniquem entre si sem precisar passar mensagens por meio de um armazenamento persistente. Para obter mais informações, [consulte Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).

A tabela a seguir lista o conjunto de requisitos SharedRuntime 1.1, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

| Conjunto de requisitos | Office 2021 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) | Office no iPad<br>(conectado a uma assinatura do Microsoft 365) | Office no Mac<br>(conectado a uma assinatura do Microsoft 365) | Office na Web | Servidor do Office Online |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Build 16.0.14326.20454 ou posterior | Versão 2002 (build 12527.20092) ou posterior | Não disponível | 16.35 ou posterior | Fevereiro de 2020 | Não disponível |

> [!IMPORTANT]
> No momento, o tempo de execução de JavaScript compartilhado não é compatível com iPad ou em versões de compra avulsa do Office 2019 ou anterior. Para obter detalhes adicionais de suporte, consulte as seções a seguir.

## <a name="support-for-version-11-on-excel"></a>Suporte para a versão 1.1 no Excel

O conjunto de requisitos SharedRuntime 1.1 é lançado para Excel na Web, Windows e Mac.

## <a name="preview-support-for-version-11-on-word-and-powerpoint"></a>Visualizar suporte para a versão 1.1 no Word e PowerPoint

A tabela a seguir lista builds de aplicativo adicionais que suportam uma visualização do tempo de execução javaScript compartilhado. A versão de visualização do tempo de execução compartilhado está sujeita a alterações. Não é compatível para uso em ambientes de produção. Para obter o build mais recente, você precisa [Ingressar no Office Insider](https://insider.office.com/join). Uma boa maneira de experimentar os recursos de pré-visualização é usando uma assinatura do Microsoft 365. Se você ainda não tem uma assinatura do Microsoft 365, pode obter uma ingressando no[ programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/office/dev-program).

|Aplicativo do Office |Compilar |
|-------------------|------|
|PowerPoint no Windows |Build 16.0.13218.10000 ou posterior |
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
