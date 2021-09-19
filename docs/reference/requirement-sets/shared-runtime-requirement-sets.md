---
title: Conjuntos de requisitos de tempo de execução compartilhados
description: Especifica as plataformas e Office que suportam as APIs sharedRuntime.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: b4e7d37e66a562799bc841fd7d7e7ad8cd6d89e7
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/19/2021
ms.locfileid: "59448716"
---
# <a name="shared-runtime-requirement-sets"></a>Conjuntos de requisitos de tempo de execução compartilhados

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Partes de um Office que executem código JavaScript, como painéis de tarefas, arquivos de função lançados a partir de comandos de Excel e funções personalizadas, podem compartilhar um único tempo de execução do JavaScript. Isso permite que todas as partes compartilhem um conjunto de variáveis globais, compartilhem um conjunto de bibliotecas carregadas e se comuniquem entre si sem precisar passar mensagens por meio de um armazenamento persistente. Para obter mais informações, [consulte Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).

A tabela a seguir lista o conjunto de requisitos SharedRuntime 1.1, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou version do aplicativo Office.

| Conjunto de requisitos | Office 2021 ou posterior no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365) | Office no iPad<br>(conectado a uma assinatura do Microsoft 365) | Office no Mac<br>(conectado a uma assinatura do Microsoft 365) | Office na Web | Servidor do Office Online |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Versão 2002 (build 12527.20092) ou posterior | Versão 2002 (build 12527.20092) ou posterior | N/A | 16.35 ou posterior | Fevereiro de 2020 | N/A |

> [!IMPORTANT]
> O conjunto de requisitos de tempo de execução do JavaScript compartilhado só está disponível nas seguintes plataformas.
>
> - Excel na web, Windows e Mac.
> - PowerPoint no Windows (build 13218.10000 ou posterior). O tempo de execução de JavaScript compartilhado para PowerPoint está atualmente em pré-visualização e sujeito a alterações. Não é compatível para uso em ambientes de produção. Para obter o build mais recente, você precisa [Ingressar no Office Insider](https://insider.office.com/join). Uma boa maneira de experimentar os recursos de pré-visualização é usando uma assinatura do Microsoft 365. Se você ainda não tem uma assinatura do Microsoft 365, pode obter uma ingressando no[ programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/office/dev-program).
>
> No momento, o tempo de execução de JavaScript compartilhado não é compatível com iPad ou em versões de compra avulsa do Office 2019 ou anterior.

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
