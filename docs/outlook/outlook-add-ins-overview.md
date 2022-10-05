---
title: Visão geral dos suplementos do Outlook
description: Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web.
ms.date: 08/09/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: fd17728f840188fbedfdeba7d3ee8f97852d702a
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467255"
---
# <a name="outlook-add-ins-overview"></a>Visão geral dos suplementos do Outlook

Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:

- O mesmo suplemento e lógica de negócios funcionam em desktop (Outlook no Windows e Mac), na Web (Microsoft 365 e Outlook.com) e em dispositivos móveis.
- Os suplementos do Outlook consistem em um manifesto, que descreve como o suplemento se integra ao Outlook (por exemplo, um botão ou um painel de tarefas), e o código JavaScript/HTML, que compõe a interface do usuário e lógica de negócios do suplemento.
- Os suplementos do Outlook podem ser adquiridos na [AppSource](https://appsource.microsoft.com) ou [sideloaded](sideload-outlook-add-ins-for-testing.md) por usuários finais ou administradores.

Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.

The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>Pontos de extensão

Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done.

- Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

    **Suplemento com botões de comando na Faixa de Opções**

    ![Comando de função de suplemento.](../images/uiless-command-shape.png)

- Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).

    **Suplemento contextual para uma entidade realçada (um endereço)**

    ![Mostra um aplicativo contextual em um cartão.](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>Itens de caixa de correio disponíveis para suplementos

Os suplementos do Outlook são ativados quando o usuário está redigindo ou lendo uma mensagem ou compromisso, mas não em outros tipos de item. Entretanto, os suplementos *não* são ativados se o item de mensagem atual, em um formulário de redação ou de leitura, estiver em uma das seguintes situações:

- Protegido pelo IRM (Gerenciamento de Direitos de Informação) ou criptografado de outras maneiras para proteção e acessado do Outlook em clientes não Windows. Uma mensagem assinada digitalmente é um exemplo, já que a assinatura digital se baseia em um desses mecanismos.

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- Um relatório de entrega ou notificação que tem a classe de mensagem IPM.Report.*, incluindo NDRs (notificações de falha na entrega) e notificações de leitura, falha na leitura e atraso.

- Um arquivo .msg que é um anexo de outra mensagem.

- Um arquivo .msg aberto no sistema de arquivos.

- Em uma [caixa de correio de grupo](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), em uma caixa de correio compartilhada\*, em uma caixa de correio de outro usuário\*, em uma [caixa de correio de arquivo](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-client-and-compliance-&-security-feature-details?tabs=Archive-features#archive-mailbox), ou em uma pasta pública.

  > [!IMPORTANT]
  > \* Suporte para cenários de acesso de delegados (por exemplo, pastas compartilhadas da caixa de correio de outro usuário) foi introduzido no [conjunto de requisitos 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8). O suporte a caixas de correio compartilhadas agora está em visualização no Outlook no Windows e no Mac. Para saber mais, confira [Habilitar pastas compartilhadas e cenários de caixa de correio compartilhada](delegate-access.md).

- Usando um formulário personalizado.

- Criado através de MAPI simples. O MAPI simples é usado quando um usuário do Office cria ou envia um email de um aplicativo do Office no Windows com o Outlook fechado. Por exemplo, um usuário pode criar um email do Outlook enquanto trabalha no Word, o que dispara uma janela de composição do Outlook sem iniciar o aplicativo Outlook completo. No entanto, se o Outlook já estiver em execução quando o usuário criar o email a partir do Word, esse não será um cenário MAPI simples para que os suplementos do Outlook funcionem no formulário de composição, desde que outros requisitos de ativação sejam atendidos.

Em geral, o Outlook pode ativar suplementos no formato de leitura para itens na pasta Itens Enviados, com exceção dos suplementos que são ativados baseados em cadeias de correspondências de entidades conhecidas. Para obter mais informações sobre os motivos por trás disso, consulte [Suporte para entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md#support-for-well-known-entities).

Atualmente, há considerações adicionais ao projetar e implementar suplementos para clientes móveis. Para saber mais, confira [Adicionar suporte móvel a um suplemento do Outlook](add-mobile-support.md#compose-mode-and-appointments).

## <a name="supported-clients"></a>Clientes com suporte

Suplementos do Outlook são compatíveis com o Outlook 2013 ou posterior no Windows, Outlook 2016 ou posterior no Mac, Outlook na Web para Exchange 2013 no local e versões posteriores, Outlook no iOS, Outlook no Android e Outlook na Web e Outlook.com. Nem todos os recursos mais recentes são compatíveis com todos os [clientes](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) ao mesmo tempo. Confira os artigos e as referências de API para esses recursos e saiba com quais aplicativos eles podem ou não ter compatibilidade.

## <a name="get-started-building-outlook-add-ins"></a>Introdução à criação de suplementos do Outlook

Para começar a construir suplementos do Outlook, tente o seguinte:

- [Início Rápido](../quickstarts/outlook-quickstart.md) - Criar um painel de tarefas simples.
- [Tutorial](../tutorials/outlook-tutorial.md) : saiba como criar um suplemento que insere gists do GitHub em uma nova mensagem.

## <a name="see-also"></a>Confira também

- [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Diretrizes de design para Suplementos do Office](../design/add-in-design.md)
- [Licenciar suplementos do Office e do SharePoint](/office/dev/store/license-your-add-ins)
- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-the-office-store)
