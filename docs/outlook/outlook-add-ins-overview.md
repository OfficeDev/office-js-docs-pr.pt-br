---
title: Visão geral dos suplementos do Outlook
description: Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web.
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 802ea8ef55fd97e938ce945fe934d2584eea836f
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741138"
---
# <a name="outlook-add-ins-overview"></a>Visão geral dos suplementos do Outlook

Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web. Os suplementos do Outlook têm três aspectos principais:

- O mesmo suplemento e lógica de negócios funcionam em desktop (Outlook no Windows e Mac), na Web (Microsoft 365 e Outlook.com) e em dispositivos móveis.
- Os suplementos do Outlook consistem em um manifesto, que descreve como o suplemento se integra ao Outlook (por exemplo, um botão ou um painel de tarefas), e o código JavaScript/HTML, que compõe a interface do usuário e lógica de negócios do suplemento.
- Os suplementos do Outlook podem ser adquiridos na [AppSource](https://appsource.microsoft.com) ou [sideloaded](sideload-outlook-add-ins-for-testing.md) por usuários finais ou administradores.

Os suplementos do Outlook são diferentes dos suplementos de COM ou VSTO, que são integrações mais antigas específicas do Outlook para Windows. Diferentemente dos suplementos de COM, os suplementos do Outlook não têm qualquer código fisicamente instalado no dispositivo do usuário ou no cliente do Outlook. No caso de um suplemento do Outlook, o Outlook lê o manifesto, conecta os controles especificados na interface do usuário e carrega o HTML e o JavaScript. Todos os componentes Web são executados no contexto do navegador em uma área restrita.

Os itens do Outlook que dão suporte a suplementos incluem mensagens de email, compromissos, solicitações, respostas e cancelamentos de reunião. Cada suplemento do Outlook define o contexto no qual está disponível, incluindo os tipos de itens e se o usuário está lendo ou redigindo um item.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>Pontos de extensão

Pontos de extensão são as formas usadas pelos suplementos para se integrar ao Outlook. Estas são as maneiras de fazer isso:

- Os suplementos podem declarar botões que aparecem nas superfícies de comando em mensagens e compromissos. Para saber mais, confira [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md).

    **Suplemento com botões de comando na Faixa de Opções**

    ![Comando de suplemento de forma sem interface do usuário](../images/uiless-command-shape.png)

- Os suplementos podem desvincular correspondências de expressões regulares ou entidades detectadas em mensagens e compromissos. Para saber mais, confira [Suplementos contextuais do Outlook](contextual-outlook-add-ins.md).

    **Suplemento contextual para uma entidade realçada (um endereço)**

    ![Mostra um aplicativo contextual em um cartão](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>Itens de caixa de correio disponíveis para suplementos

Os suplementos do Outlook são ativados quando o usuário está redigindo ou lendo uma mensagem ou compromisso, mas não em outros tipos de item. Entretanto, os suplementos *não* são ativados se o item de mensagem atual, em um formulário de redação ou de leitura, estiver em uma das seguintes situações:

- Protegido por IRM (Gerenciamento de Direitos de Informação) ou criptografado de outras maneiras para proteção. Uma mensagem assinada digitalmente é um exemplo, já que a assinatura digital se baseia em um desses mecanismos.

  > [!IMPORTANT]
  > - Os suplementos são ativados em mensagens assinadas digitalmente no Outlook associadas a uma assinatura do Microsoft 365. No Windows, esse suporte foi introduzido com a compilação 8711.1000.
  >
  > - A partir do Outlook, build 13229.10000, no Windows, os suplementos agora podem ser ativados nos itens protegidos por IRM. Para obter mais informações sobre esse recurso na visualização, consulte [Ativação de suplementos em itens protegidos pela Gestão de Direitos de Informação (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).

- Um relatório de entrega ou notificação que tem a classe de mensagem IPM.Report.*, incluindo NDRs (notificações de falha na entrega) e notificações de leitura, falha na leitura e atraso.

- Um rascunho (não tem um remetente atribuído a ele) ou está na pasta Rascunhos do Outlook.

- Um arquivo .msg que é um anexo de outra mensagem.

- Um arquivo .msg aberto no sistema de arquivos.

- Em uma caixa de correio compartilhada, na caixa de correio de outro usuário, em uma caixa de correio de arquivo morto ou em uma pasta pública.

- Usando um formulário personalizado.

Em geral, o Outlook pode ativar suplementos no formato de leitura para itens na pasta Itens Enviados, com exceção dos suplementos que são ativados baseados em cadeias de correspondências de entidades conhecidas. Para saber mais sobre os motivos por trás disso, confira "Suporte para entidades conhecidas" em [Corresponder cadeias em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).

## <a name="supported-clients"></a>Clientes com suporte

Suplementos do Outlook são compatíveis com o Outlook 2013 ou posterior no Windows, Outlook 2016 ou posterior no Mac, Outlook na Web para Exchange 2013 no local e versões posteriores, Outlook no iOS, Outlook no Android e Outlook na Web e Outlook.com. Nem todos os recursos mais recentes são compatíveis com todos os [clientes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) ao mesmo tempo. Confira os artigos e as referências de API para esses recursos e saiba com quais aplicativos eles podem ou não ter compatibilidade.


## <a name="get-started-building-outlook-add-ins"></a>Introdução à criação de suplementos do Outlook

Para começar a criar suplementos do Outlook, experimente o seguinte.

- [Início Rápido](../quickstarts/outlook-quickstart.md) - Criar um painel de tarefas simples.
- [Tutorial](../tutorials/outlook-tutorial.md) : saiba como criar um suplemento que insere gists do GitHub em uma nova mensagem.


## <a name="see-also"></a>Confira também
- [Saiba mais sobre o Programa de Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Diretrizes de design para Suplementos do Office](../design/add-in-design.md)
- [Licenciar suplementos do Office e do SharePoint](/office/dev/store/license-your-add-ins)
- [Publicar seu Suplemento do Office](../publish/publish.md)
- [Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-the-office-store)
