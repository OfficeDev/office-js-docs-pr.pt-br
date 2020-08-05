---
title: Visão geral dos suplementos do Outlook
description: Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web.
ms.date: 08/03/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: a86232d8208084ce5c5b0dd5009e8c1a9f0b5837
ms.sourcegitcommit: a3b743598025466bad19177e0ba9ca94ea66d490
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/04/2020
ms.locfileid: "46547546"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="3c164-103">Visão geral dos suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="3c164-103">Outlook add-ins overview</span></span>

<span data-ttu-id="3c164-104">Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web.</span><span class="sxs-lookup"><span data-stu-id="3c164-104">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform.</span></span> <span data-ttu-id="3c164-105">Os suplementos do Outlook têm três aspectos principais:</span><span class="sxs-lookup"><span data-stu-id="3c164-105">Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="3c164-106">O mesmo suplemento e lógica de negócios funcionam em desktop (Outlook no Windows e Mac), na Web (Microsoft 365 e Outlook.com) e em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="3c164-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="3c164-107">Os suplementos do Outlook consistem em um manifesto, que descreve como o suplemento se integra ao Outlook (por exemplo, um botão ou um painel de tarefas), e o código JavaScript/HTML, que compõe a interface do usuário e lógica de negócios do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3c164-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="3c164-108">Os suplementos do Outlook podem ser adquiridos na [AppSource](https://appsource.microsoft.com) ou [sideloaded](sideload-outlook-add-ins-for-testing.md) por usuários finais ou administradores.</span><span class="sxs-lookup"><span data-stu-id="3c164-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="3c164-109">Os suplementos do Outlook são diferentes dos suplementos de COM ou VSTO, que são integrações mais antigas específicas do Outlook para Windows.</span><span class="sxs-lookup"><span data-stu-id="3c164-109">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows.</span></span> <span data-ttu-id="3c164-110">Diferentemente dos suplementos de COM, os suplementos do Outlook não têm qualquer código fisicamente instalado no dispositivo do usuário ou no cliente do Outlook.</span><span class="sxs-lookup"><span data-stu-id="3c164-110">Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client.</span></span> <span data-ttu-id="3c164-111">No caso de um suplemento do Outlook, o Outlook lê o manifesto, conecta os controles especificados na interface do usuário e carrega o HTML e o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3c164-111">For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML.</span></span> <span data-ttu-id="3c164-112">Todos os componentes Web são executados no contexto do navegador em uma área restrita.</span><span class="sxs-lookup"><span data-stu-id="3c164-112">The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="3c164-113">Os itens do Outlook que dão suporte a suplementos incluem mensagens de email, compromissos, solicitações, respostas e cancelamentos de reunião.</span><span class="sxs-lookup"><span data-stu-id="3c164-113">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments.</span></span> <span data-ttu-id="3c164-114">Cada suplemento do Outlook define o contexto no qual está disponível, incluindo os tipos de itens e se o usuário está lendo ou redigindo um item.</span><span class="sxs-lookup"><span data-stu-id="3c164-114">Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="3c164-115">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3c164-115">Extension points</span></span>

<span data-ttu-id="3c164-p104">Pontos de extensão são as formas usadas pelos suplementos para se integrar ao Outlook. Estas são as maneiras de fazer isso:</span><span class="sxs-lookup"><span data-stu-id="3c164-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="3c164-p105">Os suplementos podem declarar botões que aparecem nas superfícies de comando em mensagens e compromissos. Para saber mais, confira [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="3c164-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="3c164-120">**Suplemento com botões de comando na Faixa de Opções**</span><span class="sxs-lookup"><span data-stu-id="3c164-120">**An add-in with command buttons on the ribbon**</span></span>

    ![Comando de suplemento de forma sem interface do usuário](../images/uiless-command-shape.png)

- <span data-ttu-id="3c164-p106">Os suplementos podem desvincular correspondências de expressões regulares ou entidades detectadas em mensagens e compromissos. Para saber mais, confira [Suplementos contextuais do Outlook](contextual-outlook-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="3c164-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="3c164-124">**Suplemento contextual para uma entidade realçada (um endereço)**</span><span class="sxs-lookup"><span data-stu-id="3c164-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![Mostra um aplicativo contextual em um cartão](../images/outlook-detected-entity-card.png)

> [!NOTE]
> <span data-ttu-id="3c164-126">[Os painéis personalizados foram preteridos](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), portanto certifique-se de que você está usando um ponto de extensão com suporte.</span><span class="sxs-lookup"><span data-stu-id="3c164-126">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using a supported extension point.</span></span>

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="3c164-127">Itens de caixa de correio disponíveis para suplementos</span><span class="sxs-lookup"><span data-stu-id="3c164-127">Mailbox items available to add-ins</span></span>

<span data-ttu-id="3c164-p107">Os suplementos do Outlook estão disponíveis nas mensagens ou compromissos durante a redação ou leitura, mas não em outros tipos de itens. O Outlook não ativa os suplementos se o item de mensagem atual, em um formato de redação ou de leitura, estiver em uma das seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="3c164-p107">Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="3c164-p108">Protegido por IRM (Gerenciamento de Direitos de Informação) ou criptografado de outras maneiras para proteção. Uma mensagem assinada digitalmente é um exemplo, já que a assinatura digital se baseia em um desses mecanismos.</span><span class="sxs-lookup"><span data-stu-id="3c164-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="3c164-132">A partir do Outlook, compilação 13120.1000, no Windows, os suplementos agora podem ser ativados nos itens protegidos por IRM.</span><span class="sxs-lookup"><span data-stu-id="3c164-132">Starting with Outlook build 13120.1000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="3c164-133">Para obter mais informações sobre esse recurso na visualização, consulte [Ativação de suplementos em itens protegidos pela Gestão de Direitos de Informação (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span><span class="sxs-lookup"><span data-stu-id="3c164-133">For more information about this feature in preview, see [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="3c164-134">Um relatório de entrega ou notificação que tem a classe de mensagem IPM.Report.\*, incluindo NDRs (notificações de falha na entrega) e notificações de leitura, falha na leitura e atraso.</span><span class="sxs-lookup"><span data-stu-id="3c164-134">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="3c164-135">Um rascunho (não tem um remetente atribuído a ele) ou está na pasta Rascunhos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="3c164-135">A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.</span></span>

- <span data-ttu-id="3c164-136">Um arquivo .msg que é um anexo de outra mensagem.</span><span class="sxs-lookup"><span data-stu-id="3c164-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="3c164-137">Um arquivo .msg aberto no sistema de arquivos.</span><span class="sxs-lookup"><span data-stu-id="3c164-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="3c164-138">Em uma caixa de correio compartilhada, na caixa de correio de outro usuário, em uma caixa de correio de arquivo morto ou em uma pasta pública.</span><span class="sxs-lookup"><span data-stu-id="3c164-138">In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.</span></span>

- <span data-ttu-id="3c164-139">Usando um formulário personalizado.</span><span class="sxs-lookup"><span data-stu-id="3c164-139">Using a custom form.</span></span>

<span data-ttu-id="3c164-140">Em geral, o Outlook pode ativar suplementos no formato de leitura para itens na pasta Itens Enviados, com exceção dos suplementos que são ativados baseados em cadeias de correspondências de entidades conhecidas.</span><span class="sxs-lookup"><span data-stu-id="3c164-140">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="3c164-141">Para saber mais sobre os motivos por trás disso, confira "Suporte para entidades conhecidas" em [Corresponder cadeias em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="3c164-141">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-hosts"></a><span data-ttu-id="3c164-142">Hosts compatíveis</span><span class="sxs-lookup"><span data-stu-id="3c164-142">Supported hosts</span></span>

<span data-ttu-id="3c164-143">Suplementos do Outlook são compatíveis com o Outlook 2013 ou posterior no Windows, Outlook 2016 ou posterior no Mac, Outlook na Web para Exchange 2013 no local e versões posteriores, Outlook no iOS, Outlook no Android e Outlook na Web e Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="3c164-143">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="3c164-144">Nem todos os recursos mais recentes são compatíveis com todos os [clientes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="3c164-144">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="3c164-145">Confira os artigos e as referências de API para esses recursos e saiba com quais hosts eles podem ou não ter compatibilidade.</span><span class="sxs-lookup"><span data-stu-id="3c164-145">Please refer to articles and API references for those features to see which hosts they may or may not be supported in.</span></span>


## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="3c164-146">Introdução à criação de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="3c164-146">Get started building Outlook add-ins</span></span>

<span data-ttu-id="3c164-147">Para começar a criar suplementos do Outlook, experimente o seguinte.</span><span class="sxs-lookup"><span data-stu-id="3c164-147">To get started building Outlook add-ins, try the following.</span></span>

- <span data-ttu-id="3c164-148">[Início Rápido](../quickstarts/outlook-quickstart.md) - Criar um painel de tarefas simples.</span><span class="sxs-lookup"><span data-stu-id="3c164-148">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="3c164-149">[Tutorial](../tutorials/outlook-tutorial.md) : saiba como criar um suplemento que insere gists do GitHub em uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="3c164-149">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>


## <a name="see-also"></a><span data-ttu-id="3c164-150">Confira também</span><span class="sxs-lookup"><span data-stu-id="3c164-150">See also</span></span>

- [<span data-ttu-id="3c164-151">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3c164-151">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="3c164-152">Diretrizes de design para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3c164-152">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="3c164-153">Licenciar suplementos do Office e do SharePoint</span><span class="sxs-lookup"><span data-stu-id="3c164-153">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="3c164-154">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="3c164-154">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="3c164-155">Disponibilizar suas soluções no AppSource e no Office</span><span class="sxs-lookup"><span data-stu-id="3c164-155">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
