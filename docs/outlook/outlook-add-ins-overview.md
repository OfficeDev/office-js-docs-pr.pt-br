---
title: Visão geral dos suplementos do Outlook
description: Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web.
ms.date: 10/09/2019
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: cb6e19788390a804b0bbacb97666a3ca8a9d5971
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165805"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="ab68e-103">Visão geral dos suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="ab68e-103">Outlook add-ins overview</span></span>

<span data-ttu-id="ab68e-104">Os suplementos do Outlook são integrações criadas por terceiros para o Outlook usando nossa plataforma baseada na Web.</span><span class="sxs-lookup"><span data-stu-id="ab68e-104">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform.</span></span> <span data-ttu-id="ab68e-105">Os suplementos do Outlook têm três aspectos principais:</span><span class="sxs-lookup"><span data-stu-id="ab68e-105">Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="ab68e-106">A mesma lógica de suplemento e de negócios funciona na área de trabalho (Outlook no Windows e Mac), na Web (Office 365 e Outlook.com) e em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="ab68e-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Office 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="ab68e-107">Os suplementos do Outlook consistem em um manifesto, que descreve como o suplemento se integra ao Outlook (por exemplo, um botão ou um painel de tarefas), e o código JavaScript/HTML, que compõe a interface do usuário e lógica de negócios do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ab68e-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="ab68e-108">Os suplementos do Outlook podem ser adquiridos na [AppSource](https://appsource.microsoft.com) ou [sideloaded](sideload-outlook-add-ins-for-testing.md) por usuários finais ou administradores.</span><span class="sxs-lookup"><span data-stu-id="ab68e-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="ab68e-109">Os suplementos do Outlook são diferentes dos suplementos de COM ou VSTO, que são integrações mais antigas específicas do Outlook para Windows.</span><span class="sxs-lookup"><span data-stu-id="ab68e-109">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows.</span></span> <span data-ttu-id="ab68e-110">Diferentemente dos suplementos de COM, os suplementos do Outlook não têm qualquer código fisicamente instalado no dispositivo do usuário ou no cliente do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ab68e-110">Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client.</span></span> <span data-ttu-id="ab68e-111">No caso de um suplemento do Outlook, o Outlook lê o manifesto, conecta os controles especificados na interface do usuário e carrega o HTML e o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ab68e-111">For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML.</span></span> <span data-ttu-id="ab68e-112">Todos os componentes Web são executados no contexto do navegador em uma área restrita.</span><span class="sxs-lookup"><span data-stu-id="ab68e-112">The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="ab68e-113">Os itens do Outlook que dão suporte a suplementos incluem mensagens de email, compromissos, solicitações, respostas e cancelamentos de reunião.</span><span class="sxs-lookup"><span data-stu-id="ab68e-113">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments.</span></span> <span data-ttu-id="ab68e-114">Cada suplemento do Outlook define o contexto no qual está disponível, incluindo os tipos de itens e se o usuário está lendo ou redigindo um item.</span><span class="sxs-lookup"><span data-stu-id="ab68e-114">Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

> [!NOTE]
> <span data-ttu-id="ab68e-p104">Caso pretenda [publicar](../publish/publish.md) o suplemento no AppSource depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade do suplemento do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="ab68e-p104">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="extension-points"></a><span data-ttu-id="ab68e-117">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ab68e-117">Extension points</span></span>

<span data-ttu-id="ab68e-p105">Pontos de extensão são as formas usadas pelos suplementos para se integrar ao Outlook. Estas são as maneiras de fazer isso:</span><span class="sxs-lookup"><span data-stu-id="ab68e-p105">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="ab68e-p106">Os suplementos podem declarar botões que aparecem nas superfícies de comando em mensagens e compromissos. Para saber mais, confira [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="ab68e-p106">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="ab68e-122">**Suplemento com botões de comando na Faixa de Opções**</span><span class="sxs-lookup"><span data-stu-id="ab68e-122">**An add-in with command buttons on the ribbon**</span></span>

    ![Comando de suplemento de forma sem interface do usuário](../images/uiless-command-shape.png)

- <span data-ttu-id="ab68e-p107">Os suplementos podem desvincular correspondências de expressões regulares ou entidades detectadas em mensagens e compromissos. Para saber mais, confira [Suplementos contextuais do Outlook](contextual-outlook-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="ab68e-p107">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="ab68e-126">**Suplemento contextual para uma entidade realçada (um endereço)**</span><span class="sxs-lookup"><span data-stu-id="ab68e-126">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![Mostra um aplicativo contextual em um cartão](../images/outlook-detected-entity-card.png)


> [!NOTE]
> <span data-ttu-id="ab68e-128">[Os painéis personalizados foram preteridos](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), portanto certifique-se de que você está usando um ponto de extensão com suporte.</span><span class="sxs-lookup"><span data-stu-id="ab68e-128">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using a supported extension point.</span></span>

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="ab68e-129">Itens de caixa de correio disponíveis para suplementos</span><span class="sxs-lookup"><span data-stu-id="ab68e-129">Mailbox items available to add-ins</span></span>

<span data-ttu-id="ab68e-p108">Os suplementos do Outlook estão disponíveis nas mensagens ou compromissos durante a redação ou leitura, mas não em outros tipos de itens. O Outlook não ativa os suplementos se o item de mensagem atual, em um formato de redação ou de leitura, estiver em uma das seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="ab68e-p108">Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="ab68e-p109">Protegido por IRM (Gerenciamento de Direitos de Informação) ou criptografado de outras maneiras para proteção. Uma mensagem assinada digitalmente é um exemplo, já que a assinatura digital se baseia em um desses mecanismos.</span><span class="sxs-lookup"><span data-stu-id="ab68e-p109">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

- <span data-ttu-id="ab68e-134">Um relatório de entrega ou notificação que tem a classe de mensagem IPM.Report.\*, incluindo NDRs (notificações de falha na entrega) e notificações de leitura, falha na leitura e atraso.</span><span class="sxs-lookup"><span data-stu-id="ab68e-134">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="ab68e-135">Um rascunho (não tem um remetente atribuído a ele) ou está na pasta Rascunhos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ab68e-135">A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.</span></span>

- <span data-ttu-id="ab68e-136">Um arquivo .msg que é um anexo de outra mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab68e-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="ab68e-137">Um arquivo .msg aberto no sistema de arquivos.</span><span class="sxs-lookup"><span data-stu-id="ab68e-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="ab68e-138">Em uma caixa de correio compartilhada, na caixa de correio de outro usuário, em uma caixa de correio de arquivo morto ou em uma pasta pública.</span><span class="sxs-lookup"><span data-stu-id="ab68e-138">In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.</span></span>

- <span data-ttu-id="ab68e-139">Usando um formulário personalizado.</span><span class="sxs-lookup"><span data-stu-id="ab68e-139">Using a custom form.</span></span>

<span data-ttu-id="ab68e-140">Em geral, o Outlook pode ativar suplementos no formato de leitura para itens na pasta Itens Enviados, com exceção dos suplementos que são ativados baseados em cadeias de correspondências de entidades conhecidas.</span><span class="sxs-lookup"><span data-stu-id="ab68e-140">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="ab68e-141">Para saber mais sobre os motivos por trás disso, confira "Suporte para entidades conhecidas" em [Corresponder cadeias em um item do Outlook como entidades conhecidas](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="ab68e-141">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-hosts"></a><span data-ttu-id="ab68e-142">Hosts compatíveis</span><span class="sxs-lookup"><span data-stu-id="ab68e-142">Supported hosts</span></span>

<span data-ttu-id="ab68e-143">Suplementos do Outlook são compatíveis com o Outlook 2013 ou posterior no Windows, Outlook 2016 ou posterior no Mac, Outlook na Web para Exchange 2013 no local e versões posteriores, Outlook no iOS, Outlook no Android e Outlook na Web no Office 365 e Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="ab68e-143">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web in Office 365 and Outlook.com.</span></span> <span data-ttu-id="ab68e-144">Nem todos os recursos mais recentes são compatíveis com todos os [clientes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="ab68e-144">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="ab68e-145">Confira os artigos e as referências de API para esses recursos e saiba com quais hosts eles podem ou não ter compatibilidade.</span><span class="sxs-lookup"><span data-stu-id="ab68e-145">Please refer to articles and API references for those features to see which hosts they may or may not be supported in.</span></span>


## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="ab68e-146">Introdução à criação de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="ab68e-146">Get started building Outlook add-ins</span></span>

<span data-ttu-id="ab68e-147">Para começar a criar suplementos do Outlook, experimente o seguinte.</span><span class="sxs-lookup"><span data-stu-id="ab68e-147">To get started building Outlook add-ins, try the following.</span></span>

- <span data-ttu-id="ab68e-148">[Início Rápido](../quickstarts/outlook-quickstart.md) - Criar um painel de tarefas simples.</span><span class="sxs-lookup"><span data-stu-id="ab68e-148">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="ab68e-149">[Tutorial](../tutorials/outlook-tutorial.md) : saiba como criar um suplemento que insere gists do GitHub em uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab68e-149">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>


## <a name="see-also"></a><span data-ttu-id="ab68e-150">Confira também</span><span class="sxs-lookup"><span data-stu-id="ab68e-150">See also</span></span>

- [<span data-ttu-id="ab68e-151">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ab68e-151">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="ab68e-152">Diretrizes de design para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ab68e-152">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="ab68e-153">Licenciar suplementos do Office e do SharePoint</span><span class="sxs-lookup"><span data-stu-id="ab68e-153">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="ab68e-154">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="ab68e-154">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="ab68e-155">Disponibilizar suas soluções no AppSource e no Office</span><span class="sxs-lookup"><span data-stu-id="ab68e-155">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
