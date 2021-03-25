---
title: Recurso Ao enviar para suplementos do Outlook
description: Fornece uma maneira de manipular um item ou impedir que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar.
ms.date: 03/17/2021
localization_priority: Normal
ms.openlocfilehash: 70e255601fd36a2f9101d56161846616691f5100
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178052"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="727c4-103">Recurso Ao enviar para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="727c4-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="727c4-p101">O recurso Ao enviar para suplementos do Outlook fornece uma maneira de manipular uma mensagem ou item de reunião, ou impede que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar. Por exemplo, você pode usar o recurso Ao enviar para:</span><span class="sxs-lookup"><span data-stu-id="727c4-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="727c4-106">Impedir que um usuário envie informações confidenciais ou deixe a linha de assunto em branco.</span><span class="sxs-lookup"><span data-stu-id="727c4-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="727c4-107">Adicionar um destinatário específico à linha CC em mensagens ou à linha destinatários opcionais em reuniões.</span><span class="sxs-lookup"><span data-stu-id="727c4-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="727c4-108">O recurso ao enviar é acionado pelo tipo de evento `ItemSend` e é sem interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="727c4-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="727c4-109">Para obter informações sobre limitações relacionadas ao recurso Ao enviar, consulte as [Limitações](#limitations) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="727c4-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="727c4-110">Clientes e plataformas com suporte</span><span class="sxs-lookup"><span data-stu-id="727c4-110">Supported clients and platforms</span></span>

<span data-ttu-id="727c4-111">A tabela a seguir mostra combinações de cliente-servidor com suporte para o recurso ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-111">The following table shows supported client-server combinations for the on-send feature.</span></span> <span data-ttu-id="727c4-112">Não há suporte para combinações excluídas.</span><span class="sxs-lookup"><span data-stu-id="727c4-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="727c4-113">Cliente</span><span class="sxs-lookup"><span data-stu-id="727c4-113">Client</span></span> | <span data-ttu-id="727c4-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="727c4-114">Exchange Online</span></span> | <span data-ttu-id="727c4-115">Exchange 2016 local</span><span class="sxs-lookup"><span data-stu-id="727c4-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="727c4-116">(Atualização Cumulativa 6 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="727c4-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="727c4-117">Exchange 2019 local</span><span class="sxs-lookup"><span data-stu-id="727c4-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="727c4-118">(Atualização Cumulativa 1 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="727c4-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="727c4-119">Windows:</span><span class="sxs-lookup"><span data-stu-id="727c4-119">Windows:</span></span><br><span data-ttu-id="727c4-120">versão 1910 (build 12130.20272) ou posterior</span><span class="sxs-lookup"><span data-stu-id="727c4-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="727c4-121">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-121">Yes</span></span>|<span data-ttu-id="727c4-122">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-122">Yes</span></span>|<span data-ttu-id="727c4-123">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-123">Yes</span></span>|
|<span data-ttu-id="727c4-124">Mac:</span><span class="sxs-lookup"><span data-stu-id="727c4-124">Mac:</span></span><br><span data-ttu-id="727c4-125">build 16.30 ou posterior</span><span class="sxs-lookup"><span data-stu-id="727c4-125">build 16.30 or later</span></span>|<span data-ttu-id="727c4-126">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-126">Yes</span></span>|<span data-ttu-id="727c4-127">Não</span><span class="sxs-lookup"><span data-stu-id="727c4-127">No</span></span>|<span data-ttu-id="727c4-128">Não</span><span class="sxs-lookup"><span data-stu-id="727c4-128">No</span></span>|
|<span data-ttu-id="727c4-129">Navegador da Web:</span><span class="sxs-lookup"><span data-stu-id="727c4-129">Web browser:</span></span><br><span data-ttu-id="727c4-130">interface do usuário moderna do Outlook</span><span class="sxs-lookup"><span data-stu-id="727c4-130">modern Outlook UI</span></span>|<span data-ttu-id="727c4-131">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-131">Yes</span></span>|<span data-ttu-id="727c4-132">Não aplicável</span><span class="sxs-lookup"><span data-stu-id="727c4-132">Not applicable</span></span>|<span data-ttu-id="727c4-133">Não aplicável</span><span class="sxs-lookup"><span data-stu-id="727c4-133">Not applicable</span></span>|
|<span data-ttu-id="727c4-134">Navegador da Web:</span><span class="sxs-lookup"><span data-stu-id="727c4-134">Web browser:</span></span><br><span data-ttu-id="727c4-135">interface do usuário clássica do Outlook</span><span class="sxs-lookup"><span data-stu-id="727c4-135">classic Outlook UI</span></span>|<span data-ttu-id="727c4-136">Não aplicável</span><span class="sxs-lookup"><span data-stu-id="727c4-136">Not applicable</span></span>|<span data-ttu-id="727c4-137">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-137">Yes</span></span>|<span data-ttu-id="727c4-138">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="727c4-139">O recurso ao enviar foi lançado oficialmente no conjunto de requisitos 1.8 (consulte o servidor atual e o suporte ao [cliente](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes).</span><span class="sxs-lookup"><span data-stu-id="727c4-139">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="727c4-140">No entanto, observe que a matriz de suporte do recurso é um superconjunto do conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="727c4-140">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="727c4-141">Os complementos que usam o recurso ao enviar não são permitidos [no AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="727c4-141">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="727c4-142">Como o recurso Ao enviar funciona?</span><span class="sxs-lookup"><span data-stu-id="727c4-142">How does the on-send feature work?</span></span>

<span data-ttu-id="727c4-143">Você pode usar o recurso Ao enviar para criar um suplemento do Outlook que integre o evento síncrono `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="727c4-143">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="727c4-144">Este evento detecta que o usuário está pressionando o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) e pode ser usado para impedir que um item seja enviado se houver falha na validação.</span><span class="sxs-lookup"><span data-stu-id="727c4-144">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="727c4-145">Por exemplo, quando um usuário dispara um evento de envio de mensagem, um suplemento do Outlook que usa o recurso Ao enviar pode:</span><span class="sxs-lookup"><span data-stu-id="727c4-145">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="727c4-146">Ler e validar o conteúdo da mensagem de email</span><span class="sxs-lookup"><span data-stu-id="727c4-146">Read and validate the email message contents</span></span>
- <span data-ttu-id="727c4-147">Verificar se a mensagem inclui uma linha de assunto</span><span class="sxs-lookup"><span data-stu-id="727c4-147">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="727c4-148">Definir um destinatário predeterminado</span><span class="sxs-lookup"><span data-stu-id="727c4-148">Set a predetermined recipient</span></span>

<span data-ttu-id="727c4-149">A validação é feita no lado do cliente no Outlook quando o evento de envio é disparado, e o complemento tem até 5 minutos antes de terminar. Se a validação falhar, o envio do item será bloqueado e uma mensagem de erro será exibida em uma barra de informações que solicita que o usuário tome medidas.</span><span class="sxs-lookup"><span data-stu-id="727c4-149">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="727c4-150">No Outlook na Web, quando o recurso ao enviar é acionado em uma mensagem que está sendo composta na guia navegador do Outlook, o item é lançado para sua própria janela ou guia do navegador para concluir a validação e outros processamentos.</span><span class="sxs-lookup"><span data-stu-id="727c4-150">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="727c4-151">A captura de tela a seguir mostra uma barra de informações que notifica que o remetente adicione um assunto.</span><span class="sxs-lookup"><span data-stu-id="727c4-151">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Captura de tela mostrando uma mensagem de erro solicitando que o usuário insira uma linha de assunto ausente](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="727c4-153">A captura de tela a seguir mostra uma barra de informações que notifica que o remetente de que foram encontradas palavras bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="727c4-153">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Captura de tela mostrando uma mensagem de erro informando ao usuário que foram encontradas palavras bloqueadas](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="727c4-155">Limitações</span><span class="sxs-lookup"><span data-stu-id="727c4-155">Limitations</span></span>

<span data-ttu-id="727c4-156">Atualmente, o recurso Ao enviar tem as seguintes limitações.</span><span class="sxs-lookup"><span data-stu-id="727c4-156">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="727c4-157">**Recurso Append-on-send** &ndash; Se você chamar [corpo. AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) no manipulador ao enviar, um erro é retornado.</span><span class="sxs-lookup"><span data-stu-id="727c4-157">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="727c4-158">**AppSource** &ndash; Você não pode publicar suplementos do Outlook que usem o recurso Ao enviar no [AppSource](https://appsource.microsoft.com), pois eles falharão na validação do AppSource.</span><span class="sxs-lookup"><span data-stu-id="727c4-158">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="727c4-159">Os suplementos que usam o recurso Ao enviar devem ser implantados pelos administradores.</span><span class="sxs-lookup"><span data-stu-id="727c4-159">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="727c4-160">**Manifesto** &ndash; Somente um evento `ItemSend` tem suporte por suplemento.</span><span class="sxs-lookup"><span data-stu-id="727c4-160">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="727c4-161">Se você tiver dois ou mais eventos `ItemSend` em um manifesto, haverá falha na validação.</span><span class="sxs-lookup"><span data-stu-id="727c4-161">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="727c4-p107">**Desempenho**&ndash; Várias idas e voltas ao servidor Web que hospeda o suplemento podem afetar o desempenho do suplemento. Considere os efeitos sobre o desempenho quando você cria suplemento que exigem várias mensagens ou operações baseadas em reuniões.</span><span class="sxs-lookup"><span data-stu-id="727c4-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="727c4-164">**Enviar mais tarde** (somente Mac) &ndash; Se houver suplementos Ao enviar, o recurso **Enviar mais tarde** ficará indisponível.</span><span class="sxs-lookup"><span data-stu-id="727c4-164">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="727c4-165">Além disso, não é recomendável que você chame o manipulador de eventos ao enviar, pois o fechamento do item deve acontecer automaticamente depois que o `item.close()` evento for concluído.</span><span class="sxs-lookup"><span data-stu-id="727c4-165">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="727c4-166">Limitações de tipo/modo de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="727c4-166">Mailbox type/mode limitations</span></span>

<span data-ttu-id="727c4-167">A funcionalidade Ao enviar é compatível apenas com caixas de correio de usuários no Outlook na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="727c4-167">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="727c4-168">Atualmente, a funcionalidade não tem suporte para os seguintes tipos e modos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="727c4-168">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="727c4-169">Caixas de correio compartilhadas\*</span><span class="sxs-lookup"><span data-stu-id="727c4-169">Shared mailboxes\*</span></span>
- <span data-ttu-id="727c4-170">Caixas de correio de grupo</span><span class="sxs-lookup"><span data-stu-id="727c4-170">Group mailboxes</span></span>
- <span data-ttu-id="727c4-171">Modo offline</span><span class="sxs-lookup"><span data-stu-id="727c4-171">Offline mode</span></span>

<span data-ttu-id="727c4-172">O Outlook não permitirá o envio se o recurso Ao enviar estiver habilitado para esses cenários de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="727c4-172">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="727c4-173">No entanto, se um usuário responder a um email em uma caixa de correio de grupo, o suplemento Ao enviar não será executado e a mensagem será enviada.</span><span class="sxs-lookup"><span data-stu-id="727c4-173">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="727c4-174">\*A funcionalidade ao enviar deve funcionar em caixas de correio ou pastas compartilhadas se o complemento também implementar suporte para cenários de acesso [de representante.](delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="727c4-174">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="727c4-175">Vários suplementos Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-175">Multiple on-send add-ins</span></span>

<span data-ttu-id="727c4-176">Se vários suplementos Ao enviar estiverem instalados, os suplementos serão executados na ordem em que são recebidos das APIs `getAppManifestCall` ou `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="727c4-176">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="727c4-177">Se o primeiro suplemento permitir envio, o segundo suplemento poderá alterar algo que faria o primeiro bloquear o envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-177">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="727c4-178">No entanto, o primeiro suplemento não será executado novamente se todos os suplementos instalados tiverem permissão de envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-178">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="727c4-179">Por exemplo, o Suplemento1 e o Suplemento2 usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-179">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="727c4-180">O Suplemento1 é instalado primeiro e o Suplemento2 é instalado depois.</span><span class="sxs-lookup"><span data-stu-id="727c4-180">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="727c4-181">O Suplemento1 verifica se a palavra Fabrikam aparece na mensagem como uma condição para o suplemento permitir o envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-181">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="727c4-182">No entanto, o Suplemento2 remove as ocorrências da palavra Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="727c4-182">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="727c4-183">A mensagem será enviada com todas as instâncias de Fabrikam removidas (devido à ordem de instalação do Suplemento1 e do Suplemento2).</span><span class="sxs-lookup"><span data-stu-id="727c4-183">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="727c4-184">Implantar suplementos do Outlook que usam Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-184">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="727c4-185">Recomendamos que os administradores implantem suplementos do Outlook que usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-185">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="727c4-186">Os administradores precisam garantir que o suplemento Ao enviar:</span><span class="sxs-lookup"><span data-stu-id="727c4-186">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="727c4-187">Esteja sempre presente a qualquer momento que um item de redigir é aberto (para email: novo, responder ou encaminhar).</span><span class="sxs-lookup"><span data-stu-id="727c4-187">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="727c4-188">Não pode ser fechado ou desabilitado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="727c4-188">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="727c4-189">Instalar suplementos do Outlook que usam Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-189">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="727c4-190">O recurso Ao enviar no Outlook exige que os suplementos sejam configurados para os tipos de eventos de envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-190">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="727c4-191">Selecione a plataforma que você deseja configurar.</span><span class="sxs-lookup"><span data-stu-id="727c4-191">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="727c4-192">Navegador da Web – Outlook clássico</span><span class="sxs-lookup"><span data-stu-id="727c4-192">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="727c4-193">Os suplementos para Outlook na Web (clássicos) que usam o recurso Ao enviar serão executados para usuários aos quais é atribuída uma política de caixa de correio do Outlook na Web que tenha o sinalizador *OnSendAddinsEnabled* definido como **true**.</span><span class="sxs-lookup"><span data-stu-id="727c4-193">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="727c4-194">Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="727c4-194">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="727c4-195">Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="727c4-195">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="727c4-196">Habilitar o recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-196">Enable the on-send feature</span></span>

<span data-ttu-id="727c4-197">Por padrão, a funcionalidade Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="727c4-197">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="727c4-198">Os administradores podem habilitar a funcionalidade Ao enviar executando os cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="727c4-198">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="727c4-199">Para habilitar suplementos Ao enviar para todos os usuários:</span><span class="sxs-lookup"><span data-stu-id="727c4-199">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="727c4-200">Criar uma nova política de caixa de correio do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-200">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="727c4-201">Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="727c4-201">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="727c4-202">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-202">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="727c4-203">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-203">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="727c4-204">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="727c4-204">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="727c4-205">Habilitar o recurso Ao enviar para um grupo de usuários</span><span class="sxs-lookup"><span data-stu-id="727c4-205">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="727c4-206">Para habilitar o recurso Ao enviar para um grupo específico de usuários, as etapas são as seguintes.</span><span class="sxs-lookup"><span data-stu-id="727c4-206">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="727c4-207">Neste exemplo, um administrador deseja habilitar apenas o recurso de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finance (em que os usuários do Finance estão no Departamento Financeiro).</span><span class="sxs-lookup"><span data-stu-id="727c4-207">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="727c4-208">Crie uma nova política de caixa de correio do Outlook na Web para o grupo.</span><span class="sxs-lookup"><span data-stu-id="727c4-208">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="727c4-209">Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações).</span><span class="sxs-lookup"><span data-stu-id="727c4-209">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="727c4-210">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-210">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="727c4-211">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-211">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="727c4-212">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="727c4-212">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="727c4-213">Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS).</span><span class="sxs-lookup"><span data-stu-id="727c4-213">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="727c4-214">Quando a política entrar em vigor, o recurso Ao enviar será habilitado para o grupo.</span><span class="sxs-lookup"><span data-stu-id="727c4-214">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="727c4-215">Desabilitar o recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-215">Disable the on-send feature</span></span>

<span data-ttu-id="727c4-216">Para desabilitar o recurso Ao enviar de um usuário ou atribuir uma política de caixa de correio do Outlook na Web que não tenha o sinalizador habilitado, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="727c4-216">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="727c4-217">Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="727c4-217">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="727c4-218">Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="727c4-218">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="727c4-219">Para desabilitar o recurso Ao enviar para todos os usuários que tenham uma política específica de caixa de correio do Outlook na Web atribuída, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="727c4-219">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="727c4-220">Navegador da Web – Outlook moderno</span><span class="sxs-lookup"><span data-stu-id="727c4-220">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="727c4-221">Os suplementos para Outlook na Web (modernos) que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="727c4-221">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="727c4-222">No entanto, se os usuários são obrigados a executar os complementos ao enviar para atender aos padrões de conformidade, a política de caixa de correio deve ter o sinalizador *OnSendAddinsEnabled* definido como para que a edição do item não seja permitida enquanto os complementos estão sendo processadas no `true` envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-222">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="727c4-223">Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="727c4-223">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="727c4-224">Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="727c4-224">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="727c4-225">Habilitar o sinalizador ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-225">Enable the on-send flag</span></span>

<span data-ttu-id="727c4-226">Os administradores podem impor a conformidade ao enviar executando cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="727c4-226">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="727c4-227">Para todos os usuários, não permitir a edição enquanto os complementos ao enviar estão processamento:</span><span class="sxs-lookup"><span data-stu-id="727c4-227">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="727c4-228">Criar uma nova política de caixa de correio do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-228">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="727c4-229">Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="727c4-229">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="727c4-230">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-230">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="727c4-231">Impor a conformidade ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-231">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="727c4-232">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="727c4-232">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="727c4-233">Ativar o sinalizador ao enviar para um grupo de usuários</span><span class="sxs-lookup"><span data-stu-id="727c4-233">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="727c4-234">Para impor a conformidade ao enviar para um grupo específico de usuários, as etapas são as seguintes.</span><span class="sxs-lookup"><span data-stu-id="727c4-234">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="727c4-235">Neste exemplo, um administrador apenas deseja habilitar uma política de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finanças (em que os usuários do Finanças estão no Departamento Financeiro).</span><span class="sxs-lookup"><span data-stu-id="727c4-235">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="727c4-236">Crie uma nova política de caixa de correio do Outlook na Web para o grupo.</span><span class="sxs-lookup"><span data-stu-id="727c4-236">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="727c4-237">Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações).</span><span class="sxs-lookup"><span data-stu-id="727c4-237">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="727c4-238">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-238">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="727c4-239">Impor a conformidade ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-239">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="727c4-240">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="727c4-240">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="727c4-241">Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS).</span><span class="sxs-lookup"><span data-stu-id="727c4-241">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="727c4-242">Quando a política entra em vigor, a conformidade ao enviar será imposta para o grupo.</span><span class="sxs-lookup"><span data-stu-id="727c4-242">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="727c4-243">Desativar o sinalizador ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-243">Turn off the on-send flag</span></span>

<span data-ttu-id="727c4-244">Para desativar a imposição de conformidade ao enviar para um usuário, atribua uma política de caixa de correio do Outlook na Web que não tenha o sinalizador habilitado executando os cmdlets a seguir.</span><span class="sxs-lookup"><span data-stu-id="727c4-244">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="727c4-245">Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="727c4-245">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="727c4-246">Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="727c4-246">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="727c4-247">Para desativar a imposição de conformidade ao enviar para todos os usuários que tenham uma política específica do Outlook na caixa de correio da Web atribuída, execute os cmdlets a seguir.</span><span class="sxs-lookup"><span data-stu-id="727c4-247">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="727c4-248">Windows</span><span class="sxs-lookup"><span data-stu-id="727c4-248">Windows</span></span>](#tab/windows)

<span data-ttu-id="727c4-249">Os suplementos para Outlook no Windows que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="727c4-249">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="727c4-250">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a política de grupo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** deve estar definida como **Habilitada** em cada máquina aplicável.</span><span class="sxs-lookup"><span data-stu-id="727c4-250">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="727c4-251">Para definir políticas de caixa de correio, os administradores podem baixar a ferramenta Modelos [Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) e acessar os modelos administrativos mais recentes executando o Editor de Política de Grupo Local, **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="727c4-251">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="727c4-252">O que a política faz</span><span class="sxs-lookup"><span data-stu-id="727c4-252">What the policy does</span></span>

<span data-ttu-id="727c4-253">Por motivos de conformidade, os administrador podem precisar garantir que os usuários não possam enviar itens de mensagem de reunião até que o último suplemento Ao enviar esteja disponível para execução.</span><span class="sxs-lookup"><span data-stu-id="727c4-253">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="727c4-254">Os administradores devem habilitar a política de grupo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** para que todos os suplementos sejam atualizados a partir do Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-254">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="727c4-255">Status da política</span><span class="sxs-lookup"><span data-stu-id="727c4-255">Policy status</span></span>|<span data-ttu-id="727c4-256">Resultado</span><span class="sxs-lookup"><span data-stu-id="727c4-256">Result</span></span>|
|---|---|
|<span data-ttu-id="727c4-257">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-257">Disabled</span></span>|<span data-ttu-id="727c4-258">Os manifestos baixados atualmente dos complementos ao enviar (não necessariamente as versões mais recentes) são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="727c4-258">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="727c4-259">Esse é o status/comportamento padrão.</span><span class="sxs-lookup"><span data-stu-id="727c4-259">This is the default status/behavior.</span></span>|
|<span data-ttu-id="727c4-260">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-260">Enabled</span></span>|<span data-ttu-id="727c4-261">Depois que os manifestos mais recentes dos complementos ao enviar são baixados do Exchange, os complementos são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="727c4-261">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="727c4-262">Caso contrário, o envio será bloqueado.</span><span class="sxs-lookup"><span data-stu-id="727c4-262">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="727c4-263">Gerenciar a política Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-263">Manage the on-send policy</span></span>

<span data-ttu-id="727c4-264">Por padrão, a política Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="727c4-264">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="727c4-265">Os administradores podem habilitar a política Ao enviar ao certificar-se de que a configuração de política de grupo do usuário **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** esteja definida como **Habilitada**.</span><span class="sxs-lookup"><span data-stu-id="727c4-265">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="727c4-266">Para desabilitar a política para um usuário, o administrador deve defini-la como **Desabilitada**.</span><span class="sxs-lookup"><span data-stu-id="727c4-266">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="727c4-267">Para gerenciar essa configuração de política, você pode fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="727c4-267">To manage this policy setting, you can do the following:</span></span>

1. <span data-ttu-id="727c4-268">Baixe a [ferramenta de Modelos Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) mais recente.</span><span class="sxs-lookup"><span data-stu-id="727c4-268">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="727c4-269">Abra o Editor de Política de Grupo Local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="727c4-269">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="727c4-270">Navegue até **Configuração do Usuário > Modelos Administrativos > Microsoft Outlook 2016 > Segurança > Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="727c4-270">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="727c4-271">Marque a configuração **Desabilitar o envio quando as extensões da Web não puderem ser carregadas**.</span><span class="sxs-lookup"><span data-stu-id="727c4-271">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="727c4-272">Abra o link para configuração Editar política.</span><span class="sxs-lookup"><span data-stu-id="727c4-272">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="727c4-273">Na caixa de diálogo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas**, selecione **Habilitado** ou **Desabilitado** conforme apropriado e selecione **OK** ou **Aplique** para colocar a atualização em vigor.</span><span class="sxs-lookup"><span data-stu-id="727c4-273">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="727c4-274">Mac</span><span class="sxs-lookup"><span data-stu-id="727c4-274">Mac</span></span>](#tab/unix)

<span data-ttu-id="727c4-275">Os suplementos para Outlook no Mac que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="727c4-275">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="727c4-276">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a configuração de caixa de correio a seguir deverá ser aplicada ao computador de cada usuário.</span><span class="sxs-lookup"><span data-stu-id="727c4-276">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="727c4-277">Esta configuração ou chave é compatível com CFPreference. Isso significa que é possível defini-la usando um software de gerenciamento empresarial para Mac, como o Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="727c4-277">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

||<span data-ttu-id="727c4-278">Valor</span><span class="sxs-lookup"><span data-stu-id="727c4-278">Value</span></span>|
|:---|:---|
|<span data-ttu-id="727c4-279">**Domínio**</span><span class="sxs-lookup"><span data-stu-id="727c4-279">**Domain**</span></span>|<span data-ttu-id="727c4-280">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="727c4-280">com.microsoft.outlook</span></span>|
|<span data-ttu-id="727c4-281">**Chave**</span><span class="sxs-lookup"><span data-stu-id="727c4-281">**Key**</span></span>|<span data-ttu-id="727c4-282">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="727c4-282">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="727c4-283">**DataType**</span><span class="sxs-lookup"><span data-stu-id="727c4-283">**DataType**</span></span>|<span data-ttu-id="727c4-284">Booliano</span><span class="sxs-lookup"><span data-stu-id="727c4-284">Boolean</span></span>|
|<span data-ttu-id="727c4-285">**Valores possíveis**</span><span class="sxs-lookup"><span data-stu-id="727c4-285">**Possible values**</span></span>|<span data-ttu-id="727c4-286">falso (padrão)</span><span class="sxs-lookup"><span data-stu-id="727c4-286">false (default)</span></span><br><span data-ttu-id="727c4-287">verdadeiro</span><span class="sxs-lookup"><span data-stu-id="727c4-287">true</span></span>|
|<span data-ttu-id="727c4-288">**Disponibilidade**</span><span class="sxs-lookup"><span data-stu-id="727c4-288">**Availability**</span></span>|<span data-ttu-id="727c4-289">16.27</span><span class="sxs-lookup"><span data-stu-id="727c4-289">16.27</span></span>|
|<span data-ttu-id="727c4-290">**Comentários**</span><span class="sxs-lookup"><span data-stu-id="727c4-290">**Comments**</span></span>|<span data-ttu-id="727c4-291">Essa chave cria uma política de onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="727c4-291">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="727c4-292">O que a configuração faz</span><span class="sxs-lookup"><span data-stu-id="727c4-292">What the setting does</span></span>

<span data-ttu-id="727c4-293">Por motivos de conformidade, os administradores podem precisar garantir que os usuários não possam enviar itens de mensagem ou de reunião até que os suplementos estejam disponíveis para execução.</span><span class="sxs-lookup"><span data-stu-id="727c4-293">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="727c4-294">Os administradores devem habilitar a chave **OnSendAddinsWaitForLoad** para que todos os suplementos sejam atualizados no Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-294">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="727c4-295">Estado da chave</span><span class="sxs-lookup"><span data-stu-id="727c4-295">Key's state</span></span>|<span data-ttu-id="727c4-296">Resultado</span><span class="sxs-lookup"><span data-stu-id="727c4-296">Result</span></span>|
|---|---|
|<span data-ttu-id="727c4-297">falso</span><span class="sxs-lookup"><span data-stu-id="727c4-297">false</span></span>|<span data-ttu-id="727c4-298">Os manifestos baixados atualmente dos complementos ao enviar (não necessariamente as versões mais recentes) são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="727c4-298">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="727c4-299">Esse é o estado/comportamento padrão.</span><span class="sxs-lookup"><span data-stu-id="727c4-299">This is the default state/behavior.</span></span>|
|<span data-ttu-id="727c4-300">verdadeiro</span><span class="sxs-lookup"><span data-stu-id="727c4-300">true</span></span>|<span data-ttu-id="727c4-301">Depois que os manifestos mais recentes dos complementos ao enviar são baixados do Exchange, os complementos são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="727c4-301">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="727c4-302">Caso contrário, o envio será bloqueado e **o botão Enviar** será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="727c4-302">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="727c4-303">Cenários do recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="727c4-303">On-send feature scenarios</span></span>

<span data-ttu-id="727c4-304">Veja a seguir os cenários com suporte e sem suporte para suplementos que usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-304">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="727c4-305">A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, mas nenhum suplemento está instalado</span><span class="sxs-lookup"><span data-stu-id="727c4-305">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="727c4-306">Neste cenário, o usuário poderá enviar itens de mensagem e de reunião sem nenhum suplemento em execução.</span><span class="sxs-lookup"><span data-stu-id="727c4-306">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="727c4-307">A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, e os suplementos compatíveis com Ao enviar estão instalados e habilitados</span><span class="sxs-lookup"><span data-stu-id="727c4-307">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="727c4-308">Os suplementos serão executados durante o evento de envio, que em seguida permitirão ou impedirão o usuário de enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-308">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="727c4-309">Delegação de caixa de correio, onde a caixa de correio 1 tem permissões de acesso total à caixa de correio 2</span><span class="sxs-lookup"><span data-stu-id="727c4-309">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="727c4-310">Navegador da Web (Outlook clássico)</span><span class="sxs-lookup"><span data-stu-id="727c4-310">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="727c4-311">Cenário</span><span class="sxs-lookup"><span data-stu-id="727c4-311">Scenario</span></span>|<span data-ttu-id="727c4-312">Recurso Ao enviar da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="727c4-312">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="727c4-313">Recurso Ao enviar da caixa de correio 2</span><span class="sxs-lookup"><span data-stu-id="727c4-313">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="727c4-314">Sessão Web do Outlook (clássico)</span><span class="sxs-lookup"><span data-stu-id="727c4-314">Outlook web session (classic)</span></span>|<span data-ttu-id="727c4-315">Resultado</span><span class="sxs-lookup"><span data-stu-id="727c4-315">Result</span></span>|<span data-ttu-id="727c4-316">Com suporte?</span><span class="sxs-lookup"><span data-stu-id="727c4-316">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="727c4-317">1</span><span class="sxs-lookup"><span data-stu-id="727c4-317">1</span></span>|<span data-ttu-id="727c4-318">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-318">Enabled</span></span>|<span data-ttu-id="727c4-319">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-319">Enabled</span></span>|<span data-ttu-id="727c4-320">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="727c4-320">New session</span></span>|<span data-ttu-id="727c4-321">A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.</span><span class="sxs-lookup"><span data-stu-id="727c4-321">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="727c4-p135">Não há suporte atualmente. Como alternativa, use o cenário 3.</span><span class="sxs-lookup"><span data-stu-id="727c4-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="727c4-324">2</span><span class="sxs-lookup"><span data-stu-id="727c4-324">2</span></span>|<span data-ttu-id="727c4-325">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-325">Disabled</span></span>|<span data-ttu-id="727c4-326">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-326">Enabled</span></span>|<span data-ttu-id="727c4-327">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="727c4-327">New session</span></span>|<span data-ttu-id="727c4-328">A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.</span><span class="sxs-lookup"><span data-stu-id="727c4-328">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="727c4-p136">Não há suporte atualmente. Como alternativa, use o cenário 3.</span><span class="sxs-lookup"><span data-stu-id="727c4-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="727c4-331">3</span><span class="sxs-lookup"><span data-stu-id="727c4-331">3</span></span>|<span data-ttu-id="727c4-332">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-332">Enabled</span></span>|<span data-ttu-id="727c4-333">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-333">Enabled</span></span>|<span data-ttu-id="727c4-334">Mesma sessão</span><span class="sxs-lookup"><span data-stu-id="727c4-334">Same session</span></span>|<span data-ttu-id="727c4-335">Os suplementos Ao enviar atribuídos à caixa de correio 1 são executados ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-335">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="727c4-336">Com suporte.</span><span class="sxs-lookup"><span data-stu-id="727c4-336">Supported.</span></span>|
|<span data-ttu-id="727c4-337">4 </span><span class="sxs-lookup"><span data-stu-id="727c4-337">4</span></span>|<span data-ttu-id="727c4-338">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-338">Enabled</span></span>|<span data-ttu-id="727c4-339">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-339">Disabled</span></span>|<span data-ttu-id="727c4-340">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="727c4-340">New session</span></span>|<span data-ttu-id="727c4-341">Nenhum suplemento Ao envio é executado; item de mensagem ou de reunião é enviado.</span><span class="sxs-lookup"><span data-stu-id="727c4-341">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="727c4-342">Com suporte.</span><span class="sxs-lookup"><span data-stu-id="727c4-342">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="727c4-343">Navegador da Web (Outlook moderno), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="727c4-343">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="727c4-344">Para impor o Ao enviar, os administradores devem garantir que a política tenha sido habilitada nas duas caixas de correio.</span><span class="sxs-lookup"><span data-stu-id="727c4-344">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="727c4-345">Para saber como oferecer suporte ao acesso de representante em um suplemento, confira [Habilitar cenários de acesso de representante em um suplemento do Outlook](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="727c4-345">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="727c4-346">O grupo 1 é uma caixa de correio do grupo moderna e a caixa de correio 1 do usuário é membro do grupo 1</span><span class="sxs-lookup"><span data-stu-id="727c4-346">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="727c4-347">Cenário</span><span class="sxs-lookup"><span data-stu-id="727c4-347">Scenario</span></span>|<span data-ttu-id="727c4-348">Política Ao enviar da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="727c4-348">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="727c4-349">Suplementos Ao enviar habilitados?</span><span class="sxs-lookup"><span data-stu-id="727c4-349">On-send add-ins enabled?</span></span>|<span data-ttu-id="727c4-350">Ação da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="727c4-350">Mailbox 1 action</span></span>|<span data-ttu-id="727c4-351">Resultado</span><span class="sxs-lookup"><span data-stu-id="727c4-351">Result</span></span>|<span data-ttu-id="727c4-352">Com suporte?</span><span class="sxs-lookup"><span data-stu-id="727c4-352">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="727c4-353">1</span><span class="sxs-lookup"><span data-stu-id="727c4-353">1</span></span>|<span data-ttu-id="727c4-354">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-354">Enabled</span></span>|<span data-ttu-id="727c4-355">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-355">Yes</span></span>|<span data-ttu-id="727c4-356">A caixa de correio 1 compõe uma nova mensagem ou reunião para o grupo 1.</span><span class="sxs-lookup"><span data-stu-id="727c4-356">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="727c4-357">Os suplementos Ao enviar são executados durante o envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-357">On-send add-ins run during send.</span></span>|<span data-ttu-id="727c4-358">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-358">Yes</span></span>|
|<span data-ttu-id="727c4-359">2</span><span class="sxs-lookup"><span data-stu-id="727c4-359">2</span></span>|<span data-ttu-id="727c4-360">Habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-360">Enabled</span></span>|<span data-ttu-id="727c4-361">Sim</span><span class="sxs-lookup"><span data-stu-id="727c4-361">Yes</span></span>|<span data-ttu-id="727c4-362">A caixa de correio 1 compõe uma nova mensagem ou reunião para o grupo 1 dentro da janela de grupo do grupo 1 no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="727c4-362">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="727c4-363">Os suplementos Ao enviar não são executados durante o envio.</span><span class="sxs-lookup"><span data-stu-id="727c4-363">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="727c4-364">Não há suporte atualmente.</span><span class="sxs-lookup"><span data-stu-id="727c4-364">Not currently supported.</span></span> <span data-ttu-id="727c4-365">Como alternativa, use o cenário 1.</span><span class="sxs-lookup"><span data-stu-id="727c4-365">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="727c4-366">Caixa de correio do usuário com recurso/política de suplemento Ao enviar habilitado, os suplementos com suporte à funcionalidade Ao enviar estão instalados e habilitados e o modo offline está habilitado</span><span class="sxs-lookup"><span data-stu-id="727c4-366">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="727c4-367">Os suplementos Ao enviar serão executados de acordo com o estado online do usuário, o back-end do suplemento e o Exchange.</span><span class="sxs-lookup"><span data-stu-id="727c4-367">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="727c4-368">Estado do usuário</span><span class="sxs-lookup"><span data-stu-id="727c4-368">User's state</span></span>

<span data-ttu-id="727c4-369">Os suplementos Ao enviar serão executados durante o envio se o usuário estiver online.</span><span class="sxs-lookup"><span data-stu-id="727c4-369">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="727c4-370">Se o usuário estiver offline, os suplementos Ao enviar não serão executados e o item de mensagem ou de reunião não será enviado.</span><span class="sxs-lookup"><span data-stu-id="727c4-370">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="727c4-371">Estado do back-end do suplemento</span><span class="sxs-lookup"><span data-stu-id="727c4-371">Add-in backend's state</span></span>

<span data-ttu-id="727c4-372">Um suplemento Ao enviar será executado se o seu back-end estiver online e acessível.</span><span class="sxs-lookup"><span data-stu-id="727c4-372">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="727c4-373">Se o back-end estiver offline, ao enviar será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="727c4-373">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="727c4-374">Estado do Exchange</span><span class="sxs-lookup"><span data-stu-id="727c4-374">Exchange's state</span></span>

<span data-ttu-id="727c4-375">Os suplementos Ao enviar serão executados durante o envio se o servidor do Exchange estiver online e acessível.</span><span class="sxs-lookup"><span data-stu-id="727c4-375">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="727c4-376">Se o suplemento Ao enviar não puder alcançar o Exchange e a política ou cmdlet aplicável estiverem ativados, o envio será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="727c4-376">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="727c4-377">No Mac, em qualquer estado offline, o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) está desabilitado e uma notificação é exibida informando que sua organização não permite envio quando o usuário está offline.</span><span class="sxs-lookup"><span data-stu-id="727c4-377">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="727c4-378">O usuário pode editar o item enquanto os complementos ao enviar estão trabalhando nele</span><span class="sxs-lookup"><span data-stu-id="727c4-378">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="727c4-379">Enquanto os complementos ao enviar estão processamento de um item, o usuário pode editar o item adicionando, por exemplo, texto ou anexos inadequados.</span><span class="sxs-lookup"><span data-stu-id="727c4-379">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="727c4-380">Se você quiser impedir que o usuário edite o item enquanto o seu add-in está processamento no envio, você pode implementar uma solução alternativa usando uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="727c4-380">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="727c4-381">Essa solução alternativa pode ser usada no Outlook na Web (clássico), Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="727c4-381">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="727c4-382">Outlook moderno na Web: para impedir que o usuário edite o item enquanto o seu complemento está sendo processada no envio, defina o sinalizador *OnSendAddinsEnabled* como conforme descrito na seção Instalar os complementos do Outlook que usam ao enviar anteriormente `true` neste artigo. [](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send)</span><span class="sxs-lookup"><span data-stu-id="727c4-382">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="727c4-383">No manipulador ao enviar:</span><span class="sxs-lookup"><span data-stu-id="727c4-383">In your on-send handler:</span></span>

1. <span data-ttu-id="727c4-384">Chame [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) para abrir uma caixa de diálogo para que os cliques do mouse e os teclas sejam desabilitados.</span><span class="sxs-lookup"><span data-stu-id="727c4-384">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="727c4-385">Para obter esse comportamento no Outlook clássico na Web, você deve definir a propriedade [displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) como no `true` parâmetro da `options` `displayDialogAsync` chamada.</span><span class="sxs-lookup"><span data-stu-id="727c4-385">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="727c4-386">Implemente o processamento do item.</span><span class="sxs-lookup"><span data-stu-id="727c4-386">Implement processing of the item.</span></span>
1. <span data-ttu-id="727c4-387">Feche a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="727c4-387">Close the dialog.</span></span> <span data-ttu-id="727c4-388">Além disso, manipular o que acontece se o usuário fechar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="727c4-388">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="727c4-389">Exemplos de código</span><span class="sxs-lookup"><span data-stu-id="727c4-389">Code examples</span></span>

<span data-ttu-id="727c4-390">Os seguintes exemplos de código mostram como criar um suplemento simples Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-390">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="727c4-391">Para baixar o exemplo de código em que esses exemplos se baseiam, consulte [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="727c4-391">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="727c4-392">Se você usar uma caixa de diálogo com o evento ao enviar, certifique-se de fechar a caixa de diálogo antes de concluir o evento.</span><span class="sxs-lookup"><span data-stu-id="727c4-392">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="727c4-393">Manifesto, versão de substituição e evento</span><span class="sxs-lookup"><span data-stu-id="727c4-393">Manifest, version override, and event</span></span>

<span data-ttu-id="727c4-394">Um exemplo de código [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) inclui dois manifestos:</span><span class="sxs-lookup"><span data-stu-id="727c4-394">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="727c4-395">`Contoso Message Body Checker.xml` &ndash; Mostra como verificar se o corpo de uma mensagem apresenta palavras restritas ou informações confidenciais ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-395">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="727c4-396">`Contoso Subject and CC Checker.xml` &ndash; Mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui uma linha de assunto ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-396">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="727c4-397">No arquivo de manifesto `Contoso Message Body Checker.xml`, inclua o arquivo de função e o nome da função que deve ser chamada no evento `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="727c4-397">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="727c4-398">A operação é executada de maneira síncrona.</span><span class="sxs-lookup"><span data-stu-id="727c4-398">The operation runs synchronously.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> <span data-ttu-id="727c4-399">Se você estiver usando o Visual Studio 2019 para desenvolver seu complemento ao enviar, poderá receber um aviso de validação como o seguinte: "Este é um xsi:type ' http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events 'inválido". Para resolver isso, você precisará de uma versão mais recente do MailAppVersionOverridesV1_1.xsd que tenha sido fornecida como um gist do GitHub em um blog sobre [esse aviso](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span><span class="sxs-lookup"><span data-stu-id="727c4-399">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="727c4-400">Para o arquivo de manifesto `Contoso Subject and CC Checker.xml`, o exemplo a seguir mostra o arquivo de função e o nome da função para chamar o evento de envio de mensagem.</span><span class="sxs-lookup"><span data-stu-id="727c4-400">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

<br/>

<span data-ttu-id="727c4-401">A API Ao enviar requer `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="727c4-401">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="727c4-402">Veja a seguir como adicionar o nó `VersionOverrides` em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="727c4-402">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="727c4-403">Para obter mais informações, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="727c4-403">For more information, see the following:</span></span>
> - [<span data-ttu-id="727c4-404">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="727c4-404">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="727c4-405">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="727c4-405">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="727c4-406">Objetos `Event` e `item`, e os métodos `body.getAsync` e `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="727c4-406">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="727c4-407">Para acessar o item de mensagem ou de reunião selecionado no momento (neste exemplo, a mensagem redigida recentemente), use o namespace `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="727c4-407">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="727c4-408">O evento `ItemSend` é passado automaticamente pelo recurso Ao enviar para a função especificada no manifesto&mdash;neste exemplo, a função `validateBody`.</span><span class="sxs-lookup"><span data-stu-id="727c4-408">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

```js
var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

<span data-ttu-id="727c4-409">A função `validateBody` obtém o corpo atual no formato especificado (HTML) e passa o objeto de evento `ItemSend` que o código deseja para acessar o método de retorno.</span><span class="sxs-lookup"><span data-stu-id="727c4-409">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="727c4-410">Além do método `getAsync`, o objeto `Body` também fornece um método `setAsync` que você pode usar para substituir o corpo pelo texto especificado.</span><span class="sxs-lookup"><span data-stu-id="727c4-410">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="727c4-411">Para saber mais, confira [Objeto do Evento](/javascript/api/office/office.addincommands.event) e [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="727c4-411">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="727c4-412">Objeto `NotificationMessages` e método `event.completed`</span><span class="sxs-lookup"><span data-stu-id="727c4-412">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="727c4-413">A função `checkBodyOnlyOnSendCallBack` usa uma expressão regular para determinar se o corpo da mensagem contém palavras bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="727c4-413">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="727c4-414">Se ela encontrar uma correspondência com uma matriz de palavras restritas, bloqueará os emails de serem enviados e notificará o remetente pela barra de informações.</span><span class="sxs-lookup"><span data-stu-id="727c4-414">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="727c4-415">Para fazer isso, ele usa a propriedade `notificationMessages` do objeto `Item` para retornar um objeto `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="727c4-415">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="727c4-416">Ele, em seguida, adiciona uma notificação ao item chamando o método `addAsync`, como mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="727c4-416">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

<span data-ttu-id="727c4-417">Veja a seguir os parâmetros para o método `addAsync`:</span><span class="sxs-lookup"><span data-stu-id="727c4-417">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="727c4-418">`NoSend` &ndash; uma cadeia de caractere que é uma chave especificada pelo desenvolvedor para fazer referência a uma mensagem de notificação.</span><span class="sxs-lookup"><span data-stu-id="727c4-418">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="727c4-419">Você pode usá-la para modificar esta mensagem mais tarde.</span><span class="sxs-lookup"><span data-stu-id="727c4-419">You can use it to modify this message later.</span></span> <span data-ttu-id="727c4-420">A chave não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="727c4-420">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="727c4-421">`type` &ndash; uma das propriedades do parâmetro de objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="727c4-421">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="727c4-422">Representa o tipo de uma mensagem; os tipos correspondem aos valores da enumeração [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="727c4-422">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="727c4-423">Os valores possíveis são indicador de progresso, mensagem informativa ou mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="727c4-423">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="727c4-424">Neste exemplo, `type` é uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="727c4-424">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="727c4-425">`message` &ndash; uma das propriedades do parâmetro de objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="727c4-425">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="727c4-426">Neste exemplo, `message` é o texto da mensagem de notificação.</span><span class="sxs-lookup"><span data-stu-id="727c4-426">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="727c4-427">Para sinalizar que o suplemento terminou de processar o evento `ItemSend` disparado pela operação enviar, chame o método `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="727c4-427">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="727c4-428">A propriedade `allowEvent` é um booleano.</span><span class="sxs-lookup"><span data-stu-id="727c4-428">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="727c4-429">Se for definido como `true`, o envio será permitido.</span><span class="sxs-lookup"><span data-stu-id="727c4-429">If set to `true`, send is allowed.</span></span> <span data-ttu-id="727c4-430">Se definido como `false`, a mensagem de email será impedida de ser enviada.</span><span class="sxs-lookup"><span data-stu-id="727c4-430">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="727c4-431">Para saber mais, confira [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="727c4-431">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="727c4-432">Métodos `replaceAsync`, `removeAsync` e `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="727c4-432">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="727c4-433">Além do método `addAsync`, o objeto `NotificationMessages` também inclui os métodos `replaceAsync`, `removeAsync` e `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="727c4-433">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="727c4-434">Esses métodos não são usados neste exemplo de código.</span><span class="sxs-lookup"><span data-stu-id="727c4-434">These methods are not used in this code sample.</span></span>  <span data-ttu-id="727c4-435">Para saber mais, veja [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="727c4-435">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="727c4-436">Código do Assunto e do verificador de CC</span><span class="sxs-lookup"><span data-stu-id="727c4-436">Subject and CC checker code</span></span>

<span data-ttu-id="727c4-437">O exemplo de código a seguir mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui um assunto ao enviar.</span><span class="sxs-lookup"><span data-stu-id="727c4-437">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="727c4-438">Este exemplo usa o recurso Ao enviar para permitir ou proibir o envio de um email.</span><span class="sxs-lookup"><span data-stu-id="727c4-438">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

<span data-ttu-id="727c4-p156">Para saber mais sobre como adicionar um destinatário à linha CC e verificar se a mensagem de e-mail inclui uma linha de assunto ao enviar e para ver as APIs que você pode usar, consulte o [exemplo Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). O código é bem comentado.</span><span class="sxs-lookup"><span data-stu-id="727c4-p156">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="727c4-441">Confira também</span><span class="sxs-lookup"><span data-stu-id="727c4-441">See also</span></span>

- [<span data-ttu-id="727c4-442">Visão geral da arquitetura e dos recursos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="727c4-442">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="727c4-443">Suplemento do Outlook para demonstração de comando de suplemento</span><span class="sxs-lookup"><span data-stu-id="727c4-443">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)