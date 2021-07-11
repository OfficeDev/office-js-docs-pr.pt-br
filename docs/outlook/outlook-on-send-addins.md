---
title: Recurso Ao enviar para suplementos do Outlook
description: Fornece uma maneira de manipular um item ou impedir que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar.
ms.date: 06/16/2021
localization_priority: Normal
ms.openlocfilehash: 80047f4c8056bafa62d467f1e69dd334d168486a
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348472"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="59c61-103">Recurso Ao enviar para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="59c61-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="59c61-p101">O recurso Ao enviar para suplementos do Outlook fornece uma maneira de manipular uma mensagem ou item de reunião, ou impede que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar. Por exemplo, você pode usar o recurso Ao enviar para:</span><span class="sxs-lookup"><span data-stu-id="59c61-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="59c61-106">Impedir que um usuário envie informações confidenciais ou deixe a linha de assunto em branco.</span><span class="sxs-lookup"><span data-stu-id="59c61-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="59c61-107">Adicionar um destinatário específico à linha CC em mensagens ou à linha destinatários opcionais em reuniões.</span><span class="sxs-lookup"><span data-stu-id="59c61-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="59c61-108">O recurso ao enviar é acionado pelo tipo de evento `ItemSend` e é sem interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="59c61-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="59c61-109">Para obter informações sobre limitações relacionadas ao recurso Ao enviar, consulte as [Limitações](#limitations) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="59c61-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="59c61-110">Clientes e plataformas com suporte</span><span class="sxs-lookup"><span data-stu-id="59c61-110">Supported clients and platforms</span></span>

<span data-ttu-id="59c61-111">A tabela a seguir mostra combinações de cliente-servidor com suporte para o recurso ao enviar, incluindo a Atualização Cumulativa mínima necessária quando aplicável.</span><span class="sxs-lookup"><span data-stu-id="59c61-111">The following table shows supported client-server combinations for the on-send feature, including the minimum required Cumulative Update where applicable.</span></span> <span data-ttu-id="59c61-112">Não há suporte para combinações excluídas.</span><span class="sxs-lookup"><span data-stu-id="59c61-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="59c61-113">Client</span><span class="sxs-lookup"><span data-stu-id="59c61-113">Client</span></span> | <span data-ttu-id="59c61-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="59c61-114">Exchange Online</span></span> | <span data-ttu-id="59c61-115">Exchange 2016 local</span><span class="sxs-lookup"><span data-stu-id="59c61-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="59c61-116">(Atualização Cumulativa 6 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="59c61-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="59c61-117">Exchange 2019 local</span><span class="sxs-lookup"><span data-stu-id="59c61-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="59c61-118">(Atualização Cumulativa 1 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="59c61-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="59c61-119">Windows:</span><span class="sxs-lookup"><span data-stu-id="59c61-119">Windows:</span></span><br><span data-ttu-id="59c61-120">versão 1910 (build 12130.20272) ou posterior</span><span class="sxs-lookup"><span data-stu-id="59c61-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="59c61-121">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-121">Yes</span></span>|<span data-ttu-id="59c61-122">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-122">Yes</span></span>|<span data-ttu-id="59c61-123">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-123">Yes</span></span>|
|<span data-ttu-id="59c61-124">Mac:</span><span class="sxs-lookup"><span data-stu-id="59c61-124">Mac:</span></span><br><span data-ttu-id="59c61-125">build 16.47 ou posterior</span><span class="sxs-lookup"><span data-stu-id="59c61-125">build 16.47 or later</span></span>|<span data-ttu-id="59c61-126">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-126">Yes</span></span>|<span data-ttu-id="59c61-127">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-127">Yes</span></span>|<span data-ttu-id="59c61-128">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-128">Yes</span></span>|
|<span data-ttu-id="59c61-129">Navegador da Web:</span><span class="sxs-lookup"><span data-stu-id="59c61-129">Web browser:</span></span><br><span data-ttu-id="59c61-130">interface do usuário Outlook moderna</span><span class="sxs-lookup"><span data-stu-id="59c61-130">modern Outlook UI</span></span>|<span data-ttu-id="59c61-131">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-131">Yes</span></span>|<span data-ttu-id="59c61-132">Não aplicável</span><span class="sxs-lookup"><span data-stu-id="59c61-132">Not applicable</span></span>|<span data-ttu-id="59c61-133">Não aplicável</span><span class="sxs-lookup"><span data-stu-id="59c61-133">Not applicable</span></span>|
|<span data-ttu-id="59c61-134">Navegador da Web:</span><span class="sxs-lookup"><span data-stu-id="59c61-134">Web browser:</span></span><br><span data-ttu-id="59c61-135">interface do usuário Outlook clássica</span><span class="sxs-lookup"><span data-stu-id="59c61-135">classic Outlook UI</span></span>|<span data-ttu-id="59c61-136">Não aplicável</span><span class="sxs-lookup"><span data-stu-id="59c61-136">Not applicable</span></span>|<span data-ttu-id="59c61-137">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-137">Yes</span></span>|<span data-ttu-id="59c61-138">Sim</span><span class="sxs-lookup"><span data-stu-id="59c61-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="59c61-139">O recurso ao enviar foi lançado oficialmente no conjunto de requisitos 1.8 (consulte o servidor atual e o suporte ao [cliente](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes).</span><span class="sxs-lookup"><span data-stu-id="59c61-139">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="59c61-140">No entanto, observe que a matriz de suporte do recurso é um superconjunto do conjunto de requisitos.</span><span class="sxs-lookup"><span data-stu-id="59c61-140">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="59c61-141">Os complementos que usam o recurso ao enviar não são permitidos [no AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="59c61-141">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="59c61-142">Como o recurso Ao enviar funciona?</span><span class="sxs-lookup"><span data-stu-id="59c61-142">How does the on-send feature work?</span></span>

<span data-ttu-id="59c61-143">Você pode usar o recurso Ao enviar para criar um suplemento do Outlook que integre o evento síncrono `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="59c61-143">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="59c61-144">Este evento detecta que o usuário está pressionando o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) e pode ser usado para impedir que um item seja enviado se houver falha na validação.</span><span class="sxs-lookup"><span data-stu-id="59c61-144">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="59c61-145">Por exemplo, quando um usuário dispara um evento de envio de mensagem, um suplemento do Outlook que usa o recurso Ao enviar pode:</span><span class="sxs-lookup"><span data-stu-id="59c61-145">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="59c61-146">Ler e validar o conteúdo da mensagem de email</span><span class="sxs-lookup"><span data-stu-id="59c61-146">Read and validate the email message contents</span></span>
- <span data-ttu-id="59c61-147">Verificar se a mensagem inclui uma linha de assunto</span><span class="sxs-lookup"><span data-stu-id="59c61-147">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="59c61-148">Definir um destinatário predeterminado</span><span class="sxs-lookup"><span data-stu-id="59c61-148">Set a predetermined recipient</span></span>

<span data-ttu-id="59c61-149">A validação é feita no lado do cliente Outlook quando o evento de envio é acionado e o complemento tem até 5 minutos antes do tempo final. Se a validação falhar, o envio do item será bloqueado e uma mensagem de erro será exibida em uma barra de informações que solicita que o usuário tome medidas.</span><span class="sxs-lookup"><span data-stu-id="59c61-149">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="59c61-150">No Outlook na Web, quando o recurso ao enviar é acionado em uma mensagem que está sendo composta na guia navegador Outlook, o item é lançado para sua própria janela ou guia do navegador para concluir a validação e outros processamentos.</span><span class="sxs-lookup"><span data-stu-id="59c61-150">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="59c61-151">A captura de tela a seguir mostra uma barra de informações que notifica que o remetente adicione um assunto.</span><span class="sxs-lookup"><span data-stu-id="59c61-151">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Captura de tela mostrando uma mensagem de erro solicitando que o usuário insira uma linha de assunto ausente.](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="59c61-153">A captura de tela a seguir mostra uma barra de informações que notifica que o remetente de que foram encontradas palavras bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="59c61-153">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Captura de tela mostrando uma mensagem de erro dizendo ao usuário que palavras bloqueadas foram encontradas.](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="59c61-155">Limitações</span><span class="sxs-lookup"><span data-stu-id="59c61-155">Limitations</span></span>

<span data-ttu-id="59c61-156">Atualmente, o recurso Ao enviar tem as seguintes limitações.</span><span class="sxs-lookup"><span data-stu-id="59c61-156">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="59c61-157">**Recurso Append-on-send** &ndash; Se você chamar [corpo. AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) no manipulador ao enviar, um erro é retornado.</span><span class="sxs-lookup"><span data-stu-id="59c61-157">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="59c61-158">**AppSource** &ndash; Você não pode publicar suplementos do Outlook que usem o recurso Ao enviar no [AppSource](https://appsource.microsoft.com), pois eles falharão na validação do AppSource.</span><span class="sxs-lookup"><span data-stu-id="59c61-158">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="59c61-159">Os suplementos que usam o recurso Ao enviar devem ser implantados pelos administradores.</span><span class="sxs-lookup"><span data-stu-id="59c61-159">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="59c61-160">**Manifesto** &ndash; Somente um evento `ItemSend` tem suporte por suplemento.</span><span class="sxs-lookup"><span data-stu-id="59c61-160">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="59c61-161">Se você tiver dois ou mais eventos `ItemSend` em um manifesto, haverá falha na validação.</span><span class="sxs-lookup"><span data-stu-id="59c61-161">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="59c61-p107">**Desempenho**&ndash; Várias idas e voltas ao servidor Web que hospeda o suplemento podem afetar o desempenho do suplemento. Considere os efeitos sobre o desempenho quando você cria suplemento que exigem várias mensagens ou operações baseadas em reuniões.</span><span class="sxs-lookup"><span data-stu-id="59c61-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="59c61-164">**Enviar mais tarde** (somente Mac) &ndash; Se houver suplementos Ao enviar, o recurso **Enviar mais tarde** ficará indisponível.</span><span class="sxs-lookup"><span data-stu-id="59c61-164">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="59c61-165">Além disso, não é recomendável que você chame o manipulador de eventos ao enviar, pois o fechamento do item deve acontecer automaticamente depois que o `item.close()` evento for concluído.</span><span class="sxs-lookup"><span data-stu-id="59c61-165">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="59c61-166">Limitações de tipo/modo de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="59c61-166">Mailbox type/mode limitations</span></span>

<span data-ttu-id="59c61-167">A funcionalidade Ao enviar é compatível apenas com caixas de correio de usuários no Outlook na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="59c61-167">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="59c61-168">Além de situações em que os complementos não são ativados conforme o que é notado nos itens de Caixa de Correio disponíveis para os [complementos](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) da página de visão geral de complementos do Outlook, a funcionalidade não tem suporte no momento para o modo offline.</span><span class="sxs-lookup"><span data-stu-id="59c61-168">In addition to situations where add-ins don't activate as noted in the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page, the functionality is not currently supported for offline mode.</span></span>

<span data-ttu-id="59c61-169">Outlook permitirá o envio se o recurso ao enviar estiver habilitado para cenários de caixa de correio sem suporte.</span><span class="sxs-lookup"><span data-stu-id="59c61-169">Outlook won't allow sending if the on-send feature is enabled for unsupported mailbox scenarios.</span></span> <span data-ttu-id="59c61-170">No entanto, em casos em que os Outlook não são ativados, o complemento ao enviar não será executado e a mensagem será enviada.</span><span class="sxs-lookup"><span data-stu-id="59c61-170">However, in cases where Outlook add-ins don't activate, the on-send add-in won't run and the message will be sent.</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="59c61-171">Vários suplementos Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-171">Multiple on-send add-ins</span></span>

<span data-ttu-id="59c61-172">Se vários suplementos Ao enviar estiverem instalados, os suplementos serão executados na ordem em que são recebidos das APIs `getAppManifestCall` ou `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="59c61-172">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="59c61-173">Se o primeiro suplemento permitir envio, o segundo suplemento poderá alterar algo que faria o primeiro bloquear o envio.</span><span class="sxs-lookup"><span data-stu-id="59c61-173">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="59c61-174">No entanto, o primeiro suplemento não será executado novamente se todos os suplementos instalados tiverem permissão de envio.</span><span class="sxs-lookup"><span data-stu-id="59c61-174">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="59c61-175">Por exemplo, o Suplemento1 e o Suplemento2 usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-175">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="59c61-176">O Suplemento1 é instalado primeiro e o Suplemento2 é instalado depois.</span><span class="sxs-lookup"><span data-stu-id="59c61-176">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="59c61-177">O Suplemento1 verifica se a palavra Fabrikam aparece na mensagem como uma condição para o suplemento permitir o envio.</span><span class="sxs-lookup"><span data-stu-id="59c61-177">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="59c61-178">No entanto, o Suplemento2 remove as ocorrências da palavra Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="59c61-178">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="59c61-179">A mensagem será enviada com todas as instâncias de Fabrikam removidas (devido à ordem de instalação do Suplemento1 e do Suplemento2).</span><span class="sxs-lookup"><span data-stu-id="59c61-179">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="59c61-180">Implantar suplementos do Outlook que usam Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-180">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="59c61-181">Recomendamos que os administradores implantem suplementos do Outlook que usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-181">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="59c61-182">Os administradores precisam garantir que o suplemento Ao enviar:</span><span class="sxs-lookup"><span data-stu-id="59c61-182">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="59c61-183">Esteja sempre presente a qualquer momento que um item de redigir é aberto (para email: novo, responder ou encaminhar).</span><span class="sxs-lookup"><span data-stu-id="59c61-183">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="59c61-184">Não pode ser fechado ou desabilitado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="59c61-184">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="59c61-185">Instalar suplementos do Outlook que usam Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-185">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="59c61-186">O recurso Ao enviar no Outlook exige que os suplementos sejam configurados para os tipos de eventos de envio.</span><span class="sxs-lookup"><span data-stu-id="59c61-186">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="59c61-187">Selecione a plataforma que você deseja configurar.</span><span class="sxs-lookup"><span data-stu-id="59c61-187">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="59c61-188">Navegador da Web – Outlook clássico</span><span class="sxs-lookup"><span data-stu-id="59c61-188">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="59c61-189">Os suplementos para Outlook na Web (clássicos) que usam o recurso Ao enviar serão executados para usuários aos quais é atribuída uma política de caixa de correio do Outlook na Web que tenha o sinalizador *OnSendAddinsEnabled* definido como **true**.</span><span class="sxs-lookup"><span data-stu-id="59c61-189">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="59c61-190">Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="59c61-190">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="59c61-191">Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="59c61-191">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="59c61-192">Habilitar o recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-192">Enable the on-send feature</span></span>

<span data-ttu-id="59c61-193">Por padrão, a funcionalidade Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="59c61-193">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="59c61-194">Os administradores podem habilitar a funcionalidade Ao enviar executando os cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="59c61-194">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="59c61-195">Para habilitar suplementos Ao enviar para todos os usuários:</span><span class="sxs-lookup"><span data-stu-id="59c61-195">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="59c61-196">Criar uma nova política de caixa de correio do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="59c61-196">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="59c61-197">Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="59c61-197">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="59c61-198">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="59c61-198">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="59c61-199">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-199">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="59c61-200">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="59c61-200">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="59c61-201">Habilitar o recurso Ao enviar para um grupo de usuários</span><span class="sxs-lookup"><span data-stu-id="59c61-201">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="59c61-202">Para habilitar o recurso Ao enviar para um grupo específico de usuários, as etapas são as seguintes.</span><span class="sxs-lookup"><span data-stu-id="59c61-202">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="59c61-203">Neste exemplo, um administrador deseja habilitar apenas o recurso de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finance (em que os usuários do Finance estão no Departamento Financeiro).</span><span class="sxs-lookup"><span data-stu-id="59c61-203">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="59c61-204">Crie uma nova política de caixa de correio do Outlook na Web para o grupo.</span><span class="sxs-lookup"><span data-stu-id="59c61-204">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="59c61-205">Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações).</span><span class="sxs-lookup"><span data-stu-id="59c61-205">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="59c61-206">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="59c61-206">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="59c61-207">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-207">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="59c61-208">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="59c61-208">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="59c61-209">Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS).</span><span class="sxs-lookup"><span data-stu-id="59c61-209">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="59c61-210">Quando a política entrar em vigor, o recurso Ao enviar será habilitado para o grupo.</span><span class="sxs-lookup"><span data-stu-id="59c61-210">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="59c61-211">Desabilitar o recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-211">Disable the on-send feature</span></span>

<span data-ttu-id="59c61-212">Para desabilitar o recurso Ao enviar de um usuário ou atribuir uma política de caixa de correio do Outlook na Web que não tenha o sinalizador habilitado, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="59c61-212">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="59c61-213">Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="59c61-213">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="59c61-214">Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="59c61-214">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="59c61-215">Para desabilitar o recurso Ao enviar para todos os usuários que tenham uma política específica de caixa de correio do Outlook na Web atribuída, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="59c61-215">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="59c61-216">Navegador da Web – Outlook moderno</span><span class="sxs-lookup"><span data-stu-id="59c61-216">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="59c61-217">Os suplementos para Outlook na Web (modernos) que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="59c61-217">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="59c61-218">No entanto, se os usuários são obrigados a executar os complementos ao enviar para atender aos padrões de conformidade, a política de caixa de correio deve ter o sinalizador *OnSendAddinsEnabled* definido como para que a edição do item não seja permitida enquanto os complementos estão sendo processadas no `true` envio.</span><span class="sxs-lookup"><span data-stu-id="59c61-218">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="59c61-219">Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="59c61-219">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="59c61-220">Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="59c61-220">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="59c61-221">Habilitar o sinalizador ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-221">Enable the on-send flag</span></span>

<span data-ttu-id="59c61-222">Os administradores podem impor a conformidade ao enviar executando Exchange Online cmdlets do PowerShell.</span><span class="sxs-lookup"><span data-stu-id="59c61-222">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="59c61-223">Para todos os usuários, não permitir a edição enquanto os complementos ao enviar estão processamento:</span><span class="sxs-lookup"><span data-stu-id="59c61-223">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="59c61-224">Criar uma nova política de caixa de correio do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="59c61-224">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="59c61-225">Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="59c61-225">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="59c61-226">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="59c61-226">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="59c61-227">Impor a conformidade ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-227">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="59c61-228">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="59c61-228">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="59c61-229">Ativar o sinalizador ao enviar para um grupo de usuários</span><span class="sxs-lookup"><span data-stu-id="59c61-229">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="59c61-230">Para impor a conformidade ao enviar para um grupo específico de usuários, as etapas são as seguintes.</span><span class="sxs-lookup"><span data-stu-id="59c61-230">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="59c61-231">Neste exemplo, um administrador apenas deseja habilitar uma política de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finanças (em que os usuários do Finanças estão no Departamento Financeiro).</span><span class="sxs-lookup"><span data-stu-id="59c61-231">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="59c61-232">Crie uma nova política de caixa de correio do Outlook na Web para o grupo.</span><span class="sxs-lookup"><span data-stu-id="59c61-232">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="59c61-233">Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações).</span><span class="sxs-lookup"><span data-stu-id="59c61-233">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="59c61-234">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="59c61-234">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="59c61-235">Impor a conformidade ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-235">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="59c61-236">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="59c61-236">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="59c61-237">Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS).</span><span class="sxs-lookup"><span data-stu-id="59c61-237">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="59c61-238">Quando a política entra em vigor, a conformidade ao enviar será imposta para o grupo.</span><span class="sxs-lookup"><span data-stu-id="59c61-238">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="59c61-239">Desativar o sinalizador ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-239">Turn off the on-send flag</span></span>

<span data-ttu-id="59c61-240">Para desativar a imposição de conformidade ao enviar para um usuário, atribua uma política de caixa de correio Outlook na Web que não tenha o sinalizador habilitado executando os cmdlets a seguir.</span><span class="sxs-lookup"><span data-stu-id="59c61-240">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="59c61-241">Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="59c61-241">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="59c61-242">Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="59c61-242">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="59c61-243">Para desativar a imposição de conformidade ao enviar para todos os usuários que tenham uma política de caixa de correio Outlook na Web específica atribuída, execute os cmdlets a seguir.</span><span class="sxs-lookup"><span data-stu-id="59c61-243">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="59c61-244">Windows</span><span class="sxs-lookup"><span data-stu-id="59c61-244">Windows</span></span>](#tab/windows)

<span data-ttu-id="59c61-245">Os suplementos para Outlook no Windows que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="59c61-245">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="59c61-246">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a política de grupo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** deve estar definida como **Habilitada** em cada máquina aplicável.</span><span class="sxs-lookup"><span data-stu-id="59c61-246">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="59c61-247">Para definir políticas de caixa de correio, os administradores podem baixar a ferramenta Modelos [Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) e acessar os modelos administrativos mais recentes executando o Editor de Política de Grupo Local, **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="59c61-247">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="59c61-248">O que a política faz</span><span class="sxs-lookup"><span data-stu-id="59c61-248">What the policy does</span></span>

<span data-ttu-id="59c61-249">Por motivos de conformidade, os administrador podem precisar garantir que os usuários não possam enviar itens de mensagem de reunião até que o último suplemento Ao enviar esteja disponível para execução.</span><span class="sxs-lookup"><span data-stu-id="59c61-249">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="59c61-250">Os administradores devem habilitar a política de grupo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** para que todos os suplementos sejam atualizados a partir do Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-250">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="59c61-251">Status da política</span><span class="sxs-lookup"><span data-stu-id="59c61-251">Policy status</span></span>|<span data-ttu-id="59c61-252">Resultado</span><span class="sxs-lookup"><span data-stu-id="59c61-252">Result</span></span>|
|---|---|
|<span data-ttu-id="59c61-253">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-253">Disabled</span></span>|<span data-ttu-id="59c61-254">Os manifestos baixados atualmente dos complementos ao enviar (não necessariamente as versões mais recentes) são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="59c61-254">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="59c61-255">Esse é o status/comportamento padrão.</span><span class="sxs-lookup"><span data-stu-id="59c61-255">This is the default status/behavior.</span></span>|
|<span data-ttu-id="59c61-256">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-256">Enabled</span></span>|<span data-ttu-id="59c61-257">Depois que os manifestos mais recentes dos complementos ao enviar são baixados do Exchange, os complementos são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="59c61-257">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="59c61-258">Caso contrário, o envio será bloqueado.</span><span class="sxs-lookup"><span data-stu-id="59c61-258">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="59c61-259">Gerenciar a política Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-259">Manage the on-send policy</span></span>

<span data-ttu-id="59c61-260">Por padrão, a política Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="59c61-260">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="59c61-261">Os administradores podem habilitar a política Ao enviar ao certificar-se de que a configuração de política de grupo do usuário **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** esteja definida como **Habilitada**.</span><span class="sxs-lookup"><span data-stu-id="59c61-261">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="59c61-262">Para desabilitar a política para um usuário, o administrador deve defini-la como **Desabilitada**.</span><span class="sxs-lookup"><span data-stu-id="59c61-262">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="59c61-263">Para gerenciar essa configuração de política, você pode fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="59c61-263">To manage this policy setting, you can do the following:</span></span>

1. <span data-ttu-id="59c61-264">Baixe a [ferramenta de Modelos Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) mais recente.</span><span class="sxs-lookup"><span data-stu-id="59c61-264">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="59c61-265">Abra o Editor de Política de Grupo Local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="59c61-265">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="59c61-266">Navegue até **Configuração do Usuário > Modelos Administrativos > Microsoft Outlook 2016 > Segurança > Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="59c61-266">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="59c61-267">Marque a configuração **Desabilitar o envio quando as extensões da Web não puderem ser carregadas**.</span><span class="sxs-lookup"><span data-stu-id="59c61-267">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="59c61-268">Abra o link para configuração Editar política.</span><span class="sxs-lookup"><span data-stu-id="59c61-268">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="59c61-269">Na caixa de diálogo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas**, selecione **Habilitado** ou **Desabilitado** conforme apropriado e selecione **OK** ou **Aplique** para colocar a atualização em vigor.</span><span class="sxs-lookup"><span data-stu-id="59c61-269">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="59c61-270">Mac</span><span class="sxs-lookup"><span data-stu-id="59c61-270">Mac</span></span>](#tab/unix)

<span data-ttu-id="59c61-271">Os suplementos para Outlook no Mac que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="59c61-271">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="59c61-272">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a configuração de caixa de correio a seguir deverá ser aplicada ao computador de cada usuário.</span><span class="sxs-lookup"><span data-stu-id="59c61-272">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="59c61-273">Esta configuração ou chave é compatível com CFPreference. Isso significa que é possível defini-la usando um software de gerenciamento empresarial para Mac, como o Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="59c61-273">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

||<span data-ttu-id="59c61-274">Valor</span><span class="sxs-lookup"><span data-stu-id="59c61-274">Value</span></span>|
|:---|:---|
|<span data-ttu-id="59c61-275">**Domínio**</span><span class="sxs-lookup"><span data-stu-id="59c61-275">**Domain**</span></span>|<span data-ttu-id="59c61-276">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="59c61-276">com.microsoft.outlook</span></span>|
|<span data-ttu-id="59c61-277">**Chave**</span><span class="sxs-lookup"><span data-stu-id="59c61-277">**Key**</span></span>|<span data-ttu-id="59c61-278">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="59c61-278">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="59c61-279">**DataType**</span><span class="sxs-lookup"><span data-stu-id="59c61-279">**DataType**</span></span>|<span data-ttu-id="59c61-280">Booliano</span><span class="sxs-lookup"><span data-stu-id="59c61-280">Boolean</span></span>|
|<span data-ttu-id="59c61-281">**Valores possíveis**</span><span class="sxs-lookup"><span data-stu-id="59c61-281">**Possible values**</span></span>|<span data-ttu-id="59c61-282">falso (padrão)</span><span class="sxs-lookup"><span data-stu-id="59c61-282">false (default)</span></span><br><span data-ttu-id="59c61-283">verdadeiro</span><span class="sxs-lookup"><span data-stu-id="59c61-283">true</span></span>|
|<span data-ttu-id="59c61-284">**Disponibilidade**</span><span class="sxs-lookup"><span data-stu-id="59c61-284">**Availability**</span></span>|<span data-ttu-id="59c61-285">16.27</span><span class="sxs-lookup"><span data-stu-id="59c61-285">16.27</span></span>|
|<span data-ttu-id="59c61-286">**Comentários**</span><span class="sxs-lookup"><span data-stu-id="59c61-286">**Comments**</span></span>|<span data-ttu-id="59c61-287">Essa chave cria uma política de onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="59c61-287">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="59c61-288">O que a configuração faz</span><span class="sxs-lookup"><span data-stu-id="59c61-288">What the setting does</span></span>

<span data-ttu-id="59c61-289">Por motivos de conformidade, os administradores podem precisar garantir que os usuários não possam enviar itens de mensagem ou de reunião até que os suplementos estejam disponíveis para execução.</span><span class="sxs-lookup"><span data-stu-id="59c61-289">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="59c61-290">Os administradores devem habilitar a chave **OnSendAddinsWaitForLoad** para que todos os suplementos sejam atualizados no Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-290">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="59c61-291">Estado da chave</span><span class="sxs-lookup"><span data-stu-id="59c61-291">Key's state</span></span>|<span data-ttu-id="59c61-292">Resultado</span><span class="sxs-lookup"><span data-stu-id="59c61-292">Result</span></span>|
|---|---|
|<span data-ttu-id="59c61-293">falso</span><span class="sxs-lookup"><span data-stu-id="59c61-293">false</span></span>|<span data-ttu-id="59c61-294">Os manifestos baixados atualmente dos complementos ao enviar (não necessariamente as versões mais recentes) são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="59c61-294">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="59c61-295">Esse é o estado/comportamento padrão.</span><span class="sxs-lookup"><span data-stu-id="59c61-295">This is the default state/behavior.</span></span>|
|<span data-ttu-id="59c61-296">verdadeiro</span><span class="sxs-lookup"><span data-stu-id="59c61-296">true</span></span>|<span data-ttu-id="59c61-297">Depois que os manifestos mais recentes dos complementos ao enviar são baixados do Exchange, os complementos são executados em itens de mensagem ou reunião que estão sendo enviados.</span><span class="sxs-lookup"><span data-stu-id="59c61-297">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="59c61-298">Caso contrário, o envio será bloqueado e **o botão Enviar** será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="59c61-298">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="59c61-299">Cenários do recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="59c61-299">On-send feature scenarios</span></span>

<span data-ttu-id="59c61-300">Veja a seguir os cenários com suporte e sem suporte para suplementos que usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-300">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="59c61-301">A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, mas nenhum suplemento está instalado</span><span class="sxs-lookup"><span data-stu-id="59c61-301">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="59c61-302">Neste cenário, o usuário poderá enviar itens de mensagem e de reunião sem nenhum suplemento em execução.</span><span class="sxs-lookup"><span data-stu-id="59c61-302">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="59c61-303">A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, e os suplementos compatíveis com Ao enviar estão instalados e habilitados</span><span class="sxs-lookup"><span data-stu-id="59c61-303">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="59c61-304">Os suplementos serão executados durante o evento de envio, que em seguida permitirão ou impedirão o usuário de enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-304">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="59c61-305">Delegação de caixa de correio, onde a caixa de correio 1 tem permissões de acesso total à caixa de correio 2</span><span class="sxs-lookup"><span data-stu-id="59c61-305">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="59c61-306">Navegador da Web (Outlook clássico)</span><span class="sxs-lookup"><span data-stu-id="59c61-306">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="59c61-307">Cenário</span><span class="sxs-lookup"><span data-stu-id="59c61-307">Scenario</span></span>|<span data-ttu-id="59c61-308">Recurso Ao enviar da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="59c61-308">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="59c61-309">Recurso Ao enviar da caixa de correio 2</span><span class="sxs-lookup"><span data-stu-id="59c61-309">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="59c61-310">Sessão Web do Outlook (clássico)</span><span class="sxs-lookup"><span data-stu-id="59c61-310">Outlook web session (classic)</span></span>|<span data-ttu-id="59c61-311">Resultado</span><span class="sxs-lookup"><span data-stu-id="59c61-311">Result</span></span>|<span data-ttu-id="59c61-312">Com suporte?</span><span class="sxs-lookup"><span data-stu-id="59c61-312">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="59c61-313">1</span><span class="sxs-lookup"><span data-stu-id="59c61-313">1</span></span>|<span data-ttu-id="59c61-314">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-314">Enabled</span></span>|<span data-ttu-id="59c61-315">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-315">Enabled</span></span>|<span data-ttu-id="59c61-316">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="59c61-316">New session</span></span>|<span data-ttu-id="59c61-317">A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.</span><span class="sxs-lookup"><span data-stu-id="59c61-317">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="59c61-p135">Não há suporte atualmente. Como alternativa, use o cenário 3.</span><span class="sxs-lookup"><span data-stu-id="59c61-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="59c61-320">2</span><span class="sxs-lookup"><span data-stu-id="59c61-320">2</span></span>|<span data-ttu-id="59c61-321">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-321">Disabled</span></span>|<span data-ttu-id="59c61-322">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-322">Enabled</span></span>|<span data-ttu-id="59c61-323">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="59c61-323">New session</span></span>|<span data-ttu-id="59c61-324">A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.</span><span class="sxs-lookup"><span data-stu-id="59c61-324">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="59c61-p136">Não há suporte atualmente. Como alternativa, use o cenário 3.</span><span class="sxs-lookup"><span data-stu-id="59c61-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="59c61-327">3</span><span class="sxs-lookup"><span data-stu-id="59c61-327">3</span></span>|<span data-ttu-id="59c61-328">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-328">Enabled</span></span>|<span data-ttu-id="59c61-329">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-329">Enabled</span></span>|<span data-ttu-id="59c61-330">Mesma sessão</span><span class="sxs-lookup"><span data-stu-id="59c61-330">Same session</span></span>|<span data-ttu-id="59c61-331">Os suplementos Ao enviar atribuídos à caixa de correio 1 são executados ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-331">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="59c61-332">Com suporte.</span><span class="sxs-lookup"><span data-stu-id="59c61-332">Supported.</span></span>|
|<span data-ttu-id="59c61-333">4 </span><span class="sxs-lookup"><span data-stu-id="59c61-333">4</span></span>|<span data-ttu-id="59c61-334">Habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-334">Enabled</span></span>|<span data-ttu-id="59c61-335">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-335">Disabled</span></span>|<span data-ttu-id="59c61-336">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="59c61-336">New session</span></span>|<span data-ttu-id="59c61-337">Nenhum suplemento Ao envio é executado; item de mensagem ou de reunião é enviado.</span><span class="sxs-lookup"><span data-stu-id="59c61-337">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="59c61-338">Com suporte.</span><span class="sxs-lookup"><span data-stu-id="59c61-338">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="59c61-339">Navegador da Web (Outlook moderno), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="59c61-339">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="59c61-340">Para impor o Ao enviar, os administradores devem garantir que a política tenha sido habilitada nas duas caixas de correio.</span><span class="sxs-lookup"><span data-stu-id="59c61-340">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="59c61-341">Para saber como dar suporte ao acesso de representante em um complemento, consulte [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="59c61-341">To learn how to support delegate access in an add-in, see [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span></span>

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="59c61-342">Caixa de correio do usuário com recurso/política de suplemento Ao enviar habilitado, os suplementos com suporte à funcionalidade Ao enviar estão instalados e habilitados e o modo offline está habilitado</span><span class="sxs-lookup"><span data-stu-id="59c61-342">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="59c61-343">Os suplementos Ao enviar serão executados de acordo com o estado online do usuário, o back-end do suplemento e o Exchange.</span><span class="sxs-lookup"><span data-stu-id="59c61-343">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="59c61-344">Estado do usuário</span><span class="sxs-lookup"><span data-stu-id="59c61-344">User's state</span></span>

<span data-ttu-id="59c61-345">Os suplementos Ao enviar serão executados durante o envio se o usuário estiver online.</span><span class="sxs-lookup"><span data-stu-id="59c61-345">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="59c61-346">Se o usuário estiver offline, os suplementos Ao enviar não serão executados e o item de mensagem ou de reunião não será enviado.</span><span class="sxs-lookup"><span data-stu-id="59c61-346">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="59c61-347">Estado do back-end do suplemento</span><span class="sxs-lookup"><span data-stu-id="59c61-347">Add-in backend's state</span></span>

<span data-ttu-id="59c61-348">Um suplemento Ao enviar será executado se o seu back-end estiver online e acessível.</span><span class="sxs-lookup"><span data-stu-id="59c61-348">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="59c61-349">Se o back-end estiver offline, ao enviar será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="59c61-349">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="59c61-350">Estado do Exchange</span><span class="sxs-lookup"><span data-stu-id="59c61-350">Exchange's state</span></span>

<span data-ttu-id="59c61-351">Os suplementos Ao enviar serão executados durante o envio se o servidor do Exchange estiver online e acessível.</span><span class="sxs-lookup"><span data-stu-id="59c61-351">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="59c61-352">Se o suplemento Ao enviar não puder alcançar o Exchange e a política ou cmdlet aplicável estiverem ativados, o envio será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="59c61-352">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="59c61-353">No Mac, em qualquer estado offline, o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) está desabilitado e uma notificação é exibida informando que sua organização não permite envio quando o usuário está offline.</span><span class="sxs-lookup"><span data-stu-id="59c61-353">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="59c61-354">O usuário pode editar o item enquanto os complementos ao enviar estão trabalhando nele</span><span class="sxs-lookup"><span data-stu-id="59c61-354">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="59c61-355">Enquanto os complementos ao enviar estão processamento de um item, o usuário pode editar o item adicionando, por exemplo, texto ou anexos inadequados.</span><span class="sxs-lookup"><span data-stu-id="59c61-355">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="59c61-356">Se você quiser impedir que o usuário edite o item enquanto o seu add-in está processamento no envio, você pode implementar uma solução alternativa usando uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="59c61-356">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="59c61-357">Essa solução alternativa pode ser usada em Outlook na Web (clássico), Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="59c61-357">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="59c61-358">Modern Outlook na Web: para impedir que o usuário edite o item enquanto o seu complemento está sendo processada no envio, você deve definir o sinalizador *OnSendAddinsEnabled* como conforme descrito na seção Instalar os complementos do Outlook que usam a seção ao enviar anteriormente `true` neste artigo. [](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send)</span><span class="sxs-lookup"><span data-stu-id="59c61-358">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="59c61-359">No manipulador ao enviar:</span><span class="sxs-lookup"><span data-stu-id="59c61-359">In your on-send handler:</span></span>

1. <span data-ttu-id="59c61-360">Chame [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) para abrir uma caixa de diálogo para que os cliques do mouse e os teclas sejam desabilitados.</span><span class="sxs-lookup"><span data-stu-id="59c61-360">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="59c61-361">Para obter esse comportamento em Outlook na Web, você deve definir a [propriedade displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) como no `true` parâmetro da `options` `displayDialogAsync` chamada.</span><span class="sxs-lookup"><span data-stu-id="59c61-361">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="59c61-362">Implemente o processamento do item.</span><span class="sxs-lookup"><span data-stu-id="59c61-362">Implement processing of the item.</span></span>
1. <span data-ttu-id="59c61-363">Feche a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="59c61-363">Close the dialog.</span></span> <span data-ttu-id="59c61-364">Além disso, manipular o que acontece se o usuário fechar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="59c61-364">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="59c61-365">Exemplos de código</span><span class="sxs-lookup"><span data-stu-id="59c61-365">Code examples</span></span>

<span data-ttu-id="59c61-366">Os seguintes exemplos de código mostram como criar um suplemento simples Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-366">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="59c61-367">Para baixar o exemplo de código em que esses exemplos se baseiam, consulte [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="59c61-367">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="59c61-368">Se você usar uma caixa de diálogo com o evento ao enviar, certifique-se de fechar a caixa de diálogo antes de concluir o evento.</span><span class="sxs-lookup"><span data-stu-id="59c61-368">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="59c61-369">Manifesto, versão de substituição e evento</span><span class="sxs-lookup"><span data-stu-id="59c61-369">Manifest, version override, and event</span></span>

<span data-ttu-id="59c61-370">Um exemplo de código [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) inclui dois manifestos:</span><span class="sxs-lookup"><span data-stu-id="59c61-370">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="59c61-371">`Contoso Message Body Checker.xml` &ndash; Mostra como verificar se o corpo de uma mensagem apresenta palavras restritas ou informações confidenciais ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-371">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="59c61-372">`Contoso Subject and CC Checker.xml` &ndash; Mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui uma linha de assunto ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-372">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="59c61-373">No arquivo de manifesto `Contoso Message Body Checker.xml`, inclua o arquivo de função e o nome da função que deve ser chamada no evento `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="59c61-373">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="59c61-374">A operação é executada de maneira síncrona.</span><span class="sxs-lookup"><span data-stu-id="59c61-374">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="59c61-375">Se você estiver usando o Visual Studio 2019 para desenvolver seu complemento ao enviar, poderá receber um aviso de validação como o seguinte: "Este é um xsi:type ' http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events 'inválido". Para resolver isso, você precisará de uma versão mais recente do MailAppVersionOverridesV1_1.xsd que tenha sido fornecida como um GitHub em um blog sobre esse [aviso](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span><span class="sxs-lookup"><span data-stu-id="59c61-375">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="59c61-376">Para o arquivo de manifesto `Contoso Subject and CC Checker.xml`, o exemplo a seguir mostra o arquivo de função e o nome da função para chamar o evento de envio de mensagem.</span><span class="sxs-lookup"><span data-stu-id="59c61-376">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="59c61-377">A API Ao enviar requer `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="59c61-377">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="59c61-378">Veja a seguir como adicionar o nó `VersionOverrides` em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="59c61-378">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="59c61-379">Para obter mais informações, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="59c61-379">For more information, see the following:</span></span>
> - [<span data-ttu-id="59c61-380">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="59c61-380">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="59c61-381">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="59c61-381">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="59c61-382">Objetos `Event` e `item`, e os métodos `body.getAsync` e `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="59c61-382">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="59c61-383">Para acessar o item de mensagem ou de reunião selecionado no momento (neste exemplo, a mensagem redigida recentemente), use o namespace `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="59c61-383">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="59c61-384">O evento `ItemSend` é passado automaticamente pelo recurso Ao enviar para a função especificada no manifesto&mdash;neste exemplo, a função `validateBody`.</span><span class="sxs-lookup"><span data-stu-id="59c61-384">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="59c61-385">A função `validateBody` obtém o corpo atual no formato especificado (HTML) e passa o objeto de evento `ItemSend` que o código deseja para acessar o método de retorno.</span><span class="sxs-lookup"><span data-stu-id="59c61-385">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="59c61-386">Além do método `getAsync`, o objeto `Body` também fornece um método `setAsync` que você pode usar para substituir o corpo pelo texto especificado.</span><span class="sxs-lookup"><span data-stu-id="59c61-386">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="59c61-387">Para saber mais, confira [Objeto do Evento](/javascript/api/office/office.addincommands.event) e [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="59c61-387">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="59c61-388">Objeto `NotificationMessages` e método `event.completed`</span><span class="sxs-lookup"><span data-stu-id="59c61-388">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="59c61-389">A função `checkBodyOnlyOnSendCallBack` usa uma expressão regular para determinar se o corpo da mensagem contém palavras bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="59c61-389">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="59c61-390">Se ela encontrar uma correspondência com uma matriz de palavras restritas, bloqueará os emails de serem enviados e notificará o remetente pela barra de informações.</span><span class="sxs-lookup"><span data-stu-id="59c61-390">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="59c61-391">Para fazer isso, ele usa a propriedade `notificationMessages` do objeto `Item` para retornar um objeto `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="59c61-391">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="59c61-392">Ele, em seguida, adiciona uma notificação ao item chamando o método `addAsync`, como mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="59c61-392">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="59c61-393">A seguir estão os parâmetros do `addAsync` método.</span><span class="sxs-lookup"><span data-stu-id="59c61-393">The following are the parameters for the `addAsync` method.</span></span>

- <span data-ttu-id="59c61-394">`NoSend` &ndash; uma cadeia de caractere que é uma chave especificada pelo desenvolvedor para fazer referência a uma mensagem de notificação.</span><span class="sxs-lookup"><span data-stu-id="59c61-394">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="59c61-395">Você pode usá-la para modificar esta mensagem mais tarde.</span><span class="sxs-lookup"><span data-stu-id="59c61-395">You can use it to modify this message later.</span></span> <span data-ttu-id="59c61-396">A chave não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="59c61-396">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="59c61-397">`type` &ndash; uma das propriedades do parâmetro de objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="59c61-397">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="59c61-398">Representa o tipo de uma mensagem; os tipos correspondem aos valores da enumeração [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="59c61-398">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="59c61-399">Os valores possíveis são indicador de progresso, mensagem informativa ou mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="59c61-399">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="59c61-400">Neste exemplo, `type` é uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="59c61-400">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="59c61-401">`message` &ndash; uma das propriedades do parâmetro de objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="59c61-401">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="59c61-402">Neste exemplo, `message` é o texto da mensagem de notificação.</span><span class="sxs-lookup"><span data-stu-id="59c61-402">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="59c61-403">Para sinalizar que o suplemento terminou de processar o evento `ItemSend` disparado pela operação enviar, chame o método `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="59c61-403">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="59c61-404">A propriedade `allowEvent` é um booleano.</span><span class="sxs-lookup"><span data-stu-id="59c61-404">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="59c61-405">Se for definido como `true`, o envio será permitido.</span><span class="sxs-lookup"><span data-stu-id="59c61-405">If set to `true`, send is allowed.</span></span> <span data-ttu-id="59c61-406">Se definido como `false`, a mensagem de email será impedida de ser enviada.</span><span class="sxs-lookup"><span data-stu-id="59c61-406">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="59c61-407">Para saber mais, confira [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="59c61-407">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="59c61-408">Métodos `replaceAsync`, `removeAsync` e `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="59c61-408">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="59c61-409">Além do método `addAsync`, o objeto `NotificationMessages` também inclui os métodos `replaceAsync`, `removeAsync` e `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="59c61-409">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="59c61-410">Esses métodos não são usados neste exemplo de código.</span><span class="sxs-lookup"><span data-stu-id="59c61-410">These methods are not used in this code sample.</span></span>  <span data-ttu-id="59c61-411">Para saber mais, veja [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="59c61-411">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="59c61-412">Código do Assunto e do verificador de CC</span><span class="sxs-lookup"><span data-stu-id="59c61-412">Subject and CC checker code</span></span>

<span data-ttu-id="59c61-413">O exemplo de código a seguir mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui um assunto ao enviar.</span><span class="sxs-lookup"><span data-stu-id="59c61-413">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="59c61-414">Este exemplo usa o recurso Ao enviar para permitir ou proibir o envio de um email.</span><span class="sxs-lookup"><span data-stu-id="59c61-414">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="59c61-p155">Para saber mais sobre como adicionar um destinatário à linha CC e verificar se a mensagem de e-mail inclui uma linha de assunto ao enviar e para ver as APIs que você pode usar, consulte o [exemplo Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). O código é bem comentado.</span><span class="sxs-lookup"><span data-stu-id="59c61-p155">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="59c61-417">Confira também</span><span class="sxs-lookup"><span data-stu-id="59c61-417">See also</span></span>

- [<span data-ttu-id="59c61-418">Visão geral da arquitetura e dos recursos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="59c61-418">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="59c61-419">Suplemento do Outlook para demonstração de comando de suplemento</span><span class="sxs-lookup"><span data-stu-id="59c61-419">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)