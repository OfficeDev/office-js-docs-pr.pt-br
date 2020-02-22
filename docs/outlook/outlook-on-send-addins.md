---
title: Recurso Ao enviar para suplementos do Outlook
description: Fornece uma maneira de manipular um item ou impedir que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar.
ms.date: 11/07/2019
localization_priority: Normal
ms.openlocfilehash: 4d6b74e01a8b9a8d2c5006ea3bfedb1b613bad86
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165714"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="ddebb-103">Recurso Ao enviar para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="ddebb-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="ddebb-p101">O recurso Ao enviar para suplementos do Outlook fornece uma maneira de manipular uma mensagem ou item de reunião, ou impede que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar. Por exemplo, você pode usar o recurso Ao enviar para:</span><span class="sxs-lookup"><span data-stu-id="ddebb-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="ddebb-106">Impedir que um usuário envie informações confidenciais ou deixe a linha de assunto em branco.</span><span class="sxs-lookup"><span data-stu-id="ddebb-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="ddebb-107">Adicionar um destinatário específico à linha CC em mensagens ou à linha destinatários opcionais em reuniões.</span><span class="sxs-lookup"><span data-stu-id="ddebb-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

> [!NOTE]
> <span data-ttu-id="ddebb-108">No momento, o recurso ao enviar é compatível com o Outlook na Web, no Exchange Online (Office 365), Exchange 2016 local (Atualização Cumulativa 6 ou posterior) e Exchange 2019 local (Atualização Cumulativa 1 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="ddebb-108">The on-send feature is currently supported for Outlook on the web in Exchange Online (Office 365), Exchange 2016 on-premises (Cumulative Update 6 or later), and Exchange 2019 on-premises (Cumulative Update 1 or later).</span></span> <span data-ttu-id="ddebb-109">Esse recurso também está disponível nas versões mais recentes do Outlook, no Windows e no Mac, conectado ao Exchange Online (Office 365).</span><span class="sxs-lookup"><span data-stu-id="ddebb-109">This feature is also available in the latest Outlook builds on Windows and Mac, connected to Exchange Online (Office 365).</span></span> <span data-ttu-id="ddebb-110">O recurso foi introduzido no conjunto de requisitos 1.8.</span><span class="sxs-lookup"><span data-stu-id="ddebb-110">The feature was introduced in requirement set 1.8.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ddebb-111">Os suplementos que usam o recurso ao enviar não são permitidos na [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="ddebb-111">Add-ins that use the on send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="ddebb-112">O recurso ao enviar é acionado pelo tipo de evento `ItemSend` e é sem interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="ddebb-112">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="ddebb-113">Para obter informações sobre limitações relacionadas ao recurso Ao enviar, consulte as [Limitações](#limitations) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="ddebb-113">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="ddebb-114">Como o recurso Ao enviar funciona?</span><span class="sxs-lookup"><span data-stu-id="ddebb-114">How does the on-send feature work?</span></span>

<span data-ttu-id="ddebb-115">Você pode usar o recurso Ao enviar para criar um suplemento do Outlook que integre o evento síncrono `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-115">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="ddebb-116">Este evento detecta que o usuário está pressionando o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) e pode ser usado para impedir que um item seja enviado se houver falha na validação.</span><span class="sxs-lookup"><span data-stu-id="ddebb-116">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="ddebb-117">Por exemplo, quando um usuário dispara um evento de envio de mensagem, um suplemento do Outlook que usa o recurso Ao enviar pode:</span><span class="sxs-lookup"><span data-stu-id="ddebb-117">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="ddebb-118">Ler e validar o conteúdo da mensagem de email</span><span class="sxs-lookup"><span data-stu-id="ddebb-118">Read and validate the email message contents</span></span>
- <span data-ttu-id="ddebb-119">Verificar se a mensagem inclui uma linha de assunto</span><span class="sxs-lookup"><span data-stu-id="ddebb-119">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="ddebb-120">Definir um destinatário predeterminado</span><span class="sxs-lookup"><span data-stu-id="ddebb-120">Set a predetermined recipient</span></span>

<span data-ttu-id="ddebb-121">A validação é feita no lado do cliente no Outlook, no acionamento do evento de envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-121">Validation is done on the client side in Outlook, when the send event is triggered.</span></span> <span data-ttu-id="ddebb-122">Se houver falha na validação, o envio do item será impedido e uma mensagem de erro será exibida na barra de informações que solicita a ação do usuário.</span><span class="sxs-lookup"><span data-stu-id="ddebb-122">If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>  

<span data-ttu-id="ddebb-123">A captura de tela a seguir mostra uma barra de informações que notifica que o remetente adicione um assunto.</span><span class="sxs-lookup"><span data-stu-id="ddebb-123">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Captura de tela mostrando uma mensagem de erro solicitando que o usuário insira uma linha de assunto ausente](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="ddebb-125">A captura de tela a seguir mostra uma barra de informações que notifica que o remetente de que foram encontradas palavras bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="ddebb-125">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Captura de tela mostrando uma mensagem de erro informando ao usuário que foram encontradas palavras bloqueadas](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="ddebb-127">Limitações</span><span class="sxs-lookup"><span data-stu-id="ddebb-127">Limitations</span></span>

<span data-ttu-id="ddebb-128">Atualmente, o recurso Ao enviar tem as seguintes limitações.</span><span class="sxs-lookup"><span data-stu-id="ddebb-128">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="ddebb-129">**AppSource** &ndash; Você não pode publicar suplementos do Outlook que usem o recurso Ao enviar no [AppSource](https://appsource.microsoft.com), pois eles falharão na validação do AppSource.</span><span class="sxs-lookup"><span data-stu-id="ddebb-129">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="ddebb-130">Os suplementos que usam o recurso Ao enviar devem ser implantados pelos administradores.</span><span class="sxs-lookup"><span data-stu-id="ddebb-130">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="ddebb-131">**Manifesto** &ndash; Somente um evento `ItemSend` tem suporte por suplemento.</span><span class="sxs-lookup"><span data-stu-id="ddebb-131">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="ddebb-132">Se você tiver dois ou mais eventos `ItemSend` em um manifesto, haverá falha na validação.</span><span class="sxs-lookup"><span data-stu-id="ddebb-132">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="ddebb-p107">**Desempenho**&ndash; Várias idas e voltas ao servidor Web que hospeda o suplemento podem afetar o desempenho do suplemento. Considere os efeitos sobre o desempenho quando você cria suplemento que exigem várias mensagens ou operações baseadas em reuniões.</span><span class="sxs-lookup"><span data-stu-id="ddebb-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="ddebb-135">**Enviar mais tarde** (somente Mac) &ndash; Se houver suplementos Ao enviar, o recurso **Enviar mais tarde** ficará indisponível.</span><span class="sxs-lookup"><span data-stu-id="ddebb-135">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="ddebb-136">Limitações de tipo/modo de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ddebb-136">Mailbox type/mode limitations</span></span>

<span data-ttu-id="ddebb-137">A funcionalidade Ao enviar é compatível apenas com caixas de correio de usuários no Outlook na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="ddebb-137">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="ddebb-138">Atualmente, a funcionalidade não tem suporte para os seguintes tipos e modos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-138">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="ddebb-139">Caixas de correio compartilhadas</span><span class="sxs-lookup"><span data-stu-id="ddebb-139">Shared mailboxes</span></span>
- <span data-ttu-id="ddebb-140">Caixas de correio de grupo</span><span class="sxs-lookup"><span data-stu-id="ddebb-140">Group mailboxes</span></span>
- <span data-ttu-id="ddebb-141">Modo offline</span><span class="sxs-lookup"><span data-stu-id="ddebb-141">Offline mode</span></span>

<span data-ttu-id="ddebb-142">O Outlook não permitirá o envio se o recurso Ao enviar estiver habilitado para esses cenários de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-142">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="ddebb-143">No entanto, se um usuário responder a um email em uma caixa de correio de grupo, o suplemento Ao enviar não será executado e a mensagem será enviada.</span><span class="sxs-lookup"><span data-stu-id="ddebb-143">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="ddebb-144">Vários suplementos Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-144">Multiple on-send add-ins</span></span>

<span data-ttu-id="ddebb-145">Se vários suplementos Ao enviar estiverem instalados, os suplementos serão executados na ordem em que são recebidos das APIs `getAppManifestCall` ou `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-145">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="ddebb-146">Se o primeiro suplemento permitir envio, o segundo suplemento poderá alterar algo que faria o primeiro bloquear o envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-146">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="ddebb-147">No entanto, o primeiro suplemento não será executado novamente se todos os suplementos instalados tiverem permissão de envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-147">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="ddebb-148">Por exemplo, o Suplemento1 e o Suplemento2 usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-148">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="ddebb-149">O Suplemento1 é instalado primeiro e o Suplemento2 é instalado depois.</span><span class="sxs-lookup"><span data-stu-id="ddebb-149">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="ddebb-150">O Suplemento1 verifica se a palavra Fabrikam aparece na mensagem como uma condição para o suplemento permitir o envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-150">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="ddebb-151">No entanto, o Suplemento2 remove as ocorrências da palavra Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="ddebb-151">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="ddebb-152">A mensagem será enviada com todas as instâncias de Fabrikam removidas (devido à ordem de instalação do Suplemento1 e do Suplemento2).</span><span class="sxs-lookup"><span data-stu-id="ddebb-152">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="ddebb-153">Implantar suplementos do Outlook que usam Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-153">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="ddebb-154">Recomendamos que os administradores implantem suplementos do Outlook que usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-154">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="ddebb-155">Os administradores precisam garantir que o suplemento Ao enviar:</span><span class="sxs-lookup"><span data-stu-id="ddebb-155">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="ddebb-156">Esteja sempre presente a qualquer momento que um item de redigir é aberto (para email: novo, responder ou encaminhar).</span><span class="sxs-lookup"><span data-stu-id="ddebb-156">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="ddebb-157">Não pode ser fechado ou desabilitado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="ddebb-157">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="ddebb-158">Instalar suplementos do Outlook que usam Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-158">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="ddebb-159">O recurso Ao enviar no Outlook exige que os suplementos sejam configurados para os tipos de eventos de envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-159">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="ddebb-160">Selecione a plataforma que você deseja configurar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-160">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="ddebb-161">Navegador da Web – Outlook clássico</span><span class="sxs-lookup"><span data-stu-id="ddebb-161">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="ddebb-162">Os suplementos para Outlook na Web (clássicos) que usam o recurso Ao enviar serão executados para usuários aos quais é atribuída uma política de caixa de correio do Outlook na Web que tenha o sinalizador *OnSendAddinsEnabled* definido como **true**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-162">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="ddebb-163">Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="ddebb-163">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="ddebb-164">Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="ddebb-164">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="ddebb-165">Habilitar o recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-165">Enable the on-send feature</span></span>

<span data-ttu-id="ddebb-166">Por padrão, a funcionalidade Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="ddebb-166">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="ddebb-167">Os administradores podem habilitar a funcionalidade Ao enviar executando os cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="ddebb-167">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="ddebb-168">Para habilitar suplementos Ao enviar para todos os usuários:</span><span class="sxs-lookup"><span data-stu-id="ddebb-168">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="ddebb-169">Criar uma nova política de caixa de correio do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-169">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="ddebb-170">Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-170">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="ddebb-171">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-171">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ddebb-172">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-172">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ddebb-173">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="ddebb-173">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="ddebb-174">Habilitar o recurso Ao enviar para um grupo de usuários</span><span class="sxs-lookup"><span data-stu-id="ddebb-174">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="ddebb-175">Para habilitar o recurso Ao enviar para um grupo específico de usuários, as etapas são as seguintes.</span><span class="sxs-lookup"><span data-stu-id="ddebb-175">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="ddebb-176">Neste exemplo, um administrador deseja habilitar apenas o recurso de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finance (em que os usuários do Finance estão no Departamento Financeiro).</span><span class="sxs-lookup"><span data-stu-id="ddebb-176">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="ddebb-177">Crie uma nova política de caixa de correio do Outlook na Web para o grupo.</span><span class="sxs-lookup"><span data-stu-id="ddebb-177">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="ddebb-178">Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações).</span><span class="sxs-lookup"><span data-stu-id="ddebb-178">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="ddebb-179">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-179">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ddebb-180">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-180">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ddebb-181">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="ddebb-181">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="ddebb-182">Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS).</span><span class="sxs-lookup"><span data-stu-id="ddebb-182">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="ddebb-183">Quando a política entrar em vigor, o recurso Ao enviar será habilitado para o grupo.</span><span class="sxs-lookup"><span data-stu-id="ddebb-183">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="ddebb-184">Desabilitar o recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-184">Disable the on-send feature</span></span>

<span data-ttu-id="ddebb-185">Para desabilitar o recurso Ao enviar de um usuário ou atribuir uma política de caixa de correio do Outlook na Web que não tenha o sinalizador habilitado, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="ddebb-185">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="ddebb-186">Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="ddebb-186">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="ddebb-187">Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="ddebb-187">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="ddebb-188">Para desabilitar o recurso Ao enviar para todos os usuários que tenham uma política específica de caixa de correio do Outlook na Web atribuída, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="ddebb-188">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="ddebb-189">Navegador da Web – Outlook moderno</span><span class="sxs-lookup"><span data-stu-id="ddebb-189">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="ddebb-190">Os suplementos para Outlook na Web (modernos) que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-190">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="ddebb-191">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, então a política de caixa de correio deve ter o sinalizador *OnSendAddinsEnabled* definido como **true**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-191">However, if users are required to run the add-in to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="ddebb-192">Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="ddebb-192">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="ddebb-193">Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="ddebb-193">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-policy"></a><span data-ttu-id="ddebb-194">Habilitar a política Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-194">Enable the on-send policy</span></span>

<span data-ttu-id="ddebb-195">Por padrão, a política Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="ddebb-195">By default, on-send policy is disabled.</span></span> <span data-ttu-id="ddebb-196">Os administradores podem habilitar a funcionalidade Ao enviar executando os cmdlets do PowerShell do Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="ddebb-196">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="ddebb-197">Para habilitar suplementos Ao enviar para todos os usuários:</span><span class="sxs-lookup"><span data-stu-id="ddebb-197">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="ddebb-198">Criar uma nova política de caixa de correio do Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-198">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="ddebb-199">Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-199">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="ddebb-200">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-200">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ddebb-201">Habilitar o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-201">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ddebb-202">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="ddebb-202">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-policy-for-a-group-of-users"></a><span data-ttu-id="ddebb-203">Habilitar a política Ao enviar para um grupo de usuários</span><span class="sxs-lookup"><span data-stu-id="ddebb-203">Enable the on-send policy for a group of users</span></span>

<span data-ttu-id="ddebb-204">Para habilitar a política Ao enviar para um grupo específico de usuários, as etapas são as seguintes.</span><span class="sxs-lookup"><span data-stu-id="ddebb-204">To enable the on-send policy for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="ddebb-205">Neste exemplo, um administrador apenas deseja habilitar uma política de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finanças (em que os usuários do Finanças estão no Departamento Financeiro).</span><span class="sxs-lookup"><span data-stu-id="ddebb-205">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="ddebb-206">Crie uma nova política de caixa de correio do Outlook na Web para o grupo.</span><span class="sxs-lookup"><span data-stu-id="ddebb-206">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="ddebb-207">Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações).</span><span class="sxs-lookup"><span data-stu-id="ddebb-207">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="ddebb-208">As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-208">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="ddebb-209">Habilitar a política Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-209">Enable the on-send policy.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="ddebb-210">Atribua a política aos usuários.</span><span class="sxs-lookup"><span data-stu-id="ddebb-210">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="ddebb-211">Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS).</span><span class="sxs-lookup"><span data-stu-id="ddebb-211">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="ddebb-212">Quando a política entrar em vigor, o recurso Ao enviar será aplicado ao grupo.</span><span class="sxs-lookup"><span data-stu-id="ddebb-212">When the policy takes effect, the on-send feature will be enforced for the group.</span></span>

#### <a name="disable-the-on-send-policy"></a><span data-ttu-id="ddebb-213">Desabilitar a política Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-213">Disable the on-send policy</span></span>

<span data-ttu-id="ddebb-214">Para desabilitar a política Ao enviar para um usuário ou atribuir uma política de caixa de correio do Outlook na Web que não tenha o sinalizador habilitado, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="ddebb-214">To disable the on-send policy for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="ddebb-215">Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="ddebb-215">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="ddebb-216">Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="ddebb-216">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="ddebb-217">Para desabilitar a política Ao enviar para todos os usuários que tenham uma política específica de caixa de correio do Outlook na Web atribuída, execute os seguintes cmdlets.</span><span class="sxs-lookup"><span data-stu-id="ddebb-217">To disable the on-send policy for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="ddebb-218">Windows</span><span class="sxs-lookup"><span data-stu-id="ddebb-218">Windows</span></span>](#tab/windows)

<span data-ttu-id="ddebb-219">Os suplementos para Outlook no Windows que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-219">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="ddebb-220">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a política de grupo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** deve estar definida como **Habilitada** em cada máquina aplicável.</span><span class="sxs-lookup"><span data-stu-id="ddebb-220">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="ddebb-221">Para definir as políticas de caixa de correio, os administradores podem baixar a [ferramenta Modelos administrativos](https://www.microsoft.com/download/details.aspx?id=49030) e acessar os modelos administrativos mais recentes, executando o editor de Política de grupo local, **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-221">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="ddebb-222">O que a política faz</span><span class="sxs-lookup"><span data-stu-id="ddebb-222">What the policy does</span></span>

<span data-ttu-id="ddebb-223">Por motivos de conformidade, os administrador podem precisar garantir que os usuários não possam enviar itens de mensagem de reunião até que o último suplemento Ao enviar esteja disponível para execução.</span><span class="sxs-lookup"><span data-stu-id="ddebb-223">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="ddebb-224">Os administradores devem habilitar a política de grupo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** para que todos os suplementos sejam atualizados a partir do Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-224">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="ddebb-225">Status da política</span><span class="sxs-lookup"><span data-stu-id="ddebb-225">Policy status</span></span>|<span data-ttu-id="ddebb-226">Resultado</span><span class="sxs-lookup"><span data-stu-id="ddebb-226">Result</span></span>|
|---|---|
|<span data-ttu-id="ddebb-227">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-227">Disabled</span></span>|<span data-ttu-id="ddebb-228">Envio permitido.</span><span class="sxs-lookup"><span data-stu-id="ddebb-228">Send allowed.</span></span> <span data-ttu-id="ddebb-229">É possível enviar uma mensagem ou item de reunião sem executar o suplemento Ao enviar, mesmo que o suplemento ainda não tenha sido atualizado no Exchange.</span><span class="sxs-lookup"><span data-stu-id="ddebb-229">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="ddebb-230">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-230">Enabled</span></span>|<span data-ttu-id="ddebb-231">É permitido enviar somente quando o suplemento foi atualizado do Exchange; caso contrário, o envio está bloqueado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-231">Send allowed only when the add-in has been updated from Exchange; otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="ddebb-232">Gerenciar a política Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-232">Manage the on-send policy</span></span>

<span data-ttu-id="ddebb-233">Por padrão, a política Ao enviar está desabilitada.</span><span class="sxs-lookup"><span data-stu-id="ddebb-233">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="ddebb-234">Os administradores podem habilitar a política Ao enviar ao certificar-se de que a configuração de política de grupo do usuário **Desabilitar o envio quando as extensões da Web não puderem ser carregadas** esteja definida como **Habilitada**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-234">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="ddebb-235">Para desabilitar a política para um usuário, o administrador deve defini-la como **Desabilitada**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-235">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="ddebb-236">Para gerenciar essa configuração de política, você pode fazer o seguinte.</span><span class="sxs-lookup"><span data-stu-id="ddebb-236">To manage this policy setting, you can do the following.</span></span>

1. <span data-ttu-id="ddebb-237">Baixe a [ferramenta de Modelos Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) mais recente.</span><span class="sxs-lookup"><span data-stu-id="ddebb-237">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="ddebb-238">Abra o editor de Política de Grupo Local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="ddebb-238">Open the Local Group Policy editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="ddebb-239">Navegue até **Configuração do Usuário > Modelos Administrativos > Microsoft Outlook 2016 > Segurança > Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-239">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="ddebb-240">Marque a configuração **Desabilitar o envio quando as extensões da Web não puderem ser carregadas**.</span><span class="sxs-lookup"><span data-stu-id="ddebb-240">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="ddebb-241">Abra o link para configuração Editar política.</span><span class="sxs-lookup"><span data-stu-id="ddebb-241">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="ddebb-242">Na caixa de diálogo **Desabilitar o envio quando as extensões da Web não puderem ser carregadas**, selecione **Habilitado** ou **Desabilitado** conforme apropriado e selecione **OK** ou **Aplique** para colocar a atualização em vigor.</span><span class="sxs-lookup"><span data-stu-id="ddebb-242">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="ddebb-243">Mac</span><span class="sxs-lookup"><span data-stu-id="ddebb-243">Mac</span></span>](#tab/unix)

<span data-ttu-id="ddebb-244">Os suplementos para Outlook no Mac que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-244">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="ddebb-245">No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a configuração de caixa de correio a seguir deverá ser aplicada ao computador de cada usuário.</span><span class="sxs-lookup"><span data-stu-id="ddebb-245">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="ddebb-246">Esta configuração ou chave é compatível com CFPreference. Isso significa que é possível defini-la usando um software de gerenciamento empresarial para Mac, como o Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="ddebb-246">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

|||
|:---|:---|
|<span data-ttu-id="ddebb-247">**Domínio**</span><span class="sxs-lookup"><span data-stu-id="ddebb-247">**Domain**</span></span>|<span data-ttu-id="ddebb-248">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="ddebb-248">com.microsoft.outlook</span></span>|
|<span data-ttu-id="ddebb-249">**Chave**</span><span class="sxs-lookup"><span data-stu-id="ddebb-249">**Key**</span></span>|<span data-ttu-id="ddebb-250">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="ddebb-250">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="ddebb-251">**DataType**</span><span class="sxs-lookup"><span data-stu-id="ddebb-251">**DataType**</span></span>|<span data-ttu-id="ddebb-252">Booliano</span><span class="sxs-lookup"><span data-stu-id="ddebb-252">Boolean</span></span>|
|<span data-ttu-id="ddebb-253">**Valores possíveis**</span><span class="sxs-lookup"><span data-stu-id="ddebb-253">**Possible values**</span></span>|<span data-ttu-id="ddebb-254">falso (padrão)</span><span class="sxs-lookup"><span data-stu-id="ddebb-254">false (default)</span></span><br><span data-ttu-id="ddebb-255">verdadeiro</span><span class="sxs-lookup"><span data-stu-id="ddebb-255">true</span></span>|
|<span data-ttu-id="ddebb-256">**Disponibilidade**</span><span class="sxs-lookup"><span data-stu-id="ddebb-256">**Availability**</span></span>|<span data-ttu-id="ddebb-257">16.27</span><span class="sxs-lookup"><span data-stu-id="ddebb-257">16.27</span></span>|
|<span data-ttu-id="ddebb-258">**Comentários**</span><span class="sxs-lookup"><span data-stu-id="ddebb-258">**Comments**</span></span>|<span data-ttu-id="ddebb-259">Essa chave cria uma política de onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="ddebb-259">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="ddebb-260">O que a configuração faz</span><span class="sxs-lookup"><span data-stu-id="ddebb-260">What the setting does</span></span>

<span data-ttu-id="ddebb-261">Por motivos de conformidade, os administradores podem precisar garantir que os usuários não possam enviar itens de mensagem ou de reunião até que os suplementos estejam disponíveis para execução.</span><span class="sxs-lookup"><span data-stu-id="ddebb-261">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="ddebb-262">Os administradores devem habilitar a chave **OnSendAddinsWaitForLoad** para que todos os suplementos sejam atualizados no Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-262">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="ddebb-263">Estado da chave</span><span class="sxs-lookup"><span data-stu-id="ddebb-263">Key's state</span></span>|<span data-ttu-id="ddebb-264">Resultado</span><span class="sxs-lookup"><span data-stu-id="ddebb-264">Result</span></span>|
|---|---|
|<span data-ttu-id="ddebb-265">falso</span><span class="sxs-lookup"><span data-stu-id="ddebb-265">false</span></span>|<span data-ttu-id="ddebb-266">Envio permitido.</span><span class="sxs-lookup"><span data-stu-id="ddebb-266">Send allowed.</span></span> <span data-ttu-id="ddebb-267">É possível enviar uma mensagem ou item de reunião sem executar o suplemento Ao enviar, mesmo que o suplemento ainda não tenha sido atualizado no Exchange.</span><span class="sxs-lookup"><span data-stu-id="ddebb-267">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="ddebb-268">verdadeiro</span><span class="sxs-lookup"><span data-stu-id="ddebb-268">true</span></span>|<span data-ttu-id="ddebb-269">É permitido enviar somente quando o suplemento foi atualizado do Exchange; caso contrário, o envio estará bloqueado e o botão **Enviar** será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-269">Send allowed only when add-ins have been updated from Exchange; otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="ddebb-270">Cenários do recurso Ao enviar</span><span class="sxs-lookup"><span data-stu-id="ddebb-270">On-send feature scenarios</span></span>

<span data-ttu-id="ddebb-271">Veja a seguir os cenários com suporte e sem suporte para suplementos que usam o recurso Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-271">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="ddebb-272">A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, mas nenhum suplemento está instalado</span><span class="sxs-lookup"><span data-stu-id="ddebb-272">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="ddebb-273">Neste cenário, o usuário poderá enviar itens de mensagem e de reunião sem nenhum suplemento em execução.</span><span class="sxs-lookup"><span data-stu-id="ddebb-273">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="ddebb-274">A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, e os suplementos compatíveis com Ao enviar estão instalados e habilitados</span><span class="sxs-lookup"><span data-stu-id="ddebb-274">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="ddebb-275">Os suplementos serão executados durante o evento de envio, que em seguida permitirão ou impedirão o usuário de enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-275">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="ddebb-276">Delegação de caixa de correio, onde a caixa de correio 1 tem permissões de acesso total à caixa de correio 2</span><span class="sxs-lookup"><span data-stu-id="ddebb-276">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="ddebb-277">Navegador da Web (Outlook clássico)</span><span class="sxs-lookup"><span data-stu-id="ddebb-277">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="ddebb-278">Cenário</span><span class="sxs-lookup"><span data-stu-id="ddebb-278">Scenario</span></span>|<span data-ttu-id="ddebb-279">Recurso Ao enviar da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="ddebb-279">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="ddebb-280">Recurso Ao enviar da caixa de correio 2</span><span class="sxs-lookup"><span data-stu-id="ddebb-280">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="ddebb-281">Sessão Web do Outlook (clássico)</span><span class="sxs-lookup"><span data-stu-id="ddebb-281">Outlook web session (classic)</span></span>|<span data-ttu-id="ddebb-282">Resultado</span><span class="sxs-lookup"><span data-stu-id="ddebb-282">Result</span></span>|<span data-ttu-id="ddebb-283">Com suporte?</span><span class="sxs-lookup"><span data-stu-id="ddebb-283">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="ddebb-284">1</span><span class="sxs-lookup"><span data-stu-id="ddebb-284">1</span></span>|<span data-ttu-id="ddebb-285">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-285">Enabled</span></span>|<span data-ttu-id="ddebb-286">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-286">Enabled</span></span>|<span data-ttu-id="ddebb-287">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="ddebb-287">New session</span></span>|<span data-ttu-id="ddebb-288">A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.</span><span class="sxs-lookup"><span data-stu-id="ddebb-288">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="ddebb-p134">Não há suporte atualmente. Como alternativa, use o cenário 3.</span><span class="sxs-lookup"><span data-stu-id="ddebb-p134">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="ddebb-291">2</span><span class="sxs-lookup"><span data-stu-id="ddebb-291">2</span></span>|<span data-ttu-id="ddebb-292">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-292">Disabled</span></span>|<span data-ttu-id="ddebb-293">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-293">Enabled</span></span>|<span data-ttu-id="ddebb-294">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="ddebb-294">New session</span></span>|<span data-ttu-id="ddebb-295">A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.</span><span class="sxs-lookup"><span data-stu-id="ddebb-295">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="ddebb-p135">Não há suporte atualmente. Como alternativa, use o cenário 3.</span><span class="sxs-lookup"><span data-stu-id="ddebb-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="ddebb-298">3</span><span class="sxs-lookup"><span data-stu-id="ddebb-298">3</span></span>|<span data-ttu-id="ddebb-299">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-299">Enabled</span></span>|<span data-ttu-id="ddebb-300">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-300">Enabled</span></span>|<span data-ttu-id="ddebb-301">Mesma sessão</span><span class="sxs-lookup"><span data-stu-id="ddebb-301">Same session</span></span>|<span data-ttu-id="ddebb-302">Os suplementos Ao enviar atribuídos à caixa de correio 1 são executados ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-302">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="ddebb-303">Com suporte.</span><span class="sxs-lookup"><span data-stu-id="ddebb-303">Supported.</span></span>|
|<span data-ttu-id="ddebb-304">4</span><span class="sxs-lookup"><span data-stu-id="ddebb-304">4</span></span>|<span data-ttu-id="ddebb-305">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-305">Enabled</span></span>|<span data-ttu-id="ddebb-306">Desabilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-306">Disabled</span></span>|<span data-ttu-id="ddebb-307">Nova sessão</span><span class="sxs-lookup"><span data-stu-id="ddebb-307">New session</span></span>|<span data-ttu-id="ddebb-308">Nenhum suplemento Ao envio é executado; item de mensagem ou de reunião é enviado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-308">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="ddebb-309">Com suporte.</span><span class="sxs-lookup"><span data-stu-id="ddebb-309">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="ddebb-310">Navegador da Web (Outlook moderno), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="ddebb-310">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="ddebb-311">Para impor o Ao enviar, os administradores devem garantir que a política tenha sido habilitada nas duas caixas de correio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-311">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="ddebb-312">Para saber como oferecer suporte ao acesso de representante em um suplemento, confira [Habilitar cenários de acesso de representante em um suplemento do Outlook](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="ddebb-312">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="ddebb-313">O grupo 1 é uma caixa de correio do grupo moderna e a caixa de correio 1 do usuário é membro do grupo 1</span><span class="sxs-lookup"><span data-stu-id="ddebb-313">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="ddebb-314">Cenário</span><span class="sxs-lookup"><span data-stu-id="ddebb-314">Scenario</span></span>|<span data-ttu-id="ddebb-315">Política Ao enviar da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="ddebb-315">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="ddebb-316">Suplementos Ao enviar habilitados?</span><span class="sxs-lookup"><span data-stu-id="ddebb-316">On-send add-ins enabled?</span></span>|<span data-ttu-id="ddebb-317">Ação da caixa de correio 1</span><span class="sxs-lookup"><span data-stu-id="ddebb-317">Mailbox 1 action</span></span>|<span data-ttu-id="ddebb-318">Resultado</span><span class="sxs-lookup"><span data-stu-id="ddebb-318">Result</span></span>|<span data-ttu-id="ddebb-319">Com suporte?</span><span class="sxs-lookup"><span data-stu-id="ddebb-319">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="ddebb-320">1</span><span class="sxs-lookup"><span data-stu-id="ddebb-320">1</span></span>|<span data-ttu-id="ddebb-321">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-321">Enabled</span></span>|<span data-ttu-id="ddebb-322">Sim</span><span class="sxs-lookup"><span data-stu-id="ddebb-322">Yes</span></span>|<span data-ttu-id="ddebb-323">A caixa de correio 1 compõe uma nova mensagem ou reunião para o grupo 1.</span><span class="sxs-lookup"><span data-stu-id="ddebb-323">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="ddebb-324">Os suplementos Ao enviar são executados durante o envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-324">On-send add-ins run during send.</span></span>|<span data-ttu-id="ddebb-325">Sim</span><span class="sxs-lookup"><span data-stu-id="ddebb-325">Yes</span></span>|
|<span data-ttu-id="ddebb-326">2</span><span class="sxs-lookup"><span data-stu-id="ddebb-326">2</span></span>|<span data-ttu-id="ddebb-327">Habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-327">Enabled</span></span>|<span data-ttu-id="ddebb-328">Sim</span><span class="sxs-lookup"><span data-stu-id="ddebb-328">Yes</span></span>|<span data-ttu-id="ddebb-329">A caixa de correio 1 compõe uma nova mensagem ou reunião para o grupo 1 dentro da janela de grupo do grupo 1 no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="ddebb-329">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="ddebb-330">Os suplementos Ao enviar não são executados durante o envio.</span><span class="sxs-lookup"><span data-stu-id="ddebb-330">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="ddebb-331">Não há suporte atualmente.</span><span class="sxs-lookup"><span data-stu-id="ddebb-331">Not currently supported.</span></span> <span data-ttu-id="ddebb-332">Como alternativa, use o cenário 1.</span><span class="sxs-lookup"><span data-stu-id="ddebb-332">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="ddebb-333">Caixa de correio do usuário com recurso/política de suplemento Ao enviar habilitado, os suplementos com suporte à funcionalidade Ao enviar estão instalados e habilitados e o modo offline está habilitado</span><span class="sxs-lookup"><span data-stu-id="ddebb-333">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="ddebb-334">Os suplementos Ao enviar serão executados de acordo com o estado online do usuário, o back-end do suplemento e o Exchange.</span><span class="sxs-lookup"><span data-stu-id="ddebb-334">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="ddebb-335">Estado do usuário</span><span class="sxs-lookup"><span data-stu-id="ddebb-335">User's state</span></span>

<span data-ttu-id="ddebb-336">Os suplementos Ao enviar serão executados durante o envio se o usuário estiver online.</span><span class="sxs-lookup"><span data-stu-id="ddebb-336">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="ddebb-337">Se o usuário estiver offline, os suplementos Ao enviar não serão executados e o item de mensagem ou de reunião não será enviado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-337">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="ddebb-338">Estado do back-end do suplemento</span><span class="sxs-lookup"><span data-stu-id="ddebb-338">Add-in backend's state</span></span>

<span data-ttu-id="ddebb-339">Um suplemento Ao enviar será executado se o seu back-end estiver online e acessível.</span><span class="sxs-lookup"><span data-stu-id="ddebb-339">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="ddebb-340">Se o back-end estiver offline, ao enviar será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-340">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="ddebb-341">Estado do Exchange</span><span class="sxs-lookup"><span data-stu-id="ddebb-341">Exchange's state</span></span>

<span data-ttu-id="ddebb-342">Os suplementos Ao enviar serão executados durante o envio se o servidor do Exchange estiver online e acessível.</span><span class="sxs-lookup"><span data-stu-id="ddebb-342">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="ddebb-343">Se o suplemento Ao enviar não puder alcançar o Exchange e a política ou cmdlet aplicável estiverem ativados, o envio será desabilitado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-343">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="ddebb-344">No Mac, em qualquer estado offline, o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) está desabilitado e uma notificação é exibida informando que sua organização não permite envio quando o usuário está offline.</span><span class="sxs-lookup"><span data-stu-id="ddebb-344">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>


## <a name="code-examples"></a><span data-ttu-id="ddebb-345">Exemplos de código</span><span class="sxs-lookup"><span data-stu-id="ddebb-345">Code examples</span></span>

<span data-ttu-id="ddebb-346">Os seguintes exemplos de código mostram como criar um suplemento simples Ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-346">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="ddebb-347">Para baixar o exemplo de código em que esses exemplos se baseiam, consulte [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="ddebb-347">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="ddebb-348">Manifesto, versão de substituição e evento</span><span class="sxs-lookup"><span data-stu-id="ddebb-348">Manifest, version override, and event</span></span>

<span data-ttu-id="ddebb-349">Um exemplo de código [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) inclui dois manifestos:</span><span class="sxs-lookup"><span data-stu-id="ddebb-349">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="ddebb-350">`Contoso Message Body Checker.xml` &ndash; Mostra como verificar se o corpo de uma mensagem apresenta palavras restritas ou informações confidenciais ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-350">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="ddebb-351">`Contoso Subject and CC Checker.xml` &ndash; Mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui uma linha de assunto ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-351">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="ddebb-352">No arquivo de manifesto `Contoso Message Body Checker.xml`, inclua o arquivo de função e o nome da função que deve ser chamada no evento `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-352">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="ddebb-353">A operação é executada de maneira síncrona.</span><span class="sxs-lookup"><span data-stu-id="ddebb-353">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="ddebb-354">Se você estiver usando o Visual Studio 2019 para desenvolver seu suplemento ao enviar, você pode receber um aviso de validação como o seguinte: "Este é um xsi: tipo inválido 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." Para contornar isso, você precisará de uma versão mais recente do MailAppVersionOverridesV1_1. xsd que tenha sido fornecida como um serviço de GitHub em um [blog sobre esse aviso](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span><span class="sxs-lookup"><span data-stu-id="ddebb-354">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="ddebb-355">Para o arquivo de manifesto `Contoso Subject and CC Checker.xml`, o exemplo a seguir mostra o arquivo de função e o nome da função para chamar o evento de envio de mensagem.</span><span class="sxs-lookup"><span data-stu-id="ddebb-355">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="ddebb-356">A API Ao enviar requer `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-356">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="ddebb-357">Veja a seguir como adicionar o nó `VersionOverrides` em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="ddebb-357">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On Send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="ddebb-358">Para obter mais informações, confira o seguinte:</span><span class="sxs-lookup"><span data-stu-id="ddebb-358">For more information, see the following:</span></span>
> - [<span data-ttu-id="ddebb-359">Manifestos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="ddebb-359">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="ddebb-360">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="ddebb-360">VersionOverrides</span></span>](../develop/create-addin-commands.md#step-3-add-versionoverrides-element)
> - [<span data-ttu-id="ddebb-361">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ddebb-361">Office Add-ins XML manifest</span></span>](../overview/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="ddebb-362">Objetos `Event` e `item`, e os métodos `body.getAsync` e `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="ddebb-362">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="ddebb-363">Para acessar o item de mensagem ou de reunião selecionado no momento (neste exemplo, a mensagem redigida recentemente), use o namespace `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-363">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="ddebb-364">O evento `ItemSend` é passado automaticamente pelo recurso Ao enviar para a função especificada no manifesto&mdash;neste exemplo, a função `validateBody`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-364">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="ddebb-365">A função `validateBody` obtém o corpo atual no formato especificado (HTML) e passa o objeto de evento `ItemSend` que o código deseja para acessar o método de retorno.</span><span class="sxs-lookup"><span data-stu-id="ddebb-365">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="ddebb-366">Além do método `getAsync`, o objeto `Body` também fornece um método `setAsync` que você pode usar para substituir o corpo pelo texto especificado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-366">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="ddebb-367">Para saber mais, confira [Objeto do Evento](/javascript/api/office/office.addincommands.event) e [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="ddebb-367">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="ddebb-368">Objeto `NotificationMessages` e método `event.completed`</span><span class="sxs-lookup"><span data-stu-id="ddebb-368">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="ddebb-369">A função `checkBodyOnlyOnSendCallBack` usa uma expressão regular para determinar se o corpo da mensagem contém palavras bloqueadas.</span><span class="sxs-lookup"><span data-stu-id="ddebb-369">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="ddebb-370">Se ela encontrar uma correspondência com uma matriz de palavras restritas, bloqueará os emails de serem enviados e notificará o remetente pela barra de informações.</span><span class="sxs-lookup"><span data-stu-id="ddebb-370">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="ddebb-371">Para fazer isso, ele usa a propriedade `notificationMessages` do objeto `Item` para retornar um objeto `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-371">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="ddebb-372">Ele, em seguida, adiciona uma notificação ao item chamando o método `addAsync`, como mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="ddebb-372">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="ddebb-373">Veja a seguir os parâmetros para o método `addAsync`:</span><span class="sxs-lookup"><span data-stu-id="ddebb-373">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="ddebb-374">`NoSend` &ndash; uma cadeia de caractere que é uma chave especificada pelo desenvolvedor para fazer referência a uma mensagem de notificação.</span><span class="sxs-lookup"><span data-stu-id="ddebb-374">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="ddebb-375">Você pode usá-la para modificar esta mensagem mais tarde.</span><span class="sxs-lookup"><span data-stu-id="ddebb-375">You can use it to modify this message later.</span></span> <span data-ttu-id="ddebb-376">A chave não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ddebb-376">The key can’t be longer than 32 characters.</span></span>
- <span data-ttu-id="ddebb-377">`type` &ndash; uma das propriedades do parâmetro de objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="ddebb-377">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="ddebb-378">Representa o tipo de uma mensagem; os tipos correspondem aos valores da enumeração [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="ddebb-378">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="ddebb-379">Os valores possíveis são indicador de progresso, mensagem informativa ou mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="ddebb-379">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="ddebb-380">Neste exemplo, `type` é uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="ddebb-380">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="ddebb-381">`message` &ndash; uma das propriedades do parâmetro de objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="ddebb-381">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="ddebb-382">Neste exemplo, `message` é o texto da mensagem de notificação.</span><span class="sxs-lookup"><span data-stu-id="ddebb-382">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="ddebb-383">Para sinalizar que o suplemento terminou de processar o evento `ItemSend` disparado pela operação enviar, chame o método `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-383">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="ddebb-384">A propriedade `allowEvent` é um booleano.</span><span class="sxs-lookup"><span data-stu-id="ddebb-384">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="ddebb-385">Se for definido como `true`, o envio será permitido.</span><span class="sxs-lookup"><span data-stu-id="ddebb-385">If set to `true`, send is allowed.</span></span> <span data-ttu-id="ddebb-386">Se definido como `false`, a mensagem de email será impedida de ser enviada.</span><span class="sxs-lookup"><span data-stu-id="ddebb-386">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="ddebb-387">Para saber mais, confira [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="ddebb-387">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="ddebb-388">Métodos `replaceAsync`, `removeAsync` e `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="ddebb-388">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="ddebb-389">Além do método `addAsync`, o objeto `NotificationMessages` também inclui os métodos `replaceAsync`, `removeAsync` e `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="ddebb-389">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="ddebb-390">Esses métodos não são usados neste exemplo de código.</span><span class="sxs-lookup"><span data-stu-id="ddebb-390">These methods are not used in this code sample.</span></span>  <span data-ttu-id="ddebb-391">Para saber mais, veja [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="ddebb-391">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="ddebb-392">Código do Assunto e do verificador de CC</span><span class="sxs-lookup"><span data-stu-id="ddebb-392">Subject and CC checker code</span></span>

<span data-ttu-id="ddebb-393">O exemplo de código a seguir mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui um assunto ao enviar.</span><span class="sxs-lookup"><span data-stu-id="ddebb-393">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="ddebb-394">Este exemplo usa o recurso Ao enviar para permitir ou proibir o envio de um email.</span><span class="sxs-lookup"><span data-stu-id="ddebb-394">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="ddebb-p153">Para saber mais sobre como adicionar um destinatário à linha CC e verificar se a mensagem de e-mail inclui uma linha de assunto ao enviar e para ver as APIs que você pode usar, consulte o [exemplo Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). O código é bem comentado.</span><span class="sxs-lookup"><span data-stu-id="ddebb-p153">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="ddebb-397">Confira também</span><span class="sxs-lookup"><span data-stu-id="ddebb-397">See also</span></span>

- [<span data-ttu-id="ddebb-398">Visão geral da arquitetura e dos recursos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="ddebb-398">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="ddebb-399">Suplemento do Outlook para demonstração de comando de suplemento</span><span class="sxs-lookup"><span data-stu-id="ddebb-399">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)