---
title: Crie suplementos do Outlook para formulários de redação
description: Saiba mais sobre os cenários e recursos dos suplementos do Outlook nos formulários de redação.
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: ea9a9bb245e8912111cad154a0cc66b0288d6eaa
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165835"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a><span data-ttu-id="b4830-103">Criar suplementos do Outlook para formulários de redação</span><span class="sxs-lookup"><span data-stu-id="b4830-103">Create Outlook add-ins for compose forms</span></span>

<span data-ttu-id="b4830-104">A partir da versão 1.1 do esquema de manifestos de suplementos do Office e da versão 1.1 do Office.js, você pode criar suplementos de redação, que são suplementos do Outlook ativados nos formulários de redação.</span><span class="sxs-lookup"><span data-stu-id="b4830-104">Starting with version 1.1 of the schema for Office Add-ins manifests and v1.1 of Office.js, you can create compose add-ins, which are Outlook add-ins activated in compose forms.</span></span> <span data-ttu-id="b4830-105">Ao contrário dos suplementos de leitura (suplementos do Outlook que são ativados no modo de leitura quando um usuário está exibindo uma mensagem ou um compromisso), os suplementos de redação estão disponíveis nos seguintes cenários de usuário:</span><span class="sxs-lookup"><span data-stu-id="b4830-105">In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios:</span></span>

- <span data-ttu-id="b4830-106">Redação de nova mensagem, solicitação de reunião ou compromisso em um formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="b4830-106">Composing a new message, meeting request, or appointment in a compose form.</span></span>

- <span data-ttu-id="b4830-107">Exibição ou edição de compromisso existente, ou item de reunião no qual o usuário seja o organizador.</span><span class="sxs-lookup"><span data-stu-id="b4830-107">Viewing or editing an existing appointment, or meeting item in which the user is the organizer.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="b4830-108">Se o usuário estiver na versão RTM do Outlook 2013 e do Exchange 2013 e estiver exibindo um item de reunião organizado pelo usuário, ele poderá encontrar suplementos de leitura disponíveis.
</span><span class="sxs-lookup"><span data-stu-id="b4830-108">If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available.</span></span> <span data-ttu-id="b4830-109">Desde a versão do Office 2013 SP1, há uma alteração que, no mesmo cenário, somente suplementos redigidos podem ativar e estar disponíveis.</span><span class="sxs-lookup"><span data-stu-id="b4830-109">Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.</span></span>

- <span data-ttu-id="b4830-110">Redação de uma mensagem de resposta embutida ou resposta a uma mensagem em um formulário de redação separado.</span><span class="sxs-lookup"><span data-stu-id="b4830-110">Composing an inline response message or replying to a message in a separate compose form.</span></span>

- <span data-ttu-id="b4830-111">Edição de uma resposta (**Aceitar**, **Provisório** ou **Recusar**) a uma solicitação de reunião ou a um item de reunião.</span><span class="sxs-lookup"><span data-stu-id="b4830-111">Editing a response (**Accept**, **Tentative**, or **Decline**) to a meeting request or meeting item.</span></span>

- <span data-ttu-id="b4830-112">Proposição de novo horário para um item de reunião.</span><span class="sxs-lookup"><span data-stu-id="b4830-112">Proposing a new time for a meeting item.</span></span>

- <span data-ttu-id="b4830-113">Encaminhamento ou resposta a uma solicitação de reunião ou a um item de reunião.</span><span class="sxs-lookup"><span data-stu-id="b4830-113">Forwarding or replying to a meeting request or meeting item.</span></span>

<span data-ttu-id="b4830-114">Em cada um desses cenários de redação, são mostrados os botões de comando do suplemento definidos por este.</span><span class="sxs-lookup"><span data-stu-id="b4830-114">In each of these compose scenarios, any add-in command buttons defined by the add-in are shown.</span></span> <span data-ttu-id="b4830-115">Para suplementos mais antigos que não implementam comandos de suplemento, os usuários podem escolher **Suplementos do Office** na faixa de opções para abrir o painel de seleção de suplementos, escolher e iniciar um suplemento de redação.</span><span class="sxs-lookup"><span data-stu-id="b4830-115">For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in.</span></span> <span data-ttu-id="b4830-116">A figura a seguir mostra comandos de suplemento em um formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="b4830-116">The following figure shows add-in commands in a compose form.</span></span>

![Mostra um fomulário de criação do Outlook com comandos de suplementos.](../images/compose-form-commands.png)

<span data-ttu-id="b4830-118">A figura a seguir mostra o painel de seleção do suplemento composto por dois suplementos de redação que não implementam comandos de suplemento, ativado quando o usuário está compondo uma resposta embutida no Outlook.</span><span class="sxs-lookup"><span data-stu-id="b4830-118">The following figure shows the add-in selection pane consisting of two compose add-ins that do not implement add-in commands, activated when the user is composing an inline reply in Outlook.</span></span>

![Aplicativo de email Modelos ativado para item redigido](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a><span data-ttu-id="b4830-120">Tipos de suplementos disponíveis no modo de redação</span><span class="sxs-lookup"><span data-stu-id="b4830-120">Types of add-ins available in compose mode</span></span>

<span data-ttu-id="b4830-121">Os suplementos de redação são implementados como [Comandos de suplemento para Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b4830-121">Compose add-ins are implemented as [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span> <span data-ttu-id="b4830-122">Para ativar suplementos para redação de emails ou respostas de reunião, os suplementos devem incluir um [elemento de ponto de extensão MessageComposeCommandSurface](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) no manifesto.</span><span class="sxs-lookup"><span data-stu-id="b4830-122">To activate add-ins for composing email or meeting responses, add-ins include a [MessageComposeCommandSurface extension point element](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) in the manifest.</span></span> <span data-ttu-id="b4830-123">Para ativar suplementos para redação ou edição de compromissos ou reuniões em que o usuário é o organizador, os suplementos devem incluir um [elemento de ponto de extensão AppointmentOrganizerCommandSurface](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span><span class="sxs-lookup"><span data-stu-id="b4830-123">To activate add-ins for composing or editing appointments or meetings where the user is the organizer, add-ins include a [AppointmentOrganizerCommandSurface extension point element](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span></span>

> [!NOTE]
> <span data-ttu-id="b4830-124">Os suplementos desenvolvidos para servidores ou clientes sem suporte para comandos de suplemento usam [regras de ativação](activation-rules.md) em um elemento [Rule](../reference/manifest/rule.md) contido no elemento [OfficeApp](../reference/manifest/officeapp.md).</span><span class="sxs-lookup"><span data-stu-id="b4830-124">Add-ins developed for servers or clients that do not support add-in commands use [activation rules](activation-rules.md) in a [Rule](../reference/manifest/rule.md) element contained in the [OfficeApp](../reference/manifest/officeapp.md) element.</span></span> <span data-ttu-id="b4830-125">Os novos suplementos devem usar comandos de suplemento, exceto quando o suplemento for desenvolvido para servidores e clientes mais antigos.</span><span class="sxs-lookup"><span data-stu-id="b4830-125">Unless the add-in is being specifically developed for older clients and servers, new add-ins should use add-in commands.</span></span>

## <a name="api-features-available-to-compose-add-ins"></a><span data-ttu-id="b4830-126">Recursos de API disponíveis para suplementos de redação</span><span class="sxs-lookup"><span data-stu-id="b4830-126">API features available to compose add-ins</span></span>

- [<span data-ttu-id="b4830-127">Adicionar e remover anexos de um item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-127">Add and remove attachments to an item in a compose form in Outlook</span></span>](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [<span data-ttu-id="b4830-128">Obter e definir dados de item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-128">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="b4830-129">Obter, configurar ou adicionar destinatários ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-129">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)
- [<span data-ttu-id="b4830-130">Obter ou definir o assunto ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-130">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="b4830-131">Inserir dados no corpo ao criar um compromisso ou uma mensagem no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-131">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="b4830-132">Obter ou definir o local ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-132">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="b4830-133">Obter ou definir a hora ao criar um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="b4830-133">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a><span data-ttu-id="b4830-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="b4830-134">See also</span></span>

- [<span data-ttu-id="b4830-135">Introdução aos suplementos do Outlook para Office 365</span><span class="sxs-lookup"><span data-stu-id="b4830-135">Get Started with Outlook add-ins for Office 365</span></span>](../quickstarts/outlook-quickstart.md)