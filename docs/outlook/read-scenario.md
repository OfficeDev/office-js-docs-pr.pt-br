---
title: Criar suplementos do Outlook para formulários de leitura
description: Suplementos de leitura são suplementos do Outlook que são ativados no painel de leitura ou no inspetor de leitura do Outlook.
ms.date: 04/12/2018
localization_priority: Priority
ms.openlocfilehash: 815234ed046b4c00b91f5acd6cd2c4dcd226dba2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605301"
---
# <a name="create-outlook-add-ins-for-read-forms"></a><span data-ttu-id="057a4-103">Criar suplementos do Outlook para formulários de leitura</span><span class="sxs-lookup"><span data-stu-id="057a4-103">Create Outlook add-ins for read forms</span></span>

<span data-ttu-id="057a4-p101">Suplementos de leitura são suplementos do Outlook que são ativados no painel de leitura ou no inspetor de leitura do Outlook. Ao contrário dos suplementos de redação (suplementos do Outlook que são ativados quando um usuário está criando uma mensagem ou um compromisso), os suplementos de leitura ficam disponíveis quando os usuários:</span><span class="sxs-lookup"><span data-stu-id="057a4-p101">Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:</span></span> 

- <span data-ttu-id="057a4-106">Visualizam um email, uma solicitação de reunião, uma resposta de reunião ou um cancelamento da reunião.</span><span class="sxs-lookup"><span data-stu-id="057a4-106">View an email message, meeting request, meeting response, or meeting cancellation.</span></span>

   > [!NOTE]
   > <span data-ttu-id="057a4-107">O Outlook não ativa suplementos no formulário de leitura para determinados tipos de mensagens, como itens que são anexos de outra mensagem, itens na pasta de rascunhos do Outlook ou itens que estão criptografados ou protegidos de outras maneiras.</span><span class="sxs-lookup"><span data-stu-id="057a4-107">Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.</span></span>
    
- <span data-ttu-id="057a4-108">Exibem um item de reunião em que o usuário é um participante.</span><span class="sxs-lookup"><span data-stu-id="057a4-108">View a meeting item in which the user is an attendee.</span></span>
    
- <span data-ttu-id="057a4-109">Exibem um item de reunião em que o usuário é o organizador (somente versão RTM do Outlook 2013 e do Exchange 2013).</span><span class="sxs-lookup"><span data-stu-id="057a4-109">View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="057a4-p102">Desde a versão Office 2013 SP1, se o usuário estiver exibindo um item de reunião que o usuário tenha organizado, apenas suplementos redigidos poderão realizar a ativação e estar disponíveis. Os suplementos de leitura não estão mais disponíveis nesse cenário.</span><span class="sxs-lookup"><span data-stu-id="057a4-p102">Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.</span></span>


<span data-ttu-id="057a4-p103">Em cada um desses cenários de leitura, o Outlook ativa suplementos quando suas condições de ativação são atendidas e os usuários podem escolher e abrir suplementos ativados na barra de suplemento no Painel de Leitura ou inspetor de leitura. A figura a seguir mostra o suplemento **Bing Mapas** ativado e aberto quando o usuário está lendo uma mensagem que contém um endereço geográfico.</span><span class="sxs-lookup"><span data-stu-id="057a4-p103">In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. The following figure shows the **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.</span></span>


<span data-ttu-id="057a4-114">**Painel do suplemento mostrando o suplemento Bing Mapas funcionando, no caso de uma mensagem selecionada do Outlook que contém um endereço**</span><span class="sxs-lookup"><span data-stu-id="057a4-114">**The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**</span></span>

![Bing Map mail app in Outlook](../images/bing-maps-add-in.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a><span data-ttu-id="057a4-116">Tipos de suplementos disponíveis no modo de leitura</span><span class="sxs-lookup"><span data-stu-id="057a4-116">Types of add-ins available in read mode</span></span>

<span data-ttu-id="057a4-117">Suplementos de leitura podem ser uma combinação dos tipos a seguir.</span><span class="sxs-lookup"><span data-stu-id="057a4-117">Read add-ins can be any combination of the following types.</span></span>

- [<span data-ttu-id="057a4-118">Comandos de suplemento para o Outlook</span><span class="sxs-lookup"><span data-stu-id="057a4-118">Add-in commands for Outlook</span></span>](add-in-commands-for-outlook.md)   
- [<span data-ttu-id="057a4-119">Suplementos contextuais do Outlook</span><span class="sxs-lookup"><span data-stu-id="057a4-119">Contextual Outlook add-ins</span></span>](contextual-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a><span data-ttu-id="057a4-120">Recursos de API disponíveis para suplementos de leitura</span><span class="sxs-lookup"><span data-stu-id="057a4-120">API features available to read add-ins</span></span>

- <span data-ttu-id="057a4-121">Para ativar suplementos em formulários de leitura: confira a Tabela 1 em [Especificar regras de ativação em um manifesto](activation-rules.md#specify-activation-rules-in-a-manifest).</span><span class="sxs-lookup"><span data-stu-id="057a4-121">For activating add-ins in read forms: see Table 1 in [Specify activation rules in a manifest](activation-rules.md#specify-activation-rules-in-a-manifest).</span></span>    
- [<span data-ttu-id="057a4-122">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="057a4-122">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [<span data-ttu-id="057a4-123">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="057a4-123">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="057a4-124">Extrair cadeias de caracteres de entidade de um item do Outlook</span><span class="sxs-lookup"><span data-stu-id="057a4-124">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)   
- [<span data-ttu-id="057a4-125">Obter anexos de um item do Outlook a partir do servidor</span><span class="sxs-lookup"><span data-stu-id="057a4-125">Get attachments of an Outlook item from the server</span></span>](get-attachments-of-an-outlook-item.md)
    

## <a name="see-also"></a><span data-ttu-id="057a4-126">Confira também</span><span class="sxs-lookup"><span data-stu-id="057a4-126">See also</span></span>

- [<span data-ttu-id="057a4-127">Escreva seu primeiro suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="057a4-127">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
